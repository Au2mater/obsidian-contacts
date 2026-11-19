import { fold, words } from "./normalize";
import type { Contact, Person, PersonNote } from "./types";

/**
 * Merge imported contacts with person notes in the vault. A note wins on name and aliases.
 *
 * `elsewhere` looks up a note outside the people folder for a contact that has none there (in the
 * plugin: the note a link to their full name would open). It is used only for full names of two or
 * more words, a note whose email belongs to someone else is skipped, and no note is claimed twice.
 */
export function mergePeople(contacts: Contact[], notes: PersonNote[], elsewhere?: (name: string) => PersonNote | null): Person[] {
	const people: Person[] = [];
	const byEmail = new Map<string, Person>();
	const byName = new Map<string, Person>();
	const claimed = new Set(notes.map((n) => n.path));

	for (const n of notes) {
		const p: Person = {
			key: `note:${n.path}`,
			name: n.basename,
			aliases: n.aliases,
			email: n.emails[0],
			company: n.company,
			path: n.path,
		};
		people.push(p);
		for (const e of n.emails) byEmail.set(e.toLowerCase(), p);
		byName.set(fold(n.basename), p);
		for (const a of n.aliases) if (!byName.has(fold(a))) byName.set(fold(a), p);
	}

	for (const c of contacts) {
		const emailHit = c.emails.map((e) => byEmail.get(e.toLowerCase())).find(Boolean);
		// Match by name only when the note has no email of its own that contradicts the contact.
		const nameHit = byName.get(fold(c.name));
		const note = emailHit ?? (nameHit && !nameHit.email ? nameHit : undefined);
		if (note) {
			if (!note.contact) {
				note.contact = c;
				// Keep the contact's key so usage ranking survives the note being created.
				note.key = `contact:${c.id}`;
			}
			note.email ??= c.emails[0];
			note.company ??= c.company;
			note.title ??= c.title;
			if (c.nickname && !note.aliases.some((a) => fold(a) === fold(c.nickname!))) {
				note.aliases = [...note.aliases, c.nickname];
			}
			continue;
		}
		const person: Person = {
			key: `contact:${c.id}`,
			name: c.name,
			aliases: c.nickname ? [c.nickname] : [],
			email: c.emails[0],
			company: c.company,
			title: c.title,
			contact: c,
		};
		const found = elsewhere && words(fold(c.name)).length >= 2 ? elsewhere(c.name) : null;
		const mine = c.emails.map((e) => e.toLowerCase());
		if (found && !claimed.has(found.path) && (!found.emails.length || found.emails.some((e) => mine.includes(e)))) {
			claimed.add(found.path);
			person.path = found.path;
			person.elsewhere = true;
			for (const a of found.aliases) if (!person.aliases.some((x) => fold(x) === fold(a))) person.aliases.push(a);
			person.company ??= found.company;
		}
		people.push(person);
	}
	return people;
}

interface Field {
	text: string;
	folded: string;
	words: string[];
	kind: "name" | "alias" | "email";
}

interface Entry {
	person: Person;
	fields: Field[];
	initials: string;
}

export interface Match {
	person: Person;
	score: number;
	/** The alias the query matched, when it matched an alias rather than the name. */
	alias?: string;
}

export interface SearchOptions {
	limit?: number;
	/** Also match on initials ("sj" → Søren Jensen). */
	initials?: boolean;
	/** Also match the local part of the email address. */
	email?: boolean;
	/** Drop matches scoring below this. */
	minScore?: number;
}

/** In-memory index over people. Built once per change; each search is a linear scan over precomputed strings. */
export class PeopleIndex {
	private entries: Entry[] = [];
	private usage: Record<string, number> = {};

	constructor(people: Person[] = [], usage: Record<string, number> = {}) {
		this.set(people, usage);
	}

	get size(): number {
		return this.entries.length;
	}

	/** Everyone in the index, in no particular order. */
	all(): Person[] {
		return this.entries.map((e) => e.person);
	}

	set(people: Person[], usage: Record<string, number> = this.usage): void {
		this.usage = usage;
		this.entries = people.map((person) => {
			const field = (text: string, kind: Field["kind"]): Field => {
				const folded = fold(text);
				return { text, folded, words: words(folded), kind };
			};
			const fields = [field(person.name, "name"), ...person.aliases.map((a) => field(a, "alias"))];
			if (person.email) fields.push(field(person.email.split("@")[0], "email"));
			const initials = fields[0].words.map((w) => w[0]).join("");
			return { person, fields, initials };
		});
	}

	setUsage(usage: Record<string, number>): void {
		this.usage = usage;
	}

	search(query: string, opts: SearchOptions = {}): Match[] {
		const { limit = 20, initials = true, email = true, minScore = 1 } = opts;
		const q = fold(query);
		if (!q) return [];
		const qWords = words(q);
		const qCompact = qWords.join("");
		const out: Match[] = [];

		for (const e of this.entries) {
			let best = 0;
			let alias: string | undefined;
			for (const f of e.fields) {
				if (f.kind === "email" && !email) continue;
				const s = scoreField(f, q, qWords);
				if (s > best) {
					best = s;
					alias = f.kind === "alias" ? f.text : undefined;
				}
			}
			if (initials && qCompact.length >= 2 && e.initials.startsWith(qCompact)) {
				const s = e.initials === qCompact ? 65 : 55;
				if (s > best) {
					best = s;
					alias = undefined;
				}
			}
			if (best < minScore) continue;
			const used = this.usage[e.person.key] ?? 0;
			const rank = e.person.contact?.rank;
			const relevance = rank === undefined ? 0 : Math.max(0, 8 - Math.floor(rank / 25));
			const score = best + Math.min(used, 10) + relevance + (e.person.path ? 3 : 0);
			out.push({ person: e.person, score, alias });
		}

		out.sort((a, b) => b.score - a.score || a.person.name.localeCompare(b.person.name));
		return out.slice(0, limit);
	}
}

/**
 * 100 whole-string prefix, 80 every query word starts a word in order, 70 every query word starts some word.
 * Aliases score a little below names; email local parts much lower.
 */
function scoreField(f: Field, q: string, qWords: string[]): number {
	let s = 0;
	if (f.folded === q) s = 110;
	else if (f.folded.startsWith(q)) s = 100;
	else if (wordsInOrder(f.words, qWords)) s = 80;
	else if (qWords.every((qw) => f.words.some((w) => w.startsWith(qw)))) s = 70;
	if (!s) return 0;
	if (f.kind === "alias") return s - 5;
	if (f.kind === "email") return s - 40;
	return s;
}

function wordsInOrder(fieldWords: string[], qWords: string[]): boolean {
	let i = 0;
	for (const qw of qWords) {
		while (i < fieldWords.length && !fieldWords[i].startsWith(qw)) i++;
		if (i === fieldWords.length) return false;
		i++;
	}
	return true;
}
