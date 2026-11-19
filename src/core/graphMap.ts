import { contactId, fold } from "./normalize";
import type { Contact } from "./types";

/** Microsoft Graph `/me/contacts` item (the fields we select). */
export interface GraphContact {
	displayName?: string | null;
	givenName?: string | null;
	middleName?: string | null;
	surname?: string | null;
	nickName?: string | null;
	emailAddresses?: { address?: string | null }[] | null;
	companyName?: string | null;
	jobTitle?: string | null;
}

/** Microsoft Graph `/me/people` item: people you work with, ranked by relevance. */
export interface GraphPerson {
	displayName?: string | null;
	givenName?: string | null;
	surname?: string | null;
	scoredEmailAddresses?: { address?: string | null }[] | null;
	companyName?: string | null;
	jobTitle?: string | null;
	personType?: { class?: string | null; subclass?: string | null } | null;
}

const clean = (s: string | null | undefined) => (s ?? "").replace(/\s+/g, " ").trim();

function emailsOf(list: { address?: string | null }[] | null | undefined): string[] {
	const seen = new Set<string>();
	const out: string[] = [];
	for (const e of list ?? []) {
		const a = clean(e.address);
		if (!a.includes("@") || seen.has(a.toLowerCase())) continue;
		seen.add(a.toLowerCase());
		out.push(a);
	}
	return out;
}

/** Outlook display names are sometimes "Last, First" or the email itself; prefer the name parts. */
function pickName(display: string, parts: string[]): string {
	const fromParts = parts.filter(Boolean).join(" ");
	if (fromParts) return fromParts;
	if (display && !display.includes("@")) return display;
	return "";
}

function make(name: string, emails: string[], extra: Partial<Contact>): Contact {
	const c: Contact = { id: contactId(name, emails), name, emails, source: "m365" };
	if (extra.nickname && fold(extra.nickname) !== fold(name)) c.nickname = extra.nickname;
	if (extra.company) c.company = extra.company;
	if (extra.title) c.title = extra.title;
	if (extra.rank !== undefined) c.rank = extra.rank;
	return c;
}

export function fromGraphContact(g: GraphContact): Contact | null {
	const emails = emailsOf(g.emailAddresses);
	const name = pickName(clean(g.displayName), [clean(g.givenName), clean(g.middleName), clean(g.surname)]) || emails[0]?.split("@")[0] || "";
	if (!name) return null;
	return make(name, emails, { nickname: clean(g.nickName), company: clean(g.companyName), title: clean(g.jobTitle) });
}

export function fromGraphPerson(g: GraphPerson, rank: number): Contact | null {
	// Groups, rooms and distribution lists also come back from /me/people.
	if (g.personType?.class && g.personType.class !== "Person") return null;
	const emails = emailsOf(g.scoredEmailAddresses);
	const name = pickName(clean(g.displayName), [clean(g.givenName), clean(g.surname)]);
	if (!name || !emails.length) return null;
	return make(name, emails, { company: clean(g.companyName), title: clean(g.jobTitle), rank });
}

/**
 * Combine saved contacts and relevant people into one list keyed by contact id.
 * A saved contact keeps its own fields (nickname etc.) and takes the relevance rank from People.
 */
export function combineSynced(contacts: Contact[], people: Contact[]): Contact[] {
	const byId = new Map<string, Contact>();
	const byEmail = new Map<string, Contact>();
	const add = (c: Contact) => {
		byId.set(c.id, c);
		for (const e of c.emails) byEmail.set(e.toLowerCase(), c);
	};
	contacts.forEach(add);
	for (const p of people) {
		const hit = p.emails.map((e) => byEmail.get(e.toLowerCase())).find(Boolean) ?? byId.get(p.id);
		if (hit) {
			hit.rank = Math.min(hit.rank ?? Infinity, p.rank ?? Infinity);
			if (!Number.isFinite(hit.rank)) delete hit.rank;
			hit.company ??= p.company;
			hit.title ??= p.title;
		} else add(p);
	}
	return [...byId.values()];
}
