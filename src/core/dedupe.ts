import { contactId, fold } from "./normalize";
import type { Contact, Source } from "./types";

/**
 * Merge contacts that are the same person, across files, Microsoft 365 and Google.
 *
 * Two contacts are the same person when they share an email address (case-insensitive, and
 * transitively: A~B and B~C puts A, B and C together). A contact without any email joins the
 * person with the same name, but only when exactly one person has that name, so two different
 * people who happen to share a name are never merged.
 *
 * With `sameName`, contacts with the same full name (ignoring case and accents) are merged too,
 * even when their emails differ: usually one person with a work and a private address.
 *
 * Earlier contacts win: the merged contact keeps the first one's id, name and nickname, so pass
 * the lists in priority order. Emails are combined; company and title come from the first contact
 * that has them; the best Microsoft relevance rank is kept.
 */
export function dedupeContacts(contacts: Contact[], opts: { sameName?: boolean } = {}): Contact[] {
	const parent = contacts.map((_, i) => i);
	const find = (i: number): number => {
		while (parent[i] !== i) {
			parent[i] = parent[parent[i]];
			i = parent[i];
		}
		return i;
	};
	// The root is always the earliest index, so it is the group's primary contact.
	const union = (a: number, b: number) => {
		const ra = find(a);
		const rb = find(b);
		if (ra < rb) parent[rb] = ra;
		else if (rb < ra) parent[ra] = rb;
	};

	const byEmail = new Map<string, number>();
	contacts.forEach((c, i) => {
		for (const e of c.emails) {
			const key = e.trim().toLowerCase();
			const seen = byEmail.get(key);
			if (seen === undefined) byEmail.set(key, i);
			else union(i, seen);
		}
	});

	if (opts.sameName) {
		const byName = new Map<string, number>();
		contacts.forEach((c, i) => {
			const name = fold(c.name);
			if (!name) return;
			const seen = byName.get(name);
			if (seen === undefined) byName.set(name, i);
			else union(i, seen);
		});
	}

	// After the email merges: which groups (with an email) carry each name.
	const emailedByName = new Map<string, Set<number>>();
	const emaillessByName = new Map<string, number>();
	contacts.forEach((c, i) => {
		const name = fold(c.name);
		if (!name) return;
		if (c.emails.length) {
			let roots = emailedByName.get(name);
			if (!roots) emailedByName.set(name, (roots = new Set()));
			roots.add(find(i));
		}
	});
	contacts.forEach((c, i) => {
		const name = fold(c.name);
		if (c.emails.length || !name) return;
		const roots = emailedByName.get(name);
		if (roots?.size === 1) union(i, [...roots][0]);
		else if (!roots) {
			const first = emaillessByName.get(name);
			if (first === undefined) emaillessByName.set(name, i);
			else union(i, first);
		}
	});

	const groups = new Map<number, Contact[]>();
	contacts.forEach((c, i) => {
		const root = find(i);
		const group = groups.get(root);
		if (group) group.push(c);
		else groups.set(root, [c]);
	});
	return [...groups.values()].map((g) => (g.length === 1 ? g[0] : mergeGroup(g)));
}

function mergeGroup(group: Contact[]): Contact {
	const [first] = group;
	const merged: Contact = { id: first.id, name: first.name, emails: [] };
	if (first.source) merged.source = first.source;
	const seenEmails = new Set<string>();
	const ids = new Set<string>([first.id]);
	const sources = new Set<Source>();
	for (const c of group) {
		for (const e of c.emails) {
			if (seenEmails.has(e.toLowerCase())) continue;
			seenEmails.add(e.toLowerCase());
			merged.emails.push(e);
		}
		ids.add(c.id);
		sources.add(c.source ?? "file");
		if (!merged.nickname && c.nickname && fold(c.nickname) !== fold(merged.name)) merged.nickname = c.nickname;
		merged.company ||= c.company;
		merged.title ||= c.title;
		if (c.rank !== undefined) merged.rank = Math.min(merged.rank ?? Infinity, c.rank);
	}
	if (!merged.company) delete merged.company;
	if (!merged.title) delete merged.title;
	ids.delete(first.id);
	if (ids.size) merged.otherIds = [...ids];
	merged.sources = [...sources];
	return merged;
}

/**
 * Usage counts keyed by `contact:<id>`. A merged contact's id can be another copy's id from before
 * (for example when a second account is connected), so it takes the highest count of its copies.
 * New picks are counted for every copy, which keeps them when the merge later changes.
 */
export function mergedUsage(contacts: Contact[], usage: Record<string, number>): Record<string, number> {
	const out = { ...usage };
	for (const c of contacts) {
		for (const id of c.otherIds ?? []) {
			const other = usage[`contact:${id}`] ?? 0;
			if (other > (out[`contact:${c.id}`] ?? 0)) out[`contact:${c.id}`] = other;
		}
	}
	return out;
}

/** Senders that are machines, not people: no-reply and notification addresses. */
const AUTOMATED_LOCAL = /^(no-?reply|do-?not-?reply|donotreply|notifications?|mailer-daemon|postmaster|bounces?)([._+-]|$)/i;
const AUTOMATED_DOMAIN = /(^|\.)(reply\.github\.com|noreply\.github\.com|mail\.support\.microsoft\.com)$/i;

export function isAutomatedEmail(email: string): boolean {
	const at = email.lastIndexOf("@");
	if (at < 0) return false;
	return AUTOMATED_LOCAL.test(email.slice(0, at)) || AUTOMATED_DOMAIN.test(email.slice(at + 1));
}

/**
 * Remove automated addresses from every contact, and drop contacts that had only automated ones
 * (GitHub notification senders that Microsoft lists among people you work with, and the like).
 */
export function dropAutomated(contacts: Contact[]): Contact[] {
	const out: Contact[] = [];
	for (const c of contacts) {
		if (!c.emails.some(isAutomatedEmail)) {
			out.push(c);
			continue;
		}
		const emails = c.emails.filter((e) => !isAutomatedEmail(e));
		// The id came from the first email, which may be the automated one.
		if (emails.length) out.push({ ...c, emails, id: contactId(c.name, emails) });
	}
	return out;
}
