import { contactId, fold } from "./normalize";
import type { Contact } from "./types";

/** Google People API `Person` (the fields we request). */
export interface GooglePerson {
	names?: { displayName?: string | null; givenName?: string | null; middleName?: string | null; familyName?: string | null }[] | null;
	nicknames?: { value?: string | null }[] | null;
	emailAddresses?: { value?: string | null }[] | null;
	organizations?: { name?: string | null; title?: string | null }[] | null;
}

const clean = (s: string | null | undefined) => (s ?? "").replace(/\s+/g, " ").trim();

/**
 * Map a Google contact. Saved contacts without a name fall back to the email's local part, as file
 * imports do. "Other contacts" (people you emailed but never saved) without a real name are dropped,
 * so automatic entries like noreply@ don't flood suggestions.
 */
export function fromGooglePerson(g: GooglePerson, kind: "saved" | "other"): Contact | null {
	const seen = new Set<string>();
	const emails: string[] = [];
	for (const e of g.emailAddresses ?? []) {
		const a = clean(e.value);
		if (!a.includes("@") || seen.has(a.toLowerCase())) continue;
		seen.add(a.toLowerCase());
		emails.push(a);
	}
	const n = g.names?.[0];
	const parts = [clean(n?.givenName), clean(n?.middleName), clean(n?.familyName)].filter(Boolean).join(" ");
	let name = parts || clean(n?.displayName);
	if (name.includes("@")) name = "";
	if (!name && kind === "saved") name = emails[0]?.split("@")[0] ?? "";
	if (!name || (kind === "other" && !emails.length)) return null;

	const c: Contact = { id: contactId(name, emails), name, emails, source: "google" };
	const nickname = clean(g.nicknames?.[0]?.value);
	if (nickname && fold(nickname) !== fold(name)) c.nickname = nickname;
	const org = g.organizations?.[0];
	if (clean(org?.name)) c.company = clean(org?.name);
	if (clean(org?.title)) c.title = clean(org?.title);
	return c;
}
