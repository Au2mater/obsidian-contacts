/** Where a contact came from. */
export type Source = "file" | "m365" | "google";

/** A contact as imported from a file or synced from Microsoft 365 or Google. */
export interface Contact {
	/** Stable identity: first email (lowercased), or `name:<folded name>` when there is no email. */
	id: string;
	name: string;
	nickname?: string;
	emails: string[];
	company?: string;
	title?: string;
	/** Where it came from. Missing means imported from a file. */
	source?: Source;
	/** Set on a contact merged from duplicates: every source it was found in. */
	sources?: Source[];
	/** Set on a contact merged from duplicates: the ids of the other copies. */
	otherIds?: string[];
	/** Position in Microsoft's "people you work with" ranking (0 = most relevant). */
	rank?: number;
}

/** One person the suggester can offer: an imported contact, a person note, or both merged. */
export interface Person {
	key: string;
	name: string;
	aliases: string[];
	email?: string;
	company?: string;
	title?: string;
	/** Vault path of the person note, when one exists. */
	path?: string;
	contact?: Contact;
}

/** Facts about one person note, read from the vault's metadata cache. */
export interface PersonNote {
	path: string;
	basename: string;
	aliases: string[];
	emails: string[];
	company?: string;
}
