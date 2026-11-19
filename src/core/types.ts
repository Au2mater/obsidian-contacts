/** A contact as imported from Outlook (CSV or vCard). */
export interface Contact {
	/** Stable identity: first email (lowercased), or `name:<folded name>` when there is no email. */
	id: string;
	name: string;
	nickname?: string;
	emails: string[];
	company?: string;
	title?: string;
	/** Where it came from. Missing means imported from a file. */
	source?: "file" | "m365";
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
