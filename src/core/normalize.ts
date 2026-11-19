const SPECIAL: Record<string, string> = {
	æ: "ae",
	ø: "o",
	œ: "oe",
	ß: "ss",
	ð: "d",
	þ: "th",
	ł: "l",
	đ: "d",
};

/**
 * Lowercase and fold accents so "soren" finds "Søren" and "aero" finds "Ærø".
 * Letters NFD cannot split (æ, ø, ß, ...) are mapped explicitly; the rest (å, é, ü, ...) lose their marks.
 */
export function fold(s: string): string {
	return s
		.toLowerCase()
		.replace(/[æøœßðþłđ]/g, (c) => SPECIAL[c] ?? c)
		.normalize("NFD")
		.replace(/\p{M}/gu, "")
		.trim();
}

/** Words of a folded string, split on anything that is not a letter or digit. */
export function words(folded: string): string[] {
	return folded.split(/[^\p{L}\p{N}]+/u).filter(Boolean);
}

/** Characters Obsidian cannot use in a note name or that break links. */
export function safeFileName(name: string): string {
	return name
		.replace(/[\\/:*?"<>|#^[\]]/g, " ")
		.replace(/\s+/g, " ")
		.trim()
		.replace(/^\.+/, "");
}

export function contactId(name: string, emails: string[]): string {
	const email = emails[0]?.trim().toLowerCase();
	return email ? email : `name:${fold(name)}`;
}
