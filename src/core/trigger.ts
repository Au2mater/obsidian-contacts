export type TriggerMode = "link" | "mention" | "bare";

export interface TriggerSettings {
	/** Suggest people after `[[`. */
	inBrackets: boolean;
	/** Character that starts a mention outside brackets, e.g. "@". Empty disables it. */
	triggerChar: string;
	/** Suggest on capitalised words without any trigger. */
	bareNames: boolean;
	bareMinChars: number;
	bracketMinChars: number;
}

export interface Trigger {
	mode: TriggerMode;
	query: string;
	/** Column where the replaced text starts (the `[[`, the trigger char, or the word). */
	start: number;
}

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/**
 * Decide from the text before the cursor whether to suggest people.
 * Pure and cheap: it runs on every keypress.
 */
export function detectTrigger(before: string, s: TriggerSettings): Trigger | null {
	if (insideInlineCode(before)) return null;

	// Inside an open wikilink: `[[query` with no closing bracket, alias or heading yet.
	const link = /\[\[([^[\]|#^]*)$/.exec(before);
	if (link) {
		if (before[link.index - 1] === "!") return null; // embeds are rarely people
		if (!s.inBrackets || link[1].length < s.bracketMinChars || /^\s/.test(link[1])) return null;
		return { mode: "link", query: link[1], start: link.index };
	}
	// Already inside some other link or after a `]]` on the same word: leave it alone.
	if (/\[\[[^\]]*$/.test(before)) return null;

	if (s.triggerChar) {
		// `@` at a word start; up to three words so "@søren jen" works. Not inside email addresses.
		const re = new RegExp(
			`(?:^|[\\s([{"'])${escapeRe(s.triggerChar)}([\\p{L}\\p{N}][\\p{L}\\p{M}\\p{N}.'-]*(?: [\\p{L}\\p{M}\\p{N}.'-]*){0,2})$`,
			"u",
		);
		const m = re.exec(before);
		if (m) {
			const start = before.length - m[1].length - s.triggerChar.length;
			return { mode: "mention", query: m[1], start };
		}
	}

	if (s.bareNames) {
		const m = /(?:^|[\s([{"'])(\p{Lu}[\p{L}\p{M}'-]*)$/u.exec(before);
		if (m && m[1].length >= s.bareMinChars) {
			return { mode: "bare", query: m[1], start: before.length - m[1].length };
		}
	}
	return null;
}

function insideInlineCode(before: string): boolean {
	return (before.match(/`/g)?.length ?? 0) % 2 === 1;
}

/** How many characters after the cursor belong to the link being completed (the auto-paired `]]`). */
export function trailingClose(after: string, mode: TriggerMode): number {
	if (mode !== "link") return 0;
	// Through the closing `]]` of this link (also when editing inside an existing one), unless another link starts first.
	const m = /^[^[\]\n]*\]\]/.exec(after);
	return m ? m[0].length : 0;
}
