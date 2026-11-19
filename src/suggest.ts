import {
	Editor,
	EditorPosition,
	EditorSuggest,
	EditorSuggestContext,
	EditorSuggestTriggerInfo,
	TFile,
} from "obsidian";
import { safeFileName } from "./core/normalize";
import type { Match, SearchOptions } from "./core/people";
import { detectTrigger, trailingClose, TriggerMode } from "./core/trigger";
import type ContactLinksPlugin from "./main";
import { ensurePersonNote } from "./notes";

const SEARCH: Record<TriggerMode, SearchOptions> = {
	// Inside [[ ]] other notes compete for the same text, so only offer clear name or alias matches.
	link: { initials: false, email: false, minScore: 70 },
	mention: { initials: true, email: true },
	// Plain words are noisy: only strong prefix matches.
	bare: { initials: false, email: false, minScore: 80 },
};

export class PeopleSuggest extends EditorSuggest<Match> {
	private mode: TriggerMode = "mention";
	private results: Match[] = [];
	/** Trigger position the user dismissed with Esc; we stay quiet there until the trigger changes. */
	private dismissed: string | null = null;

	constructor(private plugin: ContactLinksPlugin) {
		super(plugin.app);
		this.limit = 20;
		// Tab accepts the highlighted suggestion (the only way to accept in plain-name mode).
		this.scope.register([], "Tab", (evt) => {
			const chooser = (this as unknown as { suggestions?: { useSelectedItem?: (e: KeyboardEvent) => void } }).suggestions;
			if (!chooser?.useSelectedItem) return true; // internal API missing: let Tab behave normally
			chooser.useSelectedItem(evt);
			return false;
		});
	}

	/** Obsidian's PopoverSuggest calls this on Esc. Remember the spot so we stay quiet there. */
	onEscapeKey(evt: KeyboardEvent): void {
		if (this.context) this.dismissed = this.keyOf(this.context.file, this.context.start);
		const base = (EditorSuggest.prototype as unknown as { onEscapeKey?: (e: KeyboardEvent) => void }).onEscapeKey;
		if (base) base.call(this, evt);
		else this.close();
	}

	onTrigger(cursor: EditorPosition, editor: Editor, file: TFile | null): EditorSuggestTriggerInfo | null {
		if (!file || this.plugin.index.size === 0) return null;
		const line = editor.getLine(cursor.line);
		const t = detectTrigger(line.slice(0, cursor.ch), this.plugin.settings);
		if (!t) {
			this.dismissed = null;
			return null;
		}
		const start = { line: cursor.line, ch: t.start };
		if (this.dismissed === this.keyOf(file, start)) return null;

		const results = this.plugin.index.search(t.query, { ...SEARCH[t.mode], limit: this.limit });
		// Returning null lets Obsidian's own suggester (notes, headings) take over.
		if (!results.length) return null;

		this.mode = t.mode;
		this.results = results;
		this.setInstructions(
			t.mode === "bare"
				? [
						{ command: "Tab", purpose: "link person" },
						{ command: "Enter", purpose: "new line" },
						{ command: "Esc", purpose: "dismiss" },
					]
				: [
						{ command: "↵", purpose: "link person" },
						{ command: "Esc", purpose: t.mode === "link" ? "show notes instead" : "dismiss" },
					],
		);
		return { start, end: cursor, query: t.query };
	}

	getSuggestions(_ctx: EditorSuggestContext): Match[] {
		return this.results;
	}

	renderSuggestion(m: Match, el: HTMLElement): void {
		el.addClass("contact-links-suggestion");
		const top = el.createDiv({ cls: "contact-links-name" });
		top.setText(m.person.name);
		if (m.person.path) top.createSpan({ cls: "contact-links-badge", text: "note" });
		const details = [m.alias && `“${m.alias}”`, m.person.title, m.person.company, m.person.email].filter(Boolean);
		if (details.length) el.createDiv({ cls: "contact-links-details", text: details.join(" · ") });
	}

	selectSuggestion(m: Match, evt: MouseEvent | KeyboardEvent): void {
		const ctx = this.context;
		if (!ctx) return;
		if (this.mode === "bare" && evt instanceof KeyboardEvent && evt.key === "Enter") {
			this.close();
			ctx.editor.exec("newlineAndIndent");
			return;
		}
		this.close();
		void this.insert(m, ctx).catch((err) => {
			console.error("Contact Links: could not insert link", err);
		});
	}

	private async insert(m: Match, ctx: EditorSuggestContext): Promise<void> {
		const { editor, file } = ctx;
		const start = ctx.start;
		const end = { line: ctx.end.line, ch: ctx.end.ch + trailingClose(editor.getLine(ctx.end.line).slice(ctx.end.ch), this.mode) };
		const original = editor.getRange(start, end);
		const { settings } = this.plugin;
		const alias = settings.aliasAsDisplay ? m.alias : undefined;

		let link: string;
		if (settings.createNotes) {
			const note = await ensurePersonNote(this.app, settings.peopleFolder, m.person);
			// A disambiguated file name ("Anna Jensen (Acme)") still reads as the plain name in text.
			const display = alias ?? (note.basename !== m.person.name ? m.person.name : undefined);
			link = this.app.fileManager.generateMarkdownLink(note, file.path, undefined, display);
		} else {
			const target = safeFileName(m.person.name);
			link = alias ? `[[${target}|${alias}]]` : `[[${target}]]`;
		}

		// The note was created asynchronously; only replace if the typed text is still there.
		if (editor.getRange(start, end) !== original) {
			editor.replaceSelection(link);
		} else {
			editor.replaceRange(link, start, end);
			editor.setCursor({ line: start.line, ch: start.ch + link.length });
		}
		this.plugin.recordUse(m.person.key);
	}

	private keyOf(file: TFile, pos: EditorPosition): string {
		return `${file.path}:${pos.line}:${pos.ch}`;
	}
}
