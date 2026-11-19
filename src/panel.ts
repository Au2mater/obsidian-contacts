import { App, debounce, ItemView, Keymap, Menu, Modal, Notice, PaneType, setIcon, Setting, WorkspaceLeaf } from "obsidian";
import type { Person } from "./core/types";
import type ContactLinksPlugin from "./main";
import { ensurePersonNote } from "./notes";

export const VIEW_TYPE_CONTACTS = "contact-links-contacts";

const SOURCE_LABELS: Record<string, string> = { m365: "Microsoft", google: "Google", file: "File" };
/** Rendering thousands of rows would only slow the panel down; searching narrows it. */
const MAX_ROWS = 200;

/** A sidebar list of everyone the plugin knows, with search. Click someone to open their note. */
export class ContactsView extends ItemView {
	private query = "";
	private listEl!: HTMLElement;
	private countEl!: HTMLElement;

	constructor(
		leaf: WorkspaceLeaf,
		private plugin: ContactLinksPlugin,
	) {
		super(leaf);
	}

	getViewType(): string {
		return VIEW_TYPE_CONTACTS;
	}

	getDisplayText(): string {
		return "Contacts";
	}

	getIcon(): string {
		return "contact";
	}

	async onOpen(): Promise<void> {
		this.build();
	}

	private build(): void {
		const root = this.contentEl;
		root.empty();
		root.addClass("contact-links-panel");
		// Search and count stay pinned at the top while the list scrolls.
		const header = root.createDiv({ cls: "contact-links-header" });
		const bar = header.createDiv({ cls: "contact-links-bar" });
		const search = bar.createEl("input", { type: "search", cls: "contact-links-search", attr: { placeholder: "Search contacts…" } });
		const sync = bar.createDiv({ cls: "clickable-icon contact-links-sync", attr: { "aria-label": "Sync contacts now", role: "button", tabindex: "0" } });
		setIcon(sync, "refresh-cw");
		const runSync = async () => {
			if (sync.hasClass("is-syncing")) return;
			sync.addClass("is-syncing");
			sync.setAttr("aria-label", "Syncing…");
			try {
				await this.plugin.syncNow();
			} finally {
				sync.removeClass("is-syncing");
				sync.setAttr("aria-label", "Sync contacts now");
			}
		};
		sync.addEventListener("click", () => void runSync());
		sync.addEventListener("keydown", (evt) => {
			if (evt.key === "Enter" || evt.key === " ") {
				evt.preventDefault();
				void runSync();
			}
		});
		search.value = this.query;
		const refresh = debounce(() => this.render(), 100, true);
		search.addEventListener("input", () => {
			this.query = search.value;
			refresh();
		});
		this.countEl = header.createDiv({ cls: "contact-links-count" });
		this.listEl = root.createDiv({ cls: "contact-links-list" });
		this.render();
	}

	/** Called by the plugin whenever the contacts or person notes change. */
	render(): void {
		// Not opened yet (Obsidian opens sidebar views lazily); onOpen renders it then.
		if (!this.listEl) return;
		const { index } = this.plugin;
		const people = this.query.trim()
			? index.search(this.query, { limit: MAX_ROWS + 1 }).map((m) => m.person)
			: index
					.all()
					.slice()
					.sort((a, b) => a.name.localeCompare(b.name));
		const total = index.size;
		this.countEl.setText(
			this.query.trim()
				? `${people.length > MAX_ROWS ? `${MAX_ROWS}+` : people.length} of ${total} people`
				: `${total} people${total > MAX_ROWS ? `, showing the first ${MAX_ROWS}. Search to find others.` : ""}`,
		);
		this.listEl.empty();
		for (const person of people.slice(0, MAX_ROWS)) this.renderRow(person);
		if (!total) {
			this.listEl.createDiv({
				cls: "contact-links-empty",
				text: "No contacts yet. Sign in to Microsoft 365 or Google, or import a file, under Settings → Contact Links.",
			});
		}
	}

	private renderRow(person: Person): void {
		const row = this.listEl.createDiv({ cls: "contact-links-row" });
		const top = row.createDiv({ cls: "contact-links-name" });
		top.createSpan({ text: person.name });
		const sources = person.contact ? (person.contact.sources ?? [person.contact.source ?? "file"]) : [];
		for (const s of sources) top.createSpan({ cls: "contact-links-badge", text: SOURCE_LABELS[s] ?? s });
		if (person.path) top.createSpan({ cls: "contact-links-badge", text: "note" });
		const details = [person.title, person.company, person.email].filter(Boolean).join(" · ");
		if (details) row.createDiv({ cls: "contact-links-details", text: details });
		row.setAttr("aria-label", person.path ? "Open note" : "Create a note…");
		row.addEventListener("click", (evt) => this.activate(person, Keymap.isModEvent(evt)));
		row.addEventListener("contextmenu", (evt) => {
			evt.preventDefault();
			this.showMenu(person, evt);
		});
	}

	/** Open the person's note; someone without a note gets one only after confirming. */
	private activate(person: Person, paneType: PaneType | boolean = false): void {
		if (person.path) {
			void this.openPerson(person, paneType);
			return;
		}
		new CreateNoteModal(this.app, person, this.plugin.settings.peopleFolder, () => void this.openPerson(person, paneType)).open();
	}

	private showMenu(person: Person, evt: MouseEvent): void {
		const menu = new Menu();
		menu.addItem((item) =>
			item
				.setTitle(person.path ? "Open note" : "Create note…")
				.setIcon(person.path ? "file-text" : "file-plus")
				.onClick(() => this.activate(person)),
		);
		if (person.path) {
			menu.addItem((item) =>
				item
					.setTitle("Open in new tab")
					.setIcon("file-plus-2")
					.onClick(() => this.activate(person, "tab")),
			);
		}
		const email = person.email;
		if (email) {
			menu.addSeparator();
			menu.addItem((item) =>
				item
					.setTitle("Copy email address")
					.setIcon("copy")
					.onClick(() => void navigator.clipboard.writeText(email).then(() => new Notice("Email address copied."))),
			);
		}
		menu.showAtMouseEvent(evt);
	}

	// Not `open`: that name is the View lifecycle method Obsidian calls to mount the view.
	private async openPerson(person: Person, paneType: PaneType | boolean): Promise<void> {
		try {
			const file = await ensurePersonNote(this.app, this.plugin.settings.peopleFolder, person);
			// In the main area, the active leaf may be this panel itself; never replace it with the note.
			const inMainArea = this.leaf.getRoot() === this.app.workspace.rootSplit;
			await this.app.workspace.getLeaf(inMainArea && !paneType ? "tab" : paneType).openFile(file);
		} catch (err) {
			console.error("Contact Links: could not open note", err);
			new Notice("Could not open the person note. See the developer console for details.");
		}
	}
}

/** Asks before creating a person note from the panel. Enter confirms, Esc cancels. */
class CreateNoteModal extends Modal {
	constructor(
		app: App,
		private person: Person,
		private folder: string,
		private onConfirm: () => void,
	) {
		super(app);
	}

	onOpen(): void {
		const { person } = this;
		this.setTitle(`Create a note for ${person.name}?`);
		const what = [person.email && "email", person.company && "company", person.title && "role"].filter(Boolean).join(", ");
		this.contentEl.createEl("p", {
			text: `A new note will be created in ${this.folder || "the vault root"}${what ? ` with their ${what}` : ""}.`,
		});
		const confirm = () => {
			this.close();
			this.onConfirm();
		};
		this.scope.register([], "Enter", (evt) => {
			evt.preventDefault();
			confirm();
		});
		new Setting(this.contentEl)
			.addButton((b) => b.setButtonText("Cancel").onClick(() => this.close()))
			.addButton((b) => b.setButtonText("Create note").setCta().onClick(confirm));
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
