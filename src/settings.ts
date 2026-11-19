import { App, PluginSettingTab, Setting } from "obsidian";
import type ContactLinksPlugin from "./main";

export interface ContactLinksSettings {
	/** Microsoft Entra app registration, as in Outlook Calendar Panel and Microsoft To Do Panel. */
	clientId: string;
	tenant: string;
	syncContacts: boolean;
	syncPeople: boolean;
	peopleLimit: number;
	/** Hours between automatic syncs; 0 turns automatic sync off. */
	syncHours: number;
	peopleFolder: string;
	/** Create a minimal person note the first time someone is linked. */
	createNotes: boolean;
	inBrackets: boolean;
	bracketMinChars: number;
	triggerChar: string;
	bareNames: boolean;
	bareMinChars: number;
	/** When the query matched an alias, show that alias as the link's display text. */
	aliasAsDisplay: boolean;
}

export const DEFAULT_SETTINGS: ContactLinksSettings = {
	clientId: "",
	tenant: "common",
	syncContacts: true,
	syncPeople: true,
	peopleLimit: 500,
	syncHours: 12,
	peopleFolder: "People",
	createNotes: true,
	inBrackets: true,
	bracketMinChars: 2,
	triggerChar: "@",
	bareNames: false,
	bareMinChars: 3,
	aliasAsDisplay: true,
};

export class ContactLinksSettingTab extends PluginSettingTab {
	constructor(
		app: App,
		private plugin: ContactLinksPlugin,
	) {
		super(app, plugin);
	}

	display(): void {
		const { containerEl } = this;
		const s = this.plugin.settings;
		const save = () => this.plugin.saveSettings();
		containerEl.empty();

		this.microsoft(containerEl);

		new Setting(containerEl).setName("Import from a file").setHeading();
		const fileCount = this.plugin.data.contacts.filter((c) => c.source !== "m365").length;
		new Setting(containerEl)
			.setName("Contacts file")
			.setDesc(`${fileCount} imported from files. Export from Outlook as CSV or vCard (.vcf), then import it. Importing again updates existing contacts.`)
			.addButton((b) => b.setButtonText("Import file…").onClick(() => this.plugin.importFromFile(() => this.display())))
			.addButton((b) =>
				b.setButtonText("Clear all").onClick(async () => {
					await this.plugin.clearContacts();
					this.display();
				}),
			);

		new Setting(containerEl).setName("Person notes").setHeading();
		new Setting(containerEl)
			.setName("People folder")
			.setDesc("Where person notes live. Notes in this folder are suggested too, with their aliases.")
			.addText((t) =>
				t.setPlaceholder("People").setValue(s.peopleFolder).onChange(async (v) => {
					s.peopleFolder = v.trim().replace(/^\/+|\/+$/g, "");
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Create a note on first mention")
			.setDesc("Create a minimal note (aliases, email, company) the first time you link someone. Off: insert a link without creating a note.")
			.addToggle((t) =>
				t.setValue(s.createNotes).onChange(async (v) => {
					s.createNotes = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Show the alias you typed")
			.setDesc("When you type a nickname or alias, link as [[Full Name|alias]].")
			.addToggle((t) =>
				t.setValue(s.aliasAsDisplay).onChange(async (v) => {
					s.aliasAsDisplay = v;
					await save();
				}),
			);

		new Setting(containerEl).setName("Suggestions").setHeading();
		new Setting(containerEl)
			.setName("Inside [[ ]]")
			.setDesc("Suggest people while typing a link. Other notes still show once nothing matches a person, or after Esc.")
			.addToggle((t) =>
				t.setValue(s.inBrackets).onChange(async (v) => {
					s.inBrackets = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Minimum characters inside [[ ]]")
			.addSlider((sl) =>
				sl.setLimits(1, 4, 1).setValue(s.bracketMinChars).setDynamicTooltip().onChange(async (v) => {
					s.bracketMinChars = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Mention trigger")
			.setDesc("Type this character before a name to get suggestions outside brackets. Leave empty to turn it off.")
			.addText((t) =>
				t.setPlaceholder("@").setValue(s.triggerChar).onChange(async (v) => {
					s.triggerChar = v.trim().slice(0, 2);
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Suggest on plain names")
			.setDesc("Suggest on any capitalised word, with no trigger. Accept with Tab; Enter still makes a new line.")
			.addToggle((t) =>
				t.setValue(s.bareNames).onChange(async (v) => {
					s.bareNames = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Minimum characters for plain names")
			.addSlider((sl) =>
				sl.setLimits(2, 6, 1).setValue(s.bareMinChars).setDynamicTooltip().onChange(async (v) => {
					s.bareMinChars = v;
					await save();
				}),
			);
	}

	private microsoft(containerEl: HTMLElement): void {
		const s = this.plugin.settings;
		const { auth } = this.plugin;
		const save = () => this.plugin.saveSettings();
		new Setting(containerEl).setName("Microsoft 365").setHeading();

		if (!auth.available) {
			new Setting(containerEl).setName("Needs a newer Obsidian").setDesc("Microsoft sign-in uses Obsidian's secret storage (Obsidian 1.11.4 or later).");
			return;
		}
		const synced = this.plugin.data.contacts.filter((c) => c.source === "m365").length;
		const last = this.plugin.data.lastSync ? new Date(this.plugin.data.lastSync).toLocaleString() : "never";
		const account = new Setting(containerEl);
		if (auth.signedIn) {
			account
				.setName(`Signed in as ${auth.account || "Microsoft account"}`)
				.setDesc(`${synced} people synced. Last sync: ${last}.${this.plugin.lastSyncError ? ` Last error: ${this.plugin.lastSyncError}` : ""}`)
				.addButton((b) =>
					b
						.setButtonText("Sync now")
						.setCta()
						.onClick(async () => {
							b.setDisabled(true).setButtonText("Syncing…");
							await this.plugin.sync(true);
							this.display();
						}),
				)
				.addButton((b) =>
					b.setButtonText("Sign out").onClick(() => {
						this.plugin.signOut();
						this.display();
					}),
				);
		} else {
			account
				.setName("Not signed in")
				.setDesc("Sign in to sync your Outlook contacts and the people you work with. Read-only.")
				.addButton((b) => b.setButtonText("Sign in").setCta().onClick(() => this.plugin.signIn()));
		}

		new Setting(containerEl)
			.setName("Saved contacts")
			.setDesc("Your Outlook contacts.")
			.addToggle((t) =>
				t.setValue(s.syncContacts).onChange(async (v) => {
					s.syncContacts = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("People you work with")
			.setDesc("Colleagues and people you email, ranked by Microsoft by how much you work with them. They also rank higher in suggestions.")
			.addToggle((t) =>
				t.setValue(s.syncPeople).onChange(async (v) => {
					s.syncPeople = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("How many people you work with")
			.addSlider((sl) =>
				sl.setLimits(100, 1000, 100).setValue(s.peopleLimit).setDynamicTooltip().onChange(async (v) => {
					s.peopleLimit = v;
					await save();
				}),
			);
		new Setting(containerEl)
			.setName("Sync every (hours)")
			.setDesc("Automatic sync while Obsidian is open. 0 turns it off; Sync now and the command still work.")
			.addText((t) =>
				t.setPlaceholder("12").setValue(String(s.syncHours)).onChange(async (v) => {
					const n = Number(v);
					if (Number.isFinite(n) && n >= 0) {
						s.syncHours = n;
						await save();
					}
				}),
			);
	}
}
