import { App, PluginSettingTab, SettingDefinitionItem } from "obsidian";
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

type Key = keyof ContactLinksSettings;

/** Declarative settings (Obsidian 1.13+), so every option shows up in Obsidian's settings search. */
export class ContactLinksSettingTab extends PluginSettingTab {
	constructor(
		app: App,
		private plugin: ContactLinksPlugin,
	) {
		super(app, plugin);
	}

	// Settings live inside the plugin's data (next to the contacts), so read and write them here
	// instead of letting the default persist `plugin.settings` alone over data.json.
	getControlValue(key: string): unknown {
		return this.plugin.settings[key as Key];
	}

	async setControlValue(key: string, value: unknown): Promise<void> {
		const s = this.plugin.settings as unknown as Record<string, unknown>;
		if (key === "peopleFolder") value = (typeof value === "string" ? value : "").trim().replace(/^\/+|\/+$/g, "");
		if (key === "triggerChar") value = (typeof value === "string" ? value : "").trim().slice(0, 2);
		s[key] = value;
		await this.plugin.saveSettings();
	}

	getSettingDefinitions(): SettingDefinitionItem<Key>[] {
		return [this.microsoft(), this.fileImport(), this.notes(), this.suggestions()];
	}

	private microsoft(): SettingDefinitionItem<Key> {
		const { plugin } = this;
		const { auth } = plugin;
		const synced = () => plugin.data.contacts.filter((c) => c.source === "m365").length;
		const last = () => (plugin.data.lastSync ? new Date(plugin.data.lastSync).toLocaleString() : "never");
		return {
			type: "group",
			heading: "Microsoft 365",
			items: [
				{
					name: "Microsoft 365 account",
					aliases: ["sign in", "outlook", "sync"],
					render: (setting) => {
						if (!auth.available) {
							setting.setName("Needs a newer Obsidian").setDesc("Microsoft sign-in uses Obsidian's secret storage.");
							return;
						}
						if (!auth.signedIn) {
							setting
								.setName("Not signed in")
								.setDesc("Sign in to sync your Outlook contacts and the people you work with. Read-only.")
								.addButton((b) => b.setButtonText("Sign in").setCta().onClick(() => plugin.signIn()));
							return;
						}
						const error = plugin.lastSyncError ? ` Last error: ${plugin.lastSyncError}` : "";
						setting
							.setName(`Signed in as ${auth.account || "Microsoft account"}`)
							.setDesc(`${synced()} people synced. Last sync: ${last()}.${error}`)
							.addButton((b) =>
								b
									.setButtonText("Sync now")
									.setCta()
									.onClick(async () => {
										b.setDisabled(true).setButtonText("Syncing…");
										await plugin.sync(true);
										this.update();
									}),
							)
							.addButton((b) =>
								b.setButtonText("Sign out").onClick(() => {
									plugin.signOut();
									this.update();
								}),
							);
					},
				},
				{
					name: "Saved contacts",
					desc: "Your Outlook contacts.",
					control: { type: "toggle", key: "syncContacts" },
				},
				{
					name: "People you work with",
					desc: "Colleagues and people you email, ranked by Microsoft by how much you work with them. They also rank higher in suggestions.",
					control: { type: "toggle", key: "syncPeople" },
				},
				{
					name: "How many people you work with",
					control: { type: "slider", key: "peopleLimit", min: 100, max: 1000, step: 100 },
				},
				{
					name: "Sync every (hours)",
					desc: "Automatic sync while Obsidian is open. 0 turns it off; Sync now and the command still work.",
					control: { type: "number", key: "syncHours", min: 0, step: 1, placeholder: "12" },
				},
			],
		};
	}

	private fileImport(): SettingDefinitionItem<Key> {
		const { plugin } = this;
		return {
			type: "group",
			heading: "Import from a file",
			items: [
				{
					name: "Contacts file",
					aliases: ["csv", "vcard", "vcf", "import"],
					render: (setting) => {
						const count = plugin.data.contacts.filter((c) => c.source !== "m365").length;
						setting
							.setName("Contacts file")
							.setDesc(`${count} imported from files. Export from Outlook as CSV or vCard (.vcf), then import it. Importing again updates existing contacts.`)
							.addButton((b) => b.setButtonText("Import file…").onClick(() => plugin.importFromFile(() => this.update())))
							.addButton((b) =>
								b.setButtonText("Clear all").onClick(async () => {
									await plugin.clearContacts();
									this.update();
								}),
							);
					},
				},
			],
		};
	}

	private notes(): SettingDefinitionItem<Key> {
		return {
			type: "group",
			heading: "Person notes",
			items: [
				{
					name: "People folder",
					desc: "Where person notes live. Notes in this folder are suggested too, with their aliases.",
					control: { type: "folder", key: "peopleFolder", placeholder: "People" },
				},
				{
					name: "Create a note on first mention",
					desc: "Create a minimal note (aliases, email, company) the first time you link someone. Off: insert a link without creating a note.",
					control: { type: "toggle", key: "createNotes" },
				},
				{
					name: "Show the alias you typed",
					desc: "When you type a nickname or alias, link as [[Full Name|alias]].",
					control: { type: "toggle", key: "aliasAsDisplay" },
				},
			],
		};
	}

	private suggestions(): SettingDefinitionItem<Key> {
		return {
			type: "group",
			heading: "Suggestions",
			items: [
				{
					name: "Inside [[ ]]",
					desc: "Suggest people while typing a link. Other notes still show once nothing matches a person, or after Esc.",
					control: { type: "toggle", key: "inBrackets" },
				},
				{
					name: "Minimum characters inside [[ ]]",
					control: { type: "slider", key: "bracketMinChars", min: 1, max: 4, step: 1 },
				},
				{
					name: "Mention trigger",
					desc: "Type this character before a name to get suggestions outside brackets. Leave empty to turn it off.",
					control: { type: "text", key: "triggerChar", placeholder: "@" },
				},
				{
					name: "Suggest on plain names",
					desc: "Suggest on any capitalised word, with no trigger. Accept with Tab; Enter still makes a new line.",
					control: { type: "toggle", key: "bareNames" },
				},
				{
					name: "Minimum characters for plain names",
					control: { type: "slider", key: "bareMinChars", min: 2, max: 6, step: 1 },
				},
			],
		};
	}
}
