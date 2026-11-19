import { App, PluginSettingTab, SettingDefinitionItem } from "obsidian";
import type ContactLinksPlugin from "./main";

export interface ContactLinksSettings {
	/** Microsoft Entra app registration, as in Outlook Calendar Panel and Microsoft To Do Panel. */
	clientId: string;
	tenant: string;
	syncContacts: boolean;
	syncPeople: boolean;
	peopleLimit: number;
	/** Google Cloud OAuth client (Desktop app). Its client secret is in secret storage, not here. */
	googleClientId: string;
	googleSyncContacts: boolean;
	/** Google's "other contacts": people you emailed but never saved. */
	googleSyncOther: boolean;
	/** Also merge contacts with the same full name when their emails differ. */
	mergeSameName: boolean;
	/** Use an existing note with a contact's full name outside the people folder as their note. */
	findNotesElsewhere: boolean;
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
	googleClientId: "",
	googleSyncContacts: true,
	googleSyncOther: true,
	mergeSameName: false,
	findNotesElsewhere: true,
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
		return [this.microsoft(), this.googleGroup(), this.duplicates(), this.fileImport(), this.notes(), this.suggestions()];
	}

	private microsoft(): SettingDefinitionItem<Key> {
		const { plugin } = this;
		const { auth } = plugin;
		const synced = () => plugin.data.synced.m365.length;
		const last = () => lastSyncText(plugin.data.lastSync.m365);
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
						const error = plugin.lastSyncError.m365 ? ` Last error: ${plugin.lastSyncError.m365}` : "";
						setting
							.setName(`Signed in as ${auth.account || "Microsoft account"}`)
							.setDesc(`${synced()} people synced. Last sync: ${last()}.${error}`)
							.addButton((b) =>
								b
									.setButtonText("Sync now")
									.setCta()
									.onClick(async () => {
										b.setDisabled(true).setButtonText("Syncing…");
										await plugin.syncProvider("m365", true);
										this.update();
									}),
							)
							.addButton((b) =>
								b.setButtonText("Sign out").onClick(async () => {
									await plugin.signOut("m365");
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
					desc: "Automatic sync of every signed-in account while Obsidian is open. 0 turns it off; Sync now and the command still work.",
					control: { type: "number", key: "syncHours", min: 0, step: 1, placeholder: "12" },
				},
			],
		};
	}

	private googleGroup(): SettingDefinitionItem<Key> {
		const { plugin } = this;
		const { google } = plugin;
		return {
			type: "group",
			heading: "Google",
			items: [
				{
					name: "Google account",
					aliases: ["sign in", "gmail", "google contacts", "sync"],
					render: (setting) => {
						if (!google.signedIn) {
							if (!google.canSignIn) {
								setting
									.setName("Not signed in")
									.setDesc("Sign in to Google in Obsidian on desktop (1.13 or later). Contacts synced there reach this device with the vault.");
								return;
							}
							setting
								.setName("Not signed in")
								.setDesc("Sign in to sync your Google contacts. Read-only.")
								.addButton((b) => b.setButtonText("Sign in").setCta().onClick(() => plugin.signInGoogle()));
							return;
						}
						const error = plugin.lastSyncError.google ? ` Last error: ${plugin.lastSyncError.google}` : "";
						setting
							.setName(`Signed in as ${google.account || "Google account"}`)
							.setDesc(`${plugin.data.synced.google.length} people synced. Last sync: ${lastSyncText(plugin.data.lastSync.google)}.${error}`)
							.addButton((b) =>
								b
									.setButtonText("Sync now")
									.setCta()
									.onClick(async () => {
										b.setDisabled(true).setButtonText("Syncing…");
										await plugin.syncProvider("google", true);
										this.update();
									}),
							)
							.addButton((b) =>
								b.setButtonText("Sign out").onClick(async () => {
									await plugin.signOut("google");
									this.update();
								}),
							);
					},
				},
				{
					name: "Google saved contacts",
					desc: "The contacts in your Google account.",
					control: { type: "toggle", key: "googleSyncContacts" },
				},
				{
					name: "Google other contacts",
					desc: "People you have emailed in Gmail but not saved. Only those with a name are used.",
					control: { type: "toggle", key: "googleSyncOther" },
				},
			],
		};
	}

	private duplicates(): SettingDefinitionItem<Key> {
		return {
			type: "group",
			heading: "Duplicates",
			items: [
				{
					name: "Merge people with the same name",
					desc: "Contacts that share an email address are always merged. On: also merge contacts with the same full name when their addresses differ, for example a work and a private address. No-reply and notification addresses are always left out.",
					aliases: ["duplicates", "dedupe", "merge"],
					control: { type: "toggle", key: "mergeSameName" },
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
						const count = plugin.data.contacts.length;
						const merged = plugin.beforeMerge - plugin.merged.length;
						const dupes = merged > 0 ? ` ${plugin.merged.length} people in total from all sources, ${merged} duplicates merged.` : "";
						setting
							.setName("Contacts file")
							.setDesc(`${count} imported from files. Export from Outlook or Google as CSV or vCard (.vcf), then import it. Importing again updates existing contacts.${dupes}`)
							.addButton((b) => b.setButtonText("Import file…").onClick(() => plugin.importFromFile(() => this.update())))
							.addButton((b) =>
								b.setButtonText("Clear imported").onClick(async () => {
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
					name: "Use existing notes outside the folder",
					desc: "When a contact has no note in the people folder but a link to their full name already opens a note elsewhere, use that note instead of creating a second one. Move such notes into the folder from the contacts panel or with the command.",
					aliases: ["existing notes", "elsewhere", "duplicates"],
					control: { type: "toggle", key: "findNotesElsewhere" },
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

function lastSyncText(time: number | undefined): string {
	return time ? new Date(time).toLocaleString() : "never";
}
