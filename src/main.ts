import { debounce, Notice, Plugin, TAbstractFile } from "obsidian";
import { replaceSynced } from "./core/graphMap";
import { mergeContacts, parseContacts } from "./core/importers";
import { mergePeople, PeopleIndex } from "./core/people";
import type { Contact } from "./core/types";
import { MicrosoftAuth, NotSignedInError } from "./m365/auth";
import { fetchFromMicrosoft } from "./m365/graph";
import { SignInModal } from "./m365/signin";
import { readPersonNotes } from "./notes";
import { ContactLinksSettings, ContactLinksSettingTab, DEFAULT_SETTINGS } from "./settings";
import { PeopleSuggest } from "./suggest";

interface PluginData {
	settings: ContactLinksSettings;
	contacts: Contact[];
	/** How often each person was picked; a small ranking boost. */
	usage: Record<string, number>;
	/** When contacts were last synced from Microsoft 365 (ms since epoch). */
	lastSync?: number;
}

/** How often to check whether an automatic sync is due. */
const SYNC_CHECK_MS = 10 * 60 * 1000;

export default class ContactLinksPlugin extends Plugin {
	data: PluginData = { settings: { ...DEFAULT_SETTINGS }, contacts: [], usage: {} };
	index = new PeopleIndex();

	/** Same object as `data.settings`. */
	settings: ContactLinksSettings = this.data.settings;
	auth!: MicrosoftAuth;
	lastSyncError = "";
	private syncing: Promise<void> | null = null;
	private settingTab!: ContactLinksSettingTab;

	async onload(): Promise<void> {
		const saved = ((await this.loadData()) ?? {}) as Partial<PluginData>;
		this.data = {
			settings: { ...DEFAULT_SETTINGS, ...saved.settings },
			contacts: saved.contacts ?? [],
			usage: saved.usage ?? {},
			lastSync: saved.lastSync,
		};
		this.settings = this.data.settings;
		this.auth = new MicrosoftAuth(this.app, () => this.settings, () => this.sessionEnded());

		const suggest = new PeopleSuggest(this);
		this.registerEditorSuggest(suggest);
		this.preferOverBuiltInLinkSuggest(suggest);
		this.settingTab = new ContactLinksSettingTab(this.app, this);
		this.addSettingTab(this.settingTab);

		this.addCommand({
			id: "import-contacts",
			name: "Import contacts from file (CSV or vCard)",
			callback: () => this.importFromFile(),
		});
		this.addCommand({
			id: "sync-microsoft",
			name: "Sync contacts from Microsoft 365",
			callback: () => (this.auth.signedIn ? void this.sync(true) : this.signIn()),
		});

		// Person notes come from the metadata cache; rebuild when it is ready and when people notes change.
		this.app.workspace.onLayoutReady(() => {
			this.rebuild();
			// Sync after startup has settled, then check now and then whether a sync is due.
			window.setTimeout(() => void this.sync(false), 15_000);
			this.registerInterval(window.setInterval(() => void this.sync(false), SYNC_CHECK_MS));
		});
		const inFolder = (path: string) => {
			const f = this.settings.peopleFolder;
			return f ? path.startsWith(f + "/") : !path.includes("/");
		};
		const onFile = (file: TAbstractFile, oldPath?: string) => {
			if (inFolder(file.path) || (oldPath && inFolder(oldPath))) this.rebuildSoon();
		};
		this.registerEvent(this.app.metadataCache.on("changed", (file) => onFile(file)));
		this.registerEvent(this.app.vault.on("delete", (file) => onFile(file)));
		this.registerEvent(this.app.vault.on("rename", (file, oldPath) => onFile(file, oldPath)));
	}

	private rebuildSoon = debounce(() => this.rebuild(), 300, true);

	rebuild(): void {
		const notes = readPersonNotes(this.app, this.settings.peopleFolder);
		this.index.set(mergePeople(this.data.contacts, notes), this.data.usage);
	}

	/**
	 * Obsidian asks suggesters in order and the first that triggers wins, so ours must come before
	 * the built-in [[ suggester to work inside brackets. This uses an internal list; if it is not there,
	 * suggestions still work outside brackets.
	 */
	private preferOverBuiltInLinkSuggest(suggest: PeopleSuggest): void {
		const list = (this.app.workspace as unknown as { editorSuggest?: { suggests?: unknown[] } }).editorSuggest?.suggests;
		if (!Array.isArray(list)) return;
		const i = list.indexOf(suggest);
		if (i > 0) {
			list.splice(i, 1);
			list.unshift(suggest);
		}
	}

	// ── Microsoft 365 ───────────────────────────────────────────────────────

	signIn(): void {
		if (!this.auth.available) {
			new Notice("Microsoft sign-in needs Obsidian 1.11.4 or later.");
			return;
		}
		new SignInModal(this.app, this).open();
	}

	signedIn(): void {
		new Notice(`Signed in to Microsoft 365${this.auth.account ? ` as ${this.auth.account}` : ""}. Syncing contacts…`);
		void this.sync(true).then(() => {
			if (this.settingTab.containerEl.isConnected) this.settingTab.display();
		});
	}

	signOut(): void {
		this.auth.signOut();
		new Notice("Signed out of Microsoft 365. Synced contacts are kept until you clear them.");
	}

	private sessionEnded(): void {
		new Notice("Your Microsoft 365 sign-in has expired. Sign in again under Settings → Contact Links.");
	}

	/**
	 * Fetch contacts from Microsoft 365 and replace the previously synced ones.
	 * `force` syncs now; otherwise only when signed in and the interval has passed.
	 */
	sync(force: boolean): Promise<void> {
		if (this.syncing) return this.syncing;
		if (!this.auth.signedIn) return Promise.resolve();
		const hours = this.settings.syncHours;
		const due = !this.data.lastSync || (hours > 0 && Date.now() - this.data.lastSync >= hours * 3_600_000);
		if (!force && !due) return Promise.resolve();
		if (!this.settings.syncContacts && !this.settings.syncPeople) return Promise.resolve();

		this.syncing = (async () => {
			try {
				const synced = await fetchFromMicrosoft(this.auth, {
					contacts: this.settings.syncContacts,
					people: this.settings.syncPeople,
					peopleLimit: this.settings.peopleLimit,
				});
				this.data.contacts = replaceSynced(this.data.contacts, synced);
				this.data.lastSync = Date.now();
				this.lastSyncError = "";
				await this.saveSettings();
				if (force) new Notice(`Contacts: ${synced.length} synced from Microsoft 365.`);
			} catch (err) {
				this.lastSyncError = (err as Error).message;
				console.error("Contact Links: sync failed", err);
				// An expired session already showed its own notice.
				if (force && !(err instanceof NotSignedInError && err.expired)) new Notice(`Contact sync failed: ${this.lastSyncError}`);
			} finally {
				this.syncing = null;
			}
		})();
		return this.syncing;
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.data);
		this.rebuild();
	}

	private saveSoon = debounce(() => void this.saveData(this.data), 2000, true);

	recordUse(key: string): void {
		this.data.usage[key] = (this.data.usage[key] ?? 0) + 1;
		this.index.setUsage(this.data.usage);
		this.saveSoon();
	}

	importFromFile(done?: () => void): void {
		const input = document.createElement("input");
		input.type = "file";
		input.accept = ".csv,.vcf,.vcard,text/csv,text/vcard,text/x-vcard";
		input.onchange = async () => {
			const file = input.files?.[0];
			if (!file) return;
			try {
				const { contacts, skipped } = parseContacts(decode(await file.arrayBuffer()));
				if (!contacts.length) {
					new Notice("No contacts found in that file. Export from Outlook as CSV or vCard (.vcf).");
					return;
				}
				const merged = mergeContacts(this.data.contacts, contacts);
				this.data.contacts = merged.contacts;
				await this.saveSettings();
				const extra = skipped ? `, ${skipped} rows without name or email skipped` : "";
				new Notice(`Contacts: ${merged.added} added, ${merged.updated} updated${extra}.`);
				done?.();
			} catch (err) {
				console.error("Contact Links: import failed", err);
				new Notice("Contact import failed. See the developer console for details.");
			}
		};
		input.click();
	}

	async clearContacts(): Promise<void> {
		this.data.contacts = [];
		await this.saveSettings();
		new Notice("Imported contacts cleared. Person notes are kept.");
	}
}

/** Outlook on Windows may export CSV in the ANSI code page rather than UTF-8. */
function decode(buf: ArrayBuffer): string {
	try {
		return new TextDecoder("utf-8", { fatal: true }).decode(buf);
	} catch {
		return new TextDecoder("windows-1252").decode(buf);
	}
}
