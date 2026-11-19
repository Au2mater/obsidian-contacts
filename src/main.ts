import { debounce, Notice, Plugin, TAbstractFile } from "obsidian";
import { dedupeContacts, dropAutomated, mergedUsage } from "./core/dedupe";
import { mergeContacts, parseContacts } from "./core/importers";
import { mergePeople, PeopleIndex } from "./core/people";
import { migrate, PluginData, Provider } from "./core/pluginData";
import type { Contact } from "./core/types";
import { GoogleAuth } from "./google/auth";
import { fetchFromGoogle } from "./google/people";
import { GoogleSignInModal } from "./google/signin";
import { MicrosoftAuth, NotSignedInError } from "./m365/auth";
import { fetchFromMicrosoft } from "./m365/graph";
import { SignInModal } from "./m365/signin";
import { readPersonNotes } from "./notes";
import { ContactLinksSettings, ContactLinksSettingTab, DEFAULT_SETTINGS } from "./settings";
import { ContactsView, VIEW_TYPE_CONTACTS } from "./panel";
import { PeopleSuggest } from "./suggest";

export const PROVIDER_NAMES: Record<Provider, string> = { m365: "Microsoft 365", google: "Google" };

/** How often to check whether an automatic sync is due. */
const SYNC_CHECK_MS = 10 * 60 * 1000;

export default class ContactLinksPlugin extends Plugin {
	data: PluginData = { settings: { ...DEFAULT_SETTINGS }, contacts: [], synced: { m365: [], google: [] }, usage: {}, lastSync: {} };
	index = new PeopleIndex();
	/** All contacts from files and accounts, duplicates merged. */
	merged: Contact[] = [];

	/** Same object as `data.settings`. */
	settings: ContactLinksSettings = this.data.settings;
	auth!: MicrosoftAuth;
	google!: GoogleAuth;
	lastSyncError: Partial<Record<Provider, string>> = {};
	private syncing: Partial<Record<Provider, Promise<void>>> = {};
	/** Bumped on sign-out, so a sync still running then does not write its contacts back. */
	private generation: Record<Provider, number> = { m365: 0, google: 0 };
	/** Contacts left after dropping automated addresses, before merging duplicates. */
	beforeMerge = 0;
	private settingTab!: ContactLinksSettingTab;

	async onload(): Promise<void> {
		this.data = migrate(((await this.loadData()) ?? {}) as Record<string, unknown>, DEFAULT_SETTINGS);
		this.settings = this.data.settings;
		this.auth = new MicrosoftAuth(this.app, () => this.settings, () => this.sessionEnded("m365"));
		this.google = new GoogleAuth(this.app, () => this.settings.googleClientId, () => this.sessionEnded("google"));

		const suggest = new PeopleSuggest(this);
		this.registerEditorSuggest(suggest);
		this.preferOverBuiltInLinkSuggest(suggest);
		this.settingTab = new ContactLinksSettingTab(this.app, this);
		this.addSettingTab(this.settingTab);

		this.registerView(VIEW_TYPE_CONTACTS, (leaf) => new ContactsView(leaf, this));
		this.addRibbonIcon("contact", "Open contacts", () => void this.openContactsPanel());
		this.addCommand({
			id: "open-contacts-panel",
			name: "Open contacts panel",
			callback: () => void this.openContactsPanel(),
		});
		this.addCommand({
			id: "import-contacts",
			name: "Import contacts from file (CSV or vCard)",
			callback: () => this.importFromFile(),
		});
		this.addCommand({
			id: "sync-microsoft",
			name: "Sync contacts now",
			callback: () => void this.syncNow(),
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
		// Priority for the merged name and id: Microsoft 365, then Google, then file imports, so a
		// fresh sync wins over an older file import of the same person (as before 1.1).
		const people = dropAutomated([...this.data.synced.m365, ...this.data.synced.google, ...this.data.contacts]);
		this.beforeMerge = people.length;
		this.merged = dedupeContacts(people, { sameName: this.settings.mergeSameName });
		this.index.set(mergePeople(this.merged, notes), mergedUsage(this.merged, this.data.usage));
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_CONTACTS)) {
			if (leaf.view instanceof ContactsView) leaf.view.render();
		}
	}

	async openContactsPanel(): Promise<void> {
		const { workspace } = this.app;
		let leaf = workspace.getLeavesOfType(VIEW_TYPE_CONTACTS)[0];
		if (!leaf) {
			leaf = workspace.getRightLeaf(false) ?? workspace.getLeaf("tab");
			await leaf.setViewState({ type: VIEW_TYPE_CONTACTS, active: true });
		}
		await workspace.revealLeaf(leaf);
	}

	/** Open Obsidian's settings on this plugin's tab. Uses an internal API; falls back to a notice. */
	openSettings(): void {
		const setting = (this.app as unknown as { setting?: { open(): void; openTabById(id: string): void } }).setting;
		if (!setting) {
			new Notice("Open Settings → Contact Links.");
			return;
		}
		setting.open();
		setting.openTabById(this.manifest.id);
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

	// ── Accounts ────────────────────────────────────────────────────────────

	signIn(): void {
		if (!this.auth.available) {
			new Notice("Microsoft sign-in needs Obsidian 1.13 or later.");
			return;
		}
		new SignInModal(this.app, this).open();
	}

	signInGoogle(): void {
		if (!this.google.canSignIn) {
			new Notice("Google sign-in works in Obsidian on desktop, version 1.13 or later.");
			return;
		}
		new GoogleSignInModal(this.app, this).open();
	}

	signedIn(): void {
		new Notice(`Signed in to Microsoft 365${this.auth.account ? ` as ${this.auth.account}` : ""}. Syncing contacts…`);
		void this.syncProvider("m365", true).then(() => this.refreshSettings());
	}

	googleSignedIn(): void {
		new Notice(`Signed in to Google${this.google.account ? ` as ${this.google.account}` : ""}. Syncing contacts…`);
		void this.syncProvider("google", true).then(() => this.refreshSettings());
	}

	/** Sign out and remove the contacts synced from that account. Person notes are kept. */
	async signOut(provider: Provider): Promise<void> {
		if (provider === "m365") this.auth.signOut();
		else this.google.signOut();
		this.generation[provider]++;
		this.data.synced[provider] = [];
		delete this.data.lastSync[provider];
		this.lastSyncError[provider] = "";
		await this.saveSettings();
		new Notice(`Signed out of ${PROVIDER_NAMES[provider]}. Its synced contacts were removed; person notes are kept.`);
	}

	private sessionEnded(provider: Provider): void {
		new Notice(`Your ${PROVIDER_NAMES[provider]} sign-in has expired. Sign in again under Settings → Contact Links.`);
	}

	isSignedIn(provider: Provider): boolean {
		return provider === "m365" ? this.auth.signedIn : this.google.signedIn;
	}

	/** Sync every signed-in account now, as the command and the panel's button do; explain when there is nothing to sync. */
	async syncNow(): Promise<void> {
		if (!this.auth.signedIn && !this.google.signedIn) {
			new Notice("Sign in to Microsoft 365 or Google first.");
			this.openSettings();
		} else if (!(["m365", "google"] as const).some((p) => this.isSignedIn(p) && this.syncEnabled(p))) {
			new Notice("Syncing is turned off for every signed-in account. Turn on a source under Settings → Contact Links.");
		} else await this.sync(true);
	}

	syncEnabled(provider: Provider): boolean {
		const s = this.settings;
		return provider === "m365" ? s.syncContacts || s.syncPeople : s.googleSyncContacts || s.googleSyncOther;
	}

	/** Sync every signed-in account. One failing does not stop or delay the other. */
	async sync(force: boolean): Promise<void> {
		await Promise.all((["m365", "google"] as const).map((p) => this.syncProvider(p, force)));
	}

	/**
	 * Fetch one account's contacts and replace what was synced from it before.
	 * `force` syncs now; otherwise only when signed in and the interval has passed.
	 */
	syncProvider(provider: Provider, force: boolean): Promise<void> {
		const running = this.syncing[provider];
		if (running) return running;
		if (!this.isSignedIn(provider)) return Promise.resolve();
		const s = this.settings;
		if (!this.syncEnabled(provider)) return Promise.resolve();
		const last = this.data.lastSync[provider];
		const due = !last || (s.syncHours > 0 && Date.now() - last >= s.syncHours * 3_600_000);
		if (!force && !due) return Promise.resolve();

		const name = PROVIDER_NAMES[provider];
		const generation = this.generation[provider];
		const job = (async () => {
			try {
				let synced: Contact[];
				let warning = "";
				if (provider === "m365") {
					synced = await fetchFromMicrosoft(this.auth, { contacts: s.syncContacts, people: s.syncPeople, peopleLimit: s.peopleLimit });
				} else {
					const result = await fetchFromGoogle(this.google, { contacts: s.googleSyncContacts, other: s.googleSyncOther });
					synced = result.contacts;
					warning = result.errors.join(" ");
				}
				// Signed out while this was running: drop the result.
				if (generation !== this.generation[provider]) return;
				this.data.synced[provider] = synced;
				this.data.lastSync[provider] = Date.now();
				this.lastSyncError[provider] = warning;
				await this.saveSettings();
				if (warning) new Notice(`${name}: ${synced.length} contacts synced, but part of the sync failed. ${warning}`);
				else if (force) new Notice(`Contacts: ${synced.length} synced from ${name}.`);
				this.refreshSettings();
			} catch (err) {
				const message = (err as Error).message;
				this.lastSyncError[provider] = message;
				console.error(`Contact Links: ${name} sync failed`, err);
				// An expired session already showed its own notice.
				if (force && !(err instanceof NotSignedInError && err.expired)) new Notice(`${name} contact sync failed: ${message}`);
			} finally {
				delete this.syncing[provider];
			}
		})();
		this.syncing[provider] = job;
		return job;
	}

	/** Re-render the settings tab if it is open, so counts and sign-in state stay current. */
	refreshSettings(): void {
		try {
			this.settingTab.update();
		} catch {
			// The settings window has not been set up yet; it renders fresh when opened.
		}
	}

	async saveSettings(): Promise<void> {
		await this.saveData(this.data);
		this.rebuild();
	}

	private saveSoon = debounce(() => void this.saveData(this.data), 2000, true);

	recordUse(key: string): void {
		// Count the pick for every copy of a merged contact, so it survives the merge changing.
		const id = key.startsWith("contact:") ? key.slice("contact:".length) : "";
		const contact = id ? this.merged.find((c) => c.id === id) : undefined;
		for (const k of [key, ...(contact?.otherIds ?? []).map((other) => `contact:${other}`)]) {
			this.data.usage[k] = (this.data.usage[k] ?? 0) + 1;
		}
		this.index.setUsage(mergedUsage(this.merged, this.data.usage));
		this.saveSoon();
	}

	importFromFile(done?: () => void): void {
		const input = createEl("input", {
			type: "file",
			attr: { accept: ".csv,.vcf,.vcard,text/csv,text/vcard,text/x-vcard" },
		});
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

	/** Remove the contacts imported from files. Synced contacts go when you sign out of their account. */
	async clearContacts(): Promise<void> {
		this.data.contacts = [];
		await this.saveSettings();
		new Notice("Contacts imported from files cleared. Person notes are kept.");
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
