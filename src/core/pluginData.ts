import type { ContactLinksSettings } from "../settings";
import type { Contact } from "./types";

export type Provider = "m365" | "google";

/** What the plugin keeps in its data.json. */
export interface PluginData {
	settings: ContactLinksSettings;
	/** Contacts imported from files. */
	contacts: Contact[];
	/** Contacts from each connected account, as last synced. Merged with the rest only in memory. */
	synced: Record<Provider, Contact[]>;
	/** How often each person was picked; a small ranking boost. */
	usage: Record<string, number>;
	/** When each account was last synced (ms since epoch). */
	lastSync: Partial<Record<Provider, number>>;
}

/** Read saved data, including the layout from 1.0, where Microsoft contacts sat among the file imports. */
export function migrate(saved: Record<string, unknown>, defaults: ContactLinksSettings): PluginData {
	const old = saved as {
		settings?: Partial<ContactLinksSettings>;
		contacts?: Contact[];
		synced?: Partial<Record<Provider, Contact[]>>;
		usage?: Record<string, number>;
		lastSync?: number | Partial<Record<Provider, number>>;
	};
	const contacts = old.contacts ?? [];
	const synced = { m365: old.synced?.m365 ?? [], google: old.synced?.google ?? [] };
	if (!old.synced) synced.m365 = contacts.filter((c) => c.source === "m365");
	return {
		settings: { ...defaults, ...old.settings },
		contacts: contacts.filter((c) => c.source !== "m365" && c.source !== "google"),
		synced,
		usage: old.usage ?? {},
		lastSync: typeof old.lastSync === "number" ? { m365: old.lastSync } : (old.lastSync ?? {}),
	};
}
