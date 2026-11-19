import { requestUrl } from "obsidian";
import { fromGooglePerson, GooglePerson } from "../core/googleMap";
import type { Contact } from "../core/types";
import { sleep } from "../m365/auth";
import type { GoogleAuth } from "./auth";

const API = "https://people.googleapis.com/v1";
const MAX_RETRIES = 5;

interface Page {
	connections?: GooglePerson[];
	otherContacts?: GooglePerson[];
	nextPageToken?: string;
}

async function apiGet(auth: GoogleAuth, url: string): Promise<Page> {
	let token = await auth.accessToken();
	let retriedAuth = false;
	for (let attempt = 0; ; attempt++) {
		const res = await requestUrl({ url, headers: { Authorization: `Bearer ${token}` }, throw: false });
		if (res.status === 401 && !retriedAuth) {
			retriedAuth = true;
			token = await auth.accessToken(true);
			continue;
		}
		if ([429, 500, 503].includes(res.status) && attempt < MAX_RETRIES) {
			await sleep(Math.min(60, 2 ** (attempt + 1)) * 1000);
			continue;
		}
		let body: unknown = null;
		try {
			body = res.json;
		} catch {
			// not JSON
		}
		if (res.status >= 400) {
			const message = (body as { error?: { message?: string } } | null)?.error?.message || `HTTP ${res.status}`;
			if (res.status === 403) throw new Error(`Google refused access: ${message}`);
			throw new Error(`Google People API error: ${message}`);
		}
		return (body as Page) ?? {};
	}
}

async function fetchAll(auth: GoogleAuth, base: string, field: "connections" | "otherContacts", kind: "saved" | "other"): Promise<Contact[]> {
	const out: Contact[] = [];
	let pageToken = "";
	do {
		const page = await apiGet(auth, pageToken ? `${base}&pageToken=${encodeURIComponent(pageToken)}` : base);
		for (const g of page[field] ?? []) {
			const c = fromGooglePerson(g, kind);
			if (c) out.push(c);
		}
		pageToken = page.nextPageToken ?? "";
	} while (pageToken);
	return out;
}

export interface GoogleSources {
	contacts: boolean;
	other: boolean;
}

/**
 * Saved Google contacts first, then "other contacts"; duplicates are merged later with all sources.
 * Each list is fetched on its own, so one Google refuses (for example a permission left unticked on
 * the consent screen) does not cost the other. Throws only when every requested list failed.
 */
export async function fetchFromGoogle(auth: GoogleAuth, sources: GoogleSources): Promise<{ contacts: Contact[]; errors: string[] }> {
	const jobs: { label: string; run: () => Promise<Contact[]> }[] = [];
	if (sources.contacts)
		jobs.push({
			label: "Saved contacts",
			run: () => fetchAll(auth, `${API}/people/me/connections?personFields=names,nicknames,emailAddresses,organizations&pageSize=1000`, "connections", "saved"),
		});
	if (sources.other)
		jobs.push({
			label: "Other contacts",
			run: () => fetchAll(auth, `${API}/otherContacts?readMask=names,emailAddresses&pageSize=1000`, "otherContacts", "other"),
		});
	const results = await Promise.allSettled(jobs.map((j) => j.run()));
	const contacts: Contact[] = [];
	const errors: string[] = [];
	results.forEach((r, i) => {
		if (r.status === "fulfilled") contacts.push(...r.value);
		else errors.push(`${jobs[i].label}: ${(r.reason as Error).message}`);
	});
	if (errors.length && errors.length === results.length) throw (results[0] as PromiseRejectedResult).reason;
	return { contacts, errors };
}
