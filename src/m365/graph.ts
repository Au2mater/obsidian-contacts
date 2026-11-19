import { requestUrl } from "obsidian";
import { combineSynced, fromGraphContact, fromGraphPerson, GraphContact, GraphPerson } from "../core/graphMap";
import type { Contact } from "../core/types";
import { MicrosoftAuth, sleep } from "./auth";

const GRAPH = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES = 5;

interface Page<T> {
	value?: T[];
	"@odata.nextLink"?: string;
}

async function graphGet<T>(auth: MicrosoftAuth, url: string): Promise<Page<T>> {
	let token = await auth.accessToken();
	let retriedAuth = false;
	for (let attempt = 0; ; attempt++) {
		const res = await requestUrl({ url, headers: { Authorization: `Bearer ${token}` }, throw: false });
		if (res.status === 401 && !retriedAuth) {
			retriedAuth = true;
			token = await auth.accessToken(true);
			continue;
		}
		if ([429, 503, 504].includes(res.status) && attempt < MAX_RETRIES) {
			await sleep(retryDelay(res.headers, attempt) * 1000);
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
			if (res.status === 403) throw new Error(`Microsoft refused access: ${message}`);
			throw new Error(`Microsoft Graph error: ${message}`);
		}
		return (body as Page<T>) ?? {};
	}
}

function retryDelay(headers: Record<string, string>, attempt: number): number {
	const header = Object.entries(headers ?? {}).find(([k]) => k.toLowerCase() === "retry-after");
	const seconds = header ? parseInt(header[1], 10) : NaN;
	return Number.isFinite(seconds) && seconds > 0 ? Math.min(seconds, 60) : Math.min(60, 2 ** (attempt + 1));
}

/** Your saved Outlook contacts (default contacts folder). */
async function fetchContacts(auth: MicrosoftAuth): Promise<Contact[]> {
	const select = "displayName,givenName,middleName,surname,nickName,emailAddresses,companyName,jobTitle";
	let url: string | undefined = `${GRAPH}/me/contacts?$select=${select}&$top=500`;
	const out: Contact[] = [];
	while (url) {
		const page: Page<GraphContact> = await graphGet<GraphContact>(auth, url);
		for (const g of page.value ?? []) {
			const c = fromGraphContact(g);
			if (c) out.push(c);
		}
		url = page["@odata.nextLink"];
	}
	return out;
}

/** People you work with, most relevant first: colleagues and people you email, saved or not. */
async function fetchPeople(auth: MicrosoftAuth, limit: number): Promise<Contact[]> {
	const select = "displayName,givenName,surname,scoredEmailAddresses,companyName,jobTitle,personType";
	const top = 100;
	const out: Contact[] = [];
	let rank = 0;
	for (let skip = 0; skip < limit; skip += top) {
		const page = await graphGet<GraphPerson>(auth, `${GRAPH}/me/people?$select=${select}&$top=${top}&$skip=${skip}`);
		const items = page.value ?? [];
		for (const g of items) {
			const c = fromGraphPerson(g, rank++);
			if (c) out.push(c);
		}
		if (items.length < top) break;
	}
	return out;
}

export interface SyncSources {
	contacts: boolean;
	people: boolean;
	peopleLimit: number;
}

export async function fetchFromMicrosoft(auth: MicrosoftAuth, sources: SyncSources): Promise<Contact[]> {
	const [contacts, people] = await Promise.all([
		sources.contacts ? fetchContacts(auth) : Promise.resolve([]),
		sources.people ? fetchPeople(auth, sources.peopleLimit) : Promise.resolve([]),
	]);
	return combineSynced(contacts, people);
}
