import { contactId, fold } from "./normalize";
import type { Contact } from "./types";

export interface ImportResult {
	contacts: Contact[];
	skipped: number;
}

/** Detect the format from the content and parse it. */
export function parseContacts(text: string): ImportResult {
	const clean = text.replace(/^\uFEFF/, "");
	return /^\s*BEGIN:VCARD/im.test(clean) ? parseVCard(clean) : parseCsv(clean);
}

// ---------------------------------------------------------------- CSV

/** Header names (folded, letters and digits only) for each field, in English and Danish Outlook exports. */
const HEADERS = {
	first: ["firstname", "givenname", "fornavn"],
	middle: ["middlename", "mellemnavn"],
	last: ["lastname", "surname", "familyname", "efternavn"],
	name: ["name", "displayname", "fullname", "navn", "visningsnavn", "fuldtnavn"],
	nickname: ["nickname", "kaldenavn"],
	company: ["company", "organization", "organisation", "firma", "virksomhed"],
	title: ["jobtitle", "stilling", "jobtitel"],
};

const headerKey = (h: string) => fold(h).replace(/[^a-z0-9]/g, "");

/** An email column is any header mentioning mail that is not a type or display-name column. */
const isEmailHeader = (k: string) => /mail/.test(k) && !/type|display|vist|visning/.test(k);

export function parseCsv(text: string): ImportResult {
	const rows = csvRows(text, detectDelimiter(text));
	if (rows.length < 2) return { contacts: [], skipped: 0 };
	const header = rows[0].map(headerKey);
	const col = (names: string[]) => header.findIndex((h) => names.includes(h));
	const idx = Object.fromEntries(
		Object.entries(HEADERS).map(([field, names]) => [field, col(names)]),
	) as Record<keyof typeof HEADERS, number>;
	const emailCols = header.map((h, i) => (isEmailHeader(h) ? i : -1)).filter((i) => i >= 0);

	const contacts: Contact[] = [];
	let skipped = 0;
	for (const row of rows.slice(1)) {
		const get = (i: number) => (i >= 0 ? (row[i] ?? "").trim() : "");
		const emails = unique(emailCols.map(get).filter((e) => e.includes("@")));
		const contact = buildContact({
			name: get(idx.name),
			first: get(idx.first),
			middle: get(idx.middle),
			last: get(idx.last),
			nickname: get(idx.nickname),
			company: get(idx.company),
			title: get(idx.title),
			emails,
		});
		if (contact) contacts.push(contact);
		else if (row.some((c) => c.trim())) skipped++;
	}
	return { contacts, skipped };
}

function detectDelimiter(text: string): string {
	const firstLine = text.split(/\r?\n/, 1)[0] ?? "";
	const unquoted = firstLine.replace(/"[^"]*"/g, "");
	const counts = [",", ";", "\t"].map((d) => [d, unquoted.split(d).length - 1] as const);
	counts.sort((a, b) => b[1] - a[1]);
	return counts[0][1] > 0 ? counts[0][0] : ",";
}

/** RFC 4180 CSV: quoted fields, doubled quotes, newlines inside quotes. */
export function csvRows(text: string, delimiter = ","): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let field = "";
	let quoted = false;
	for (let i = 0; i < text.length; i++) {
		const c = text[i];
		if (quoted) {
			if (c === '"') {
				if (text[i + 1] === '"') {
					field += '"';
					i++;
				} else quoted = false;
			} else field += c;
		} else if (c === '"') quoted = true;
		else if (c === delimiter) {
			row.push(field);
			field = "";
		} else if (c === "\n" || c === "\r") {
			if (c === "\r" && text[i + 1] === "\n") i++;
			row.push(field);
			rows.push(row);
			row = [];
			field = "";
		} else field += c;
	}
	if (field || row.length) {
		row.push(field);
		rows.push(row);
	}
	return rows.filter((r) => r.some((c) => c.trim()));
}

// ---------------------------------------------------------------- vCard

export function parseVCard(text: string): ImportResult {
	const contacts: Contact[] = [];
	let skipped = 0;
	let card: Record<string, string[]> | null = null;
	for (const line of unfoldVCard(text)) {
		if (/^BEGIN:VCARD/i.test(line)) {
			card = {};
			continue;
		}
		if (/^END:VCARD/i.test(line)) {
			if (card) {
				const c = cardToContact(card);
				if (c) contacts.push(c);
				else skipped++;
			}
			card = null;
			continue;
		}
		if (!card) continue;
		const colon = line.indexOf(":");
		if (colon < 0) continue;
		const [rawName, ...params] = line.slice(0, colon).split(";");
		const prop = rawName.replace(/^item\d+\./i, "").toUpperCase();
		const value = decodeValue(line.slice(colon + 1), params);
		(card[prop] ??= []).push(value);
	}
	return { contacts, skipped };
}

/**
 * Join continuation lines: a line starting with space or tab continues the previous one,
 * and a quoted-printable value (vCard 2.1, as Outlook exports it) continues while it ends with "=".
 */
function unfoldVCard(text: string): string[] {
	const out: string[] = [];
	let qpOpen = false;
	for (const line of text.split(/\r?\n/)) {
		if (out.length && qpOpen) {
			out[out.length - 1] = out[out.length - 1].slice(0, -1) + line;
		} else if (out.length && /^[ \t]/.test(line)) {
			out[out.length - 1] += line.slice(1);
		} else {
			out.push(line);
		}
		const cur = out[out.length - 1];
		const head = cur.slice(0, Math.max(cur.indexOf(":"), 0));
		qpOpen = /QUOTED-PRINTABLE/i.test(head) && cur.endsWith("=");
	}
	return out;
}

function decodeValue(value: string, params: string[]): string {
	const p = params.join(";").toUpperCase();
	if (p.includes("QUOTED-PRINTABLE")) {
		const charset = /CHARSET=([^;]+)/.exec(p)?.[1]?.toLowerCase() ?? "utf-8";
		const bytes: number[] = [];
		for (let i = 0; i < value.length; i++) {
			if (value[i] === "=" && /^[0-9A-F]{2}$/i.test(value.slice(i + 1, i + 3))) {
				bytes.push(parseInt(value.slice(i + 1, i + 3), 16));
				i += 2;
			} else bytes.push(value.charCodeAt(i) & 0xff);
		}
		try {
			return new TextDecoder(charset).decode(new Uint8Array(bytes));
		} catch {
			return new TextDecoder("utf-8").decode(new Uint8Array(bytes));
		}
	}
	return value;
}

/** Split a structured value on unescaped semicolons and unescape each part. */
function parts(value: string): string[] {
	// Manual split: lookbehind regexes break older iOS Safari, which Obsidian mobile uses.
	const out: string[] = [];
	let cur = "";
	for (let i = 0; i < value.length; i++) {
		if (value[i] === "\\" && i + 1 < value.length) {
			cur += value[i] + value[i + 1];
			i++;
		} else if (value[i] === ";") {
			out.push(cur);
			cur = "";
		} else cur += value[i];
	}
	out.push(cur);
	return out.map((s) =>
		s
			.replace(/\\n/gi, " ")
			.replace(/\\([,;\\])/g, "$1")
			.trim(),
	);
}

function cardToContact(card: Record<string, string[]>): Contact | null {
	const [last = "", first = "", middle = ""] = parts(card.N?.[0] ?? "");
	return buildContact({
		name: parts(card.FN?.[0] ?? "")[0] ?? "",
		first,
		middle,
		last,
		nickname: parts(card.NICKNAME?.[0] ?? "")[0] ?? "",
		company: parts(card.ORG?.[0] ?? "")[0] ?? "",
		title: parts(card.TITLE?.[0] ?? "")[0] ?? "",
		emails: unique((card.EMAIL ?? []).map((e) => e.trim()).filter((e) => e.includes("@"))),
	});
}

// ---------------------------------------------------------------- shared

interface RawContact {
	name: string;
	first: string;
	middle: string;
	last: string;
	nickname: string;
	company: string;
	title: string;
	emails: string[];
}

function buildContact(r: RawContact): Contact | null {
	const fromParts = [r.first, r.middle, r.last].filter(Boolean).join(" ");
	// Outlook sometimes stores "Last, First" or the email itself as the display name.
	const display = r.name && !r.name.includes("@") ? r.name : "";
	const name = (fromParts || display || r.nickname || r.emails[0]?.split("@")[0] || "")
		.replace(/\s+/g, " ")
		.trim();
	if (!name) return null;
	const emails = r.emails.map((e) => e.replace(/^mailto:/i, ""));
	const contact: Contact = { id: contactId(name, emails), name, emails };
	if (r.nickname && fold(r.nickname) !== fold(name)) contact.nickname = r.nickname;
	if (r.company) contact.company = r.company;
	if (r.title) contact.title = r.title;
	return contact;
}

function unique(xs: string[]): string[] {
	const seen = new Set<string>();
	return xs.filter((x) => {
		const k = x.toLowerCase();
		if (seen.has(k)) return false;
		seen.add(k);
		return true;
	});
}

/** Merge newly imported contacts into the existing list, matching on id. */
export function mergeContacts(existing: Contact[], incoming: Contact[]): { contacts: Contact[]; added: number; updated: number } {
	const byId = new Map(existing.map((c) => [c.id, c]));
	let added = 0;
	let updated = 0;
	for (const c of incoming) {
		if (byId.has(c.id)) updated++;
		else added++;
		byId.set(c.id, c);
	}
	return { contacts: [...byId.values()], added, updated };
}
