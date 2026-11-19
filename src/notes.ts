import { App, normalizePath, TFile, TFolder } from "obsidian";
import { safeFileName } from "./core/normalize";
import type { Person, PersonNote } from "./core/types";

const asList = (v: unknown): string[] =>
	(Array.isArray(v) ? v : typeof v === "string" ? v.split(",") : [])
		.map((x) => String(x).trim())
		.filter(Boolean);

/** Read every note in the people folder from the metadata cache. No file reads. */
export function readPersonNotes(app: App, folder: string): PersonNote[] {
	const prefix = folder ? normalizePath(folder) + "/" : "";
	const notes: PersonNote[] = [];
	for (const file of app.vault.getMarkdownFiles()) {
		if (prefix && !file.path.startsWith(prefix)) continue;
		if (!prefix && file.parent && !file.parent.isRoot()) continue;
		const fm = app.metadataCache.getFileCache(file)?.frontmatter ?? {};
		notes.push({
			path: file.path,
			basename: file.basename,
			aliases: asList(fm.aliases ?? fm.alias),
			emails: asList(fm.email ?? fm.emails).map((e) => e.toLowerCase()),
			company: typeof fm.company === "string" ? fm.company : undefined,
		});
	}
	return notes;
}

/** Return the person's note, creating a minimal one if it does not exist yet. */
export async function ensurePersonNote(app: App, folder: string, person: Person): Promise<TFile> {
	if (person.path) {
		const existing = app.vault.getAbstractFileByPath(person.path);
		if (existing instanceof TFile) return existing;
	}
	await ensureFolder(app, folder);

	const base = safeFileName(person.name) || "Unnamed";
	const candidates = [base];
	if (person.company) candidates.push(`${base} (${safeFileName(person.company)})`);
	if (person.email) candidates.push(`${base} (${safeFileName(person.email)})`);

	for (const name of candidates) {
		const path = normalizePath(folder ? `${folder}/${name}.md` : `${name}.md`);
		const file = app.vault.getAbstractFileByPath(path);
		if (!file) return createNote(app, path, person);
		// Same name already taken: reuse it if it is this person (or has no email to say otherwise).
		if (file instanceof TFile && isSamePerson(app, file, person)) return file;
	}
	// Every candidate is taken by someone else: number it.
	for (let i = 2; ; i++) {
		const path = normalizePath(folder ? `${folder}/${base} ${i}.md` : `${base} ${i}.md`);
		if (!app.vault.getAbstractFileByPath(path)) return createNote(app, path, person);
	}
}

function isSamePerson(app: App, file: TFile, person: Person): boolean {
	const fm = app.metadataCache.getFileCache(file)?.frontmatter ?? {};
	const emails = asList(fm.email ?? fm.emails).map((e) => e.toLowerCase());
	if (!emails.length) return true;
	const mine = (person.contact?.emails ?? (person.email ? [person.email] : [])).map((e) => e.toLowerCase());
	return mine.some((e) => emails.includes(e));
}

async function createNote(app: App, path: string, person: Person): Promise<TFile> {
	const file = await app.vault.create(path, "");
	await app.fileManager.processFrontMatter(file, (fm) => {
		if (person.aliases.length) fm.aliases = person.aliases;
		if (person.email) fm.email = person.email;
		if (person.company) fm.company = person.company;
		if (person.title) fm.role = person.title;
	});
	return file;
}

async function ensureFolder(app: App, folder: string): Promise<void> {
	if (!folder) return;
	const path = normalizePath(folder);
	const existing = app.vault.getAbstractFileByPath(path);
	if (existing instanceof TFolder) return;
	try {
		await app.vault.createFolder(path);
	} catch {
		// Created concurrently, or a file is in the way; vault.create will report the real problem.
	}
}
