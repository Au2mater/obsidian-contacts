import { App, normalizePath, TFile, TFolder } from "obsidian";
import { safeFileName } from "./core/normalize";
import type { Person, PersonNote } from "./core/types";

const asList = (v: unknown): string[] =>
	(Array.isArray(v) ? v : typeof v === "string" ? v.split(",") : [])
		.map((x) => String(x).trim())
		.filter(Boolean);

/**
 * Read the notes in the people folder (and its subfolders) from the metadata cache.
 * Walks only that folder, never the whole vault, and reads no file contents.
 */
export function readPersonNotes(app: App, folder: string): PersonNote[] {
	const root = folder ? app.vault.getFolderByPath(normalizePath(folder)) : app.vault.getRoot();
	const notes: PersonNote[] = [];
	if (!root) return notes;
	// At the vault root, only top-level notes count; a named folder includes its subfolders.
	const stack: TFolder[] = [root];
	while (stack.length) {
		const dir = stack.pop()!;
		for (const child of dir.children) {
			if (child instanceof TFolder) {
				if (folder) stack.push(child);
			} else if (child instanceof TFile && child.extension === "md") {
				notes.push(toPersonNote(app, child));
			}
		}
	}
	return notes;
}

function toPersonNote(app: App, file: TFile): PersonNote {
	const fm = app.metadataCache.getFileCache(file)?.frontmatter ?? {};
	return {
		path: file.path,
		basename: file.basename,
		aliases: asList(fm.aliases ?? fm.alias),
		emails: asList(fm.email ?? fm.emails).map((e) => e.toLowerCase()),
		company: typeof fm.company === "string" ? fm.company : undefined,
	};
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
	await app.fileManager.processFrontMatter(file, (fm: Record<string, unknown>) => {
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
