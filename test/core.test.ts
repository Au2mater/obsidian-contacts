import { describe, expect, it } from "vitest";
import { fold, safeFileName } from "../src/core/normalize";
import { mergeContacts, parseContacts } from "../src/core/importers";
import { mergePeople, PeopleIndex } from "../src/core/people";
import { detectTrigger } from "../src/core/trigger";

const T = { inBrackets: true, triggerChar: "@", bareNames: false, bareMinChars: 3, bracketMinChars: 1 };

describe("fold", () => {
	it("folds Danish and accented letters", () => {
		expect(fold("Søren Ærø Åse")).toBe("soren aero ase");
		expect(fold("José Müller")).toBe("jose muller");
	});
	it("makes safe file names", () => {
		expect(safeFileName('A/B: "C" [x]')).toBe("A B C x");
	});
});

describe("CSV import", () => {
	it("parses an English Outlook export", () => {
		const csv = '﻿"First Name","Middle Name","Last Name","Title","Nickname","Company","Job Title","E-mail Address","E-mail Type","E-mail 2 Address"\r\n' +
			'"Søren","","Jensen","Mr.","Sø","Acme, Inc.","CTO","soren@acme.dk","SMTP","sj@home.dk"\r\n' +
			'"","","","","","","","only@mail.dk","SMTP",""\r\n';
		const { contacts } = parseContacts(csv);
		expect(contacts).toHaveLength(2);
		expect(contacts[0]).toEqual({
			id: "soren@acme.dk", name: "Søren Jensen", nickname: "Sø",
			emails: ["soren@acme.dk", "sj@home.dk"], company: "Acme, Inc.", title: "CTO",
		});
		expect(contacts[1].name).toBe("only");
	});
	it("parses a Danish semicolon export with quoted newlines", () => {
		const csv = 'Fornavn;Efternavn;Firma;Mailadresse\n"Anne";"Ørsted";"Firma\nA/S";anne@x.dk\n';
		const { contacts } = parseContacts(csv);
		expect(contacts[0].name).toBe("Anne Ørsted");
		expect(contacts[0].company).toBe("Firma\nA/S");
		expect(contacts[0].emails).toEqual(["anne@x.dk"]);
	});
	it("merges by id", () => {
		const a = parseContacts("First Name,Last Name,E-mail Address\nA,B,a@b.dk\n").contacts;
		const b = parseContacts("First Name,Last Name,E-mail Address\nA,Bee,A@B.dk\nC,D,c@d.dk\n").contacts;
		const r = mergeContacts(a, b);
		expect(r).toMatchObject({ added: 1, updated: 1 });
		expect(r.contacts.map((c) => c.name)).toEqual(["A Bee", "C D"]);
	});
});

describe("vCard import", () => {
	it("parses vCard 3 and Outlook's quoted-printable vCard 2.1", () => {
		const vcf = [
			"BEGIN:VCARD", "VERSION:3.0", "N:Jensen;Søren;;;", "FN:Søren Jensen", "NICKNAME:Sø",
			"ORG:Acme\\, Inc.;Dev", "EMAIL;TYPE=INTERNET:soren@acme.dk", "END:VCARD",
			"BEGIN:VCARD", "VERSION:2.1", "N;CHARSET=utf-8;ENCODING=QUOTED-PRINTABLE:=C3=98rsted;Anne=",
			"", "FN;CHARSET=windows-1252;ENCODING=QUOTED-PRINTABLE:Anne =D8rs=", "ted",
			"EMAIL;PREF;INTERNET:anne@x.dk", "END:VCARD",
		].join("\r\n");
		const { contacts } = parseContacts(vcf);
		expect(contacts[0]).toMatchObject({ name: "Søren Jensen", nickname: "Sø", company: "Acme, Inc.", emails: ["soren@acme.dk"] });
		expect(contacts[1]).toMatchObject({ name: "Anne Ørsted", emails: ["anne@x.dk"] });
	});
});

describe("people index", () => {
	const contacts = parseContacts(
		"First Name,Last Name,Nickname,E-mail Address\nSøren,Jensen,,soren@acme.dk\nMette,Frederiksen,Mette F,mf@x.dk\nAnna,Jensen,,anna@y.dk\n",
	).contacts;
	const notes = [{ path: "People/Søren Jensen.md", basename: "Søren Jensen", aliases: ["SJ"], emails: [] }];
	const people = mergePeople(contacts, notes);

	it("merges a note with its contact by name", () => {
		expect(people).toHaveLength(3);
		expect(people[0]).toMatchObject({ path: "People/Søren Jensen.md", email: "soren@acme.dk" });
	});
	it("finds by folded prefix, word prefix, initials and alias", () => {
		const idx = new PeopleIndex(people);
		expect(idx.search("soren")[0].person.name).toBe("Søren Jensen");
		expect(idx.search("jens").map((m) => m.person.name)).toEqual(["Søren Jensen", "Anna Jensen"]);
		expect(idx.search("aj")[0].person.name).toBe("Anna Jensen");
		expect(idx.search("mette f")[0]).toMatchObject({ alias: "Mette F" });
		expect(idx.search("mf")[0].person.name).toBe("Mette Frederiksen");
		expect(idx.search("zzz")).toEqual([]);
	});
	it("boosts frequently used people", () => {
		const idx = new PeopleIndex(people, { "contact:anna@y.dk": 10 });
		expect(idx.search("jens")[0].person.name).toBe("Anna Jensen");
	});
});

describe("trigger", () => {
	it("detects open wikilinks", () => {
		expect(detectTrigger("met [[Sør", T)).toEqual({ mode: "link", query: "Sør", start: 4 });
		expect(detectTrigger("met [[Sør|x", T)).toBeNull();
		expect(detectTrigger("[[a]] then", T)).toBeNull();
	});
	it("detects @mentions but not emails", () => {
		expect(detectTrigger("call @søren je", T)).toEqual({ mode: "mention", query: "søren je", start: 5 });
		expect(detectTrigger("mail a@b", T)).toBeNull();
		expect(detectTrigger("`@x", T)).toBeNull();
	});
	it("detects bare names only when enabled", () => {
		expect(detectTrigger("met Søren", T)).toBeNull();
		expect(detectTrigger("met Søren", { ...T, bareNames: true })).toEqual({ mode: "bare", query: "Søren", start: 4 });
		expect(detectTrigger("met So", { ...T, bareNames: true })).toBeNull();
	});
});

describe("link edges", () => {
	it("skips embeds and consumes the rest of an existing link", async () => {
		const { trailingClose } = await import("../src/core/trigger");
		expect(detectTrigger("see ![[Sø", T)).toBeNull();
		expect(trailingClose("]] more", "link")).toBe(2);
		expect(trailingClose("ren Jensen]] more", "link")).toBe(12);
		expect(trailingClose(" and [[x]]", "link")).toBe(0);
		expect(trailingClose("]]", "mention")).toBe(0);
	});
});
