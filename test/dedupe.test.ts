import { describe, expect, it } from "vitest";
import { dedupeContacts, dropAutomated, isAutomatedEmail, mergedUsage } from "../src/core/dedupe";
import { fromGooglePerson } from "../src/core/googleMap";
import type { Contact } from "../src/core/types";

const c = (name: string, emails: string[], extra: Partial<Contact> = {}): Contact => ({
	id: emails[0]?.toLowerCase() ?? `name:${name.toLowerCase()}`,
	name,
	emails,
	...extra,
});

describe("dedupe across sources", () => {
	it("merges on a shared email, keeping the first contact's name and id", () => {
		const ms = c("Søren Jensen", ["soren@acme.dk"], { source: "m365", rank: 4, title: "CTO" });
		const g = c("Soren J", ["sj@gmail.com", "SOREN@acme.dk"], { source: "google", nickname: "Sø", company: "Acme" });
		const [one, ...rest] = dedupeContacts([ms, g]);
		expect(rest).toHaveLength(0);
		expect(one).toEqual({
			id: "soren@acme.dk", name: "Søren Jensen", emails: ["soren@acme.dk", "sj@gmail.com"],
			source: "m365", sources: ["m365", "google"], otherIds: ["sj@gmail.com"],
			nickname: "Sø", company: "Acme", title: "CTO", rank: 4,
		});
	});
	it("merges chains of shared emails", () => {
		const out = dedupeContacts([c("A", ["a@x.dk"]), c("B", ["b@x.dk"]), c("AB", ["b@x.dk", "a@x.dk"])]);
		expect(out).toHaveLength(1);
		expect(out[0].emails).toEqual(["a@x.dk", "b@x.dk"]);
	});
	it("keeps different people with the same name apart", () => {
		expect(dedupeContacts([c("Anna Berg", ["anna@x.dk"]), c("Anna Berg", ["anna@y.dk"])])).toHaveLength(2);
	});
	it("joins a contact without email to the one person with that name", () => {
		const out = dedupeContacts([c("Anna Berg", ["anna@x.dk"], { source: "m365" }), c("Anna Berg", [], { source: "google", company: "Y" })]);
		expect(out).toHaveLength(1);
		expect(out[0]).toMatchObject({ emails: ["anna@x.dk"], company: "Y", sources: ["m365", "google"] });
		// Ambiguous: two people named Anna Berg with emails, so the emailless one stays separate.
		expect(dedupeContacts([c("Anna Berg", ["a@x.dk"]), c("Anna Berg", ["a@y.dk"]), c("anna berg", [])])).toHaveLength(3);
	});
	it("leaves unique contacts untouched", () => {
		const a = c("A", ["a@x.dk"], { source: "google" });
		expect(dedupeContacts([a])[0]).toBe(a);
	});
	it("takes the highest usage of the copies", () => {
		const [m] = dedupeContacts([c("A", ["a@work.dk"]), c("A", ["a@home.dk", "a@work.dk"])]);
		expect(mergedUsage([m], { "contact:a@home.dk": 3, "contact:a@work.dk": 1 })["contact:a@work.dk"]).toBe(3);
	});
});

describe("same-name merging and automated addresses", () => {
	it("merges the same full name across different emails only when asked", () => {
		const list = [c("Søren Jensen", ["sj@work.dk"], { source: "m365" }), c("soren jensen", ["sj@gmail.com"], { source: "google" }), c("Søren Jensen", [], { source: "google" })];
		expect(dedupeContacts(list)).toHaveLength(3);
		const merged = dedupeContacts(list, { sameName: true });
		expect(merged).toHaveLength(1);
		expect(merged[0]).toMatchObject({ name: "Søren Jensen", emails: ["sj@work.dk", "sj@gmail.com"], sources: ["m365", "google"] });
	});
	it("recognises no-reply and notification senders", () => {
		for (const e of ["noreply@x.dk", "no-reply@x.dk", "do-not-reply@x.dk", "notifications@github.com", "abc@reply.github.com", "1+me@users.noreply.github.com", "x@mail.support.microsoft.com"])
			expect(isAutomatedEmail(e), e).toBe(true);
		for (const e of ["norene@x.dk", "anna@github.com", "replyto@x.dk", "support@microsoft.com"]) expect(isAutomatedEmail(e), e).toBe(false);
	});
	it("drops contacts with only automated addresses and strips them from others", () => {
		const out = dropAutomated([c("Bot", ["noreply@x.dk"]), c("Anna", ["anna@x.dk", "a@reply.github.com"]), c("Kim", [])]);
		expect(out.map((x) => [x.name, x.emails])).toEqual([["Anna", ["anna@x.dk"]], ["Kim", []]]);
		expect(dropAutomated([c("Jane", ["notifications@acme.com", "jane@acme.com"])])[0].id).toBe("jane@acme.com");
	});
});

describe("Google mapping", () => {
	it("maps a saved contact", () => {
		expect(
			fromGooglePerson({
				names: [{ displayName: "Søren Jensen", givenName: "Søren", familyName: "Jensen" }],
				nicknames: [{ value: "Sø" }],
				emailAddresses: [{ value: "soren@acme.dk" }, { value: "Soren@acme.dk" }],
				organizations: [{ name: "Acme", title: "CTO" }],
			}, "saved"),
		).toEqual({ id: "soren@acme.dk", name: "Søren Jensen", nickname: "Sø", emails: ["soren@acme.dk"], company: "Acme", title: "CTO", source: "google" });
	});
	it("drops other contacts without a real name", () => {
		expect(fromGooglePerson({ emailAddresses: [{ value: "noreply@x.dk" }] }, "other")).toBeNull();
		expect(fromGooglePerson({ names: [{ displayName: "noreply@x.dk" }], emailAddresses: [{ value: "noreply@x.dk" }] }, "other")).toBeNull();
		expect(fromGooglePerson({ emailAddresses: [{ value: "kim@x.dk" }] }, "saved")?.name).toBe("kim");
		expect(fromGooglePerson({ names: [{ displayName: "Kim Lund" }], emailAddresses: [{ value: "kim@x.dk" }] }, "other")?.name).toBe("Kim Lund");
	});
});
