import { describe, expect, it } from "vitest";
import { combineSynced, fromGraphContact, fromGraphPerson } from "../src/core/graphMap";

describe("Microsoft Graph mapping", () => {
	it("maps a saved contact", () => {
		expect(
			fromGraphContact({
				displayName: "Jensen, Søren", givenName: "Søren", surname: "Jensen", nickName: "Sø",
				emailAddresses: [{ address: "soren@acme.dk" }, { address: "SOREN@acme.dk" }, { address: "x" }],
				companyName: "Acme", jobTitle: "CTO",
			}),
		).toEqual({ id: "soren@acme.dk", name: "Søren Jensen", nickname: "Sø", emails: ["soren@acme.dk"], company: "Acme", title: "CTO", source: "m365" });
	});
	it("skips groups and people without email", () => {
		expect(fromGraphPerson({ displayName: "Team", personType: { class: "Group" }, scoredEmailAddresses: [{ address: "t@x.dk" }] }, 0)).toBeNull();
		expect(fromGraphPerson({ displayName: "Anna", personType: { class: "Person" } }, 0)).toBeNull();
		expect(fromGraphPerson({ displayName: "Anna Berg", personType: { class: "Person" }, scoredEmailAddresses: [{ address: "a@x.dk" }] }, 3))
			.toMatchObject({ name: "Anna Berg", rank: 3, source: "m365" });
	});
	it("combines contacts with people, taking the rank", () => {
		const c = fromGraphContact({ givenName: "Søren", surname: "Jensen", nickName: "Sø", emailAddresses: [{ address: "soren@acme.dk" }] })!;
		const p1 = fromGraphPerson({ displayName: "Søren Jensen", scoredEmailAddresses: [{ address: "Soren@acme.dk" }], jobTitle: "CTO" }, 0)!;
		const p2 = fromGraphPerson({ displayName: "Anna Berg", scoredEmailAddresses: [{ address: "a@x.dk" }] }, 1)!;
		const all = combineSynced([c], [p1, p2]);
		expect(all).toHaveLength(2);
		expect(all[0]).toMatchObject({ nickname: "Sø", rank: 0, title: "CTO" });
	});
});
