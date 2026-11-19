import { createHash } from "node:crypto";
import { get } from "node:http";
import { beforeEach, describe, expect, it, vi } from "vitest";

// The parts of the Obsidian API the Google code uses.
const requestUrl = vi.fn<(req: { url: string; body?: string }) => unknown>();
vi.mock("obsidian", () => ({ Platform: { isDesktop: true, isDesktopApp: true }, requestUrl: (r: { url: string; body?: string }) => requestUrl(r), PluginSettingTab: class {} }));
(globalThis as unknown as { window: typeof globalThis }).window = globalThis;

const { GoogleAuth } = await import("../src/google/auth");
const { fetchFromGoogle } = await import("../src/google/people");
const { migrate } = await import("../src/core/pluginData");
const { DEFAULT_SETTINGS } = await import("../src/settings");

function fakeApp() {
	const secrets = new Map<string, string>();
	return {
		vault: { getName: () => "Vault" },
		secretStorage: { getSecret: (k: string) => secrets.get(k) ?? null, setSecret: (k: string, v: string) => void secrets.set(k, v) },
	};
}

const idToken = (claims: object) => `x.${Buffer.from(JSON.stringify(claims)).toString("base64url")}.y`;
const b64url = (b: Buffer) => b.toString("base64url");
/** Plays the browser coming back to the loopback listener; resolves with the HTTP status. */
const browse = (url: string) =>
	new Promise<number>((resolve, reject) => {
		get(url, (res) => {
			res.resume();
			resolve(res.statusCode ?? 0);
		}).on("error", reject);
	});

describe("Google sign-in (loopback + PKCE)", () => {
	beforeEach(() => requestUrl.mockReset());

	it("exchanges the code from the browser for a session", async () => {
		const auth = new GoogleAuth(fakeApp() as never, () => "cid", () => {});
		auth.clientSecret = "sec";
		requestUrl.mockResolvedValue({ status: 200, json: { access_token: "at", refresh_token: "rt", expires_in: 3600, id_token: idToken({ email: "me@gmail.com" }) } });
		let authUrl!: URL;
		await auth.signIn({ cancelled: false }, (url) => {
			authUrl = new URL(url);
			const redirect = authUrl.searchParams.get("redirect_uri")!;
			const state = authUrl.searchParams.get("state")!;
			void (async () => {
				// A stray request and a forged state are ignored; the real one completes sign-in.
				expect(await browse(`${redirect}/favicon.ico`)).toBe(404);
				expect(await browse(`${redirect}/?code=evil&state=wrong`)).toBe(404);
				expect(await browse(`${redirect}/?code=good&state=${state}`)).toBe(200);
			})();
		});
		expect(authUrl.searchParams.get("code_challenge_method")).toBe("S256");
		expect(authUrl.searchParams.get("redirect_uri")).toMatch(/^http:\/\/127\.0\.0\.1:\d+$/);
		const sent = new URLSearchParams(requestUrl.mock.calls[0][0].body);
		expect(sent.get("code")).toBe("good");
		expect(sent.get("client_secret")).toBe("sec");
		expect(b64url(createHash("sha256").update(sent.get("code_verifier")!).digest())).toBe(authUrl.searchParams.get("code_challenge"));
		expect(auth.signedIn).toBe(true);
		expect(auth.account).toBe("me@gmail.com");
		expect(await auth.accessToken()).toBe("at");
	});

	it("stops listening when cancelled", async () => {
		const auth = new GoogleAuth(fakeApp() as never, () => "cid", () => {});
		const cancel: { cancelled: boolean; onCancel?: () => void } = { cancelled: false };
		let redirect = "";
		const done = auth.signIn(cancel, (url) => {
			redirect = new URL(url).searchParams.get("redirect_uri")!;
			cancel.cancelled = true;
			cancel.onCancel?.();
		});
		await expect(done).rejects.toThrow("cancelled");
		await expect(browse(redirect)).rejects.toThrow();
		expect(auth.signedIn).toBe(false);
	});

	it("fetches every page of saved and other contacts", async () => {
		const app = fakeApp();
		app.secretStorage.setSecret(
			"contact-links-google-session-" + (await import("../src/m365/auth")).hashString("Vault"),
			JSON.stringify({ client_id: "cid", access_token: "at", refresh_token: "rt", expires_at: Date.now() + 3_600_000, account: "" }),
		);
		const auth = new GoogleAuth(app as never, () => "cid", () => {});
		requestUrl.mockImplementation((req?: { url: string }) => {
			const url = req?.url ?? "";
			if (url.includes("/connections") && !url.includes("pageToken"))
				return { status: 200, json: { connections: [{ names: [{ displayName: "A A" }], emailAddresses: [{ value: "a@x.dk" }] }], nextPageToken: "p2" } };
			if (url.includes("/connections")) return { status: 200, json: { connections: [{ names: [{ displayName: "B B" }], emailAddresses: [{ value: "b@x.dk" }] }] } };
			return { status: 200, json: { otherContacts: [{ emailAddresses: [{ value: "noreply@x.dk" }] }, { names: [{ displayName: "C C" }], emailAddresses: [{ value: "c@x.dk" }] }] } };
		});
		const out = await fetchFromGoogle(auth, { contacts: true, other: true });
		expect(out.contacts.map((c) => c.name)).toEqual(["A A", "B B", "C C"]);
		expect(out.errors).toEqual([]);
		expect(requestUrl.mock.calls.map((c) => c[0].url).some((u) => u.includes("pageToken=p2"))).toBe(true);
	});
});

describe("Google partial failures", () => {
	it("keeps saved contacts when other contacts are refused, and throws when everything fails", async () => {
		const app = fakeApp();
		app.secretStorage.setSecret(
			"contact-links-google-session-" + (await import("../src/m365/auth")).hashString("Vault"),
			JSON.stringify({ client_id: "cid", access_token: "at", refresh_token: "rt", expires_at: Date.now() + 3_600_000, account: "" }),
		);
		const auth = new GoogleAuth(app as never, () => "cid", () => {});
		requestUrl.mockImplementation((req?: { url: string }) =>
			req?.url.includes("/connections")
				? { status: 200, json: { connections: [{ names: [{ displayName: "A A" }], emailAddresses: [{ value: "a@x.dk" }] }] } }
				: { status: 403, json: { error: { message: "Insufficient scopes" } } },
		);
		const out = await fetchFromGoogle(auth, { contacts: true, other: true });
		expect(out.contacts.map((c) => c.name)).toEqual(["A A"]);
		expect(out.errors).toEqual(["Other contacts: Google refused access: Insufficient scopes"]);
		await expect(fetchFromGoogle(auth, { contacts: false, other: true })).rejects.toThrow("Insufficient scopes");
	});
});

describe("data migration", () => {
	it("moves 1.0's Microsoft contacts out of the file imports", () => {
		const data = migrate(
			{
				settings: { peopleFolder: "Folk" },
				contacts: [{ id: "f@x.dk", name: "F", emails: ["f@x.dk"] }, { id: "m@x.dk", name: "M", emails: ["m@x.dk"], source: "m365" }],
				usage: { "contact:m@x.dk": 2 },
				lastSync: 123,
			},
			DEFAULT_SETTINGS,
		);
		expect(data.contacts.map((c) => c.name)).toEqual(["F"]);
		expect(data.synced).toEqual({ m365: [{ id: "m@x.dk", name: "M", emails: ["m@x.dk"], source: "m365" }], google: [] });
		expect(data.lastSync).toEqual({ m365: 123 });
		expect(data.usage).toEqual({ "contact:m@x.dk": 2 });
		expect(data.settings).toMatchObject({ peopleFolder: "Folk", googleSyncContacts: true });
		// Saving and loading again changes nothing.
		expect(migrate(JSON.parse(JSON.stringify(data)) as Record<string, unknown>, DEFAULT_SETTINGS)).toEqual(data);
	});
});
