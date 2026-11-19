// Google sign-in for a desktop app: OAuth authorization code with PKCE and a loopback redirect.
// Google's device-code flow does not allow the contacts scopes, so the browser returns to a
// short-lived listener on 127.0.0.1. That needs Node's http module, so signing in works on desktop
// only; the synced contacts reach mobile through the vault's plugin data.
// Your own Google Cloud OAuth client (type "Desktop app"); the session and the client secret are
// kept in Obsidian's secret storage on this device, outside the vault.
import { App, Platform, requestUrl } from "obsidian";
import type { AddressInfo } from "net";
import { hashString, NotSignedInError } from "../m365/auth";

/** Read-only: your saved contacts and the "other contacts" Gmail collects. */
export const GOOGLE_SCOPES = [
	"https://www.googleapis.com/auth/contacts.readonly",
	"https://www.googleapis.com/auth/contacts.other.readonly",
];

const AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const SIGN_IN_TIMEOUT_MS = 5 * 60 * 1000;

interface Session {
	client_id: string;
	access_token: string;
	refresh_token: string;
	expires_at: number;
	account: string;
}

interface TokenResponse {
	access_token?: string;
	refresh_token?: string;
	expires_in?: number;
	id_token?: string;
	error?: string;
	error_description?: string;
}

export class GoogleAuth {
	private _session: Session | null | undefined = undefined;
	private refreshing: Promise<void> | null = null;
	/** Bumped when a session begins or ends, so late responses from an older session are dropped. */
	private epoch = 0;
	private readonly sessionId: string;
	private readonly secretKeyId: string;

	constructor(
		private app: App,
		private clientId: () => string,
		private onSessionEnded: () => void,
	) {
		const vault = hashString(app.vault.getName());
		this.sessionId = `contact-links-google-session-${vault}`;
		this.secretKeyId = `contact-links-google-client-secret-${vault}`;
	}

	/** Secret storage (Obsidian 1.11.4+) holds the session. */
	get available(): boolean {
		return typeof this.app.secretStorage?.getSecret === "function";
	}

	/** Signing in needs a local listener for the browser to return to. */
	get canSignIn(): boolean {
		return this.available && Platform.isDesktopApp;
	}

	private get session(): Session | null {
		if (this._session === undefined) {
			try {
				const raw = this.available ? this.app.secretStorage.getSecret(this.sessionId) : null;
				this._session = raw ? (JSON.parse(raw) as Session) : null;
			} catch {
				this._session = null;
			}
		}
		return this._session;
	}

	private save(session: Session | null): void {
		this._session = session;
		// Secret storage has no delete; an empty value reads back as signed out.
		if (this.available) this.app.secretStorage.setSecret(this.sessionId, session ? JSON.stringify(session) : "");
	}

	get signedIn(): boolean {
		const s = this.session;
		return Boolean(s?.refresh_token && s.client_id === this.clientId().trim());
	}

	get account(): string {
		return this.session?.account ?? "";
	}

	get clientSecret(): string {
		return (this.available && this.app.secretStorage.getSecret(this.secretKeyId)) || "";
	}

	set clientSecret(value: string) {
		if (this.available) this.app.secretStorage.setSecret(this.secretKeyId, value.trim());
	}

	signOut(): void {
		this.epoch++;
		this.save(null);
	}

	/**
	 * Open Google's consent page in the browser and wait for it to come back to a local listener.
	 * Resolves once signed in; rejects on error, timeout or `cancel.cancelled`.
	 */
	async signIn(cancel: { cancelled: boolean; onCancel?: () => void }, openUrl: (url: string) => void): Promise<void> {
		if (!this.canSignIn) throw new Error("Google sign-in works in Obsidian on desktop.");
		const clientId = this.clientId().trim();
		const verifier = randomString(64);
		const state = randomString(24);
		const challenge = await sha256Base64Url(verifier);

		// Required here, on desktop only, so the plugin still loads on mobile, where Node modules do not exist.
		// require rather than import(): esbuild leaves it as is, and Obsidian's renderer provides it.
		if (!Platform.isDesktop) throw new Error("Google sign-in works in Obsidian on desktop.");
		const http = require("http") as typeof import("http");
		let finish: (result: { code?: string; error?: string }) => void = () => {};
		const result = new Promise<{ code?: string; error?: string }>((resolve) => (finish = resolve));
		const server = http.createServer((req, res) => {
			const url = new URL(req.url ?? "/", "http://127.0.0.1");
			const code = url.searchParams.get("code");
			const error = url.searchParams.get("error");
			// The browser also asks for /favicon.ico and the like.
			if ((!code && !error) || url.searchParams.get("state") !== state) {
				res.writeHead(404).end();
				return;
			}
			res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
			res.end(
				`<!doctype html><meta charset="utf-8"><title>Contact Links</title>` +
					`<p style="font-family:sans-serif">${error ? "Sign-in was not completed." : "Signed in. You can close this tab and return to Obsidian."}</p>`,
			);
			finish(code ? { code } : { error: error ?? "unknown" });
		});
		const timer = window.setTimeout(() => finish({ error: "timeout" }), SIGN_IN_TIMEOUT_MS);
		cancel.onCancel = () => finish({ error: "cancelled" });

		try {
			await new Promise<void>((resolve, reject) => {
				server.once("error", reject);
				server.listen(0, "127.0.0.1", () => resolve());
			});
			const redirectUri = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
			const params = new URLSearchParams({
				client_id: clientId,
				redirect_uri: redirectUri,
				response_type: "code",
				scope: [...GOOGLE_SCOPES, "openid", "email"].join(" "),
				code_challenge: challenge,
				code_challenge_method: "S256",
				state,
				// Always return a refresh token, also when the user has consented before.
				access_type: "offline",
				prompt: "consent",
			});
			openUrl(`${AUTH_URL}?${params.toString()}`);

			const { code, error } = await result;
			if (cancel.cancelled || error === "cancelled") throw new Error("Sign-in cancelled.");
			if (error === "timeout") throw new Error("Sign-in took too long. Start again.");
			if (error === "access_denied") throw new Error("Sign-in was declined.");
			if (!code) throw new Error(`Google sign-in failed: ${error}`);

			const res = await this.postToken({
				grant_type: "authorization_code",
				code,
				code_verifier: verifier,
				redirect_uri: redirectUri,
				client_id: clientId,
			});
			if (res.status !== 200) throw new Error(describeGoogleError(res.json));
			this.epoch++;
			this.storeTokens(res.json, clientId);
		} finally {
			window.clearTimeout(timer);
			server.close();
		}
	}

	private storeTokens(tokens: TokenResponse, clientId: string): void {
		const previous = this.session;
		const claims = jwtClaims(tokens.id_token);
		const email = typeof claims?.email === "string" ? claims.email : "";
		this.save({
			client_id: clientId,
			access_token: tokens.access_token ?? "",
			refresh_token: tokens.refresh_token || previous?.refresh_token || "",
			expires_at: Date.now() + (Number(tokens.expires_in) || 3600) * 1000,
			account: email || previous?.account || "",
		});
	}

	async accessToken(force = false): Promise<string> {
		if (!this.signedIn) throw new NotSignedInError("Not signed in to Google.");
		const s = this.session!;
		if (!force && s.access_token && s.expires_at - 120_000 > Date.now()) return s.access_token;
		this.refreshing ??= this.refresh().finally(() => (this.refreshing = null));
		await this.refreshing;
		if (!this.signedIn) throw new NotSignedInError("Not signed in to Google.");
		return this.session!.access_token;
	}

	private async refresh(): Promise<void> {
		const s = this.session!;
		const epoch = this.epoch;
		const res = await this.postToken({ grant_type: "refresh_token", client_id: s.client_id, refresh_token: s.refresh_token });
		if (epoch !== this.epoch) throw new NotSignedInError("Not signed in to Google.");
		if (res.status === 200) {
			this.storeTokens(res.json, s.client_id);
			return;
		}
		if (res.json?.error === "invalid_grant" || res.json?.error === "invalid_client") {
			this.signOut();
			this.onSessionEnded();
			const expired = new NotSignedInError(`Your Google sign-in has expired. Sign in again. (${describeGoogleError(res.json)})`);
			expired.expired = true;
			throw expired;
		}
		throw new Error(describeGoogleError(res.json));
	}

	/** Google's "Desktop app" clients come with a client secret; send it when one is set. */
	private async postToken(fields: Record<string, string>): Promise<{ status: number; json: TokenResponse }> {
		const secret = this.clientSecret;
		if (secret) fields = { ...fields, client_secret: secret };
		const res = await requestUrl({
			url: TOKEN_URL,
			method: "POST",
			contentType: "application/x-www-form-urlencoded",
			body: new URLSearchParams(fields).toString(),
			throw: false,
		});
		let json: TokenResponse = {};
		try {
			json = res.json as TokenResponse;
		} catch {
			// not JSON
		}
		return { status: res.status, json };
	}
}

const GOOGLE_HINTS: Record<string, string> = {
	invalid_client: "Check the client ID and client secret of your Google OAuth client.",
	invalid_grant: "If your Google Cloud app is in Testing, Google ends sign-ins after 7 days; publish the app or set it to Internal.",
	redirect_uri_mismatch: "Create the OAuth client as type Desktop app, not Web application.",
};

export function describeGoogleError(body: TokenResponse | null | undefined): string {
	if (!body) return "No response from Google.";
	const code = String(body.error ?? "");
	const message = String(body.error_description || code || "Unknown error.");
	return GOOGLE_HINTS[code] ? `${message} ${GOOGLE_HINTS[code]}` : message;
}

function randomString(bytes: number): string {
	return base64Url(crypto.getRandomValues(new Uint8Array(bytes)));
}

async function sha256Base64Url(text: string): Promise<string> {
	return base64Url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text))));
}

function base64Url(bytes: Uint8Array): string {
	let s = "";
	for (const b of bytes) s += String.fromCharCode(b);
	return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function jwtClaims(token: string | undefined): Record<string, unknown> | null {
	try {
		const part = token!.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
		const bytes = Uint8Array.from(atob(part), (c) => c.charCodeAt(0));
		return JSON.parse(new TextDecoder().decode(bytes)) as Record<string, unknown>;
	} catch {
		return null;
	}
}
