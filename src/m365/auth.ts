// Microsoft sign-in with the OAuth device code flow, the same way the Outlook Calendar Panel and
// Microsoft To Do Panel plugins do it: your own Entra app registration (client ID + tenant), and the
// session kept in Obsidian's secret storage on this device, outside the vault.
import { App, requestUrl } from "obsidian";

export class NotSignedInError extends Error {
	expired = false;
}

export interface AuthSettings {
	clientId: string;
	tenant: string;
}

interface Session {
	client_id: string;
	tenant: string;
	access_token: string;
	refresh_token: string;
	expires_at: number;
	account: string;
}

export interface DeviceFlow {
	device_code: string;
	user_code: string;
	verification_uri: string;
	interval?: number;
	expires_in?: number;
	ctx: { clientId: string; tenant: string };
}

/** Read-only: your saved contacts and the people you work with. */
export const SCOPES = ["Contacts.Read", "People.Read"];

const SECRET_PREFIX = "contact-links-m365-session";

export class MicrosoftAuth {
	private _session: Session | null | undefined = undefined;
	private refreshing: Promise<void> | null = null;
	/** Bumped when a session begins or ends, so late responses from an older session are dropped. */
	private epoch = 0;
	private readonly secretId: string;

	constructor(
		private app: App,
		private settings: () => AuthSettings,
		private onSessionEnded: () => void,
	) {
		this.secretId = `${SECRET_PREFIX}-${hashString(app.vault.getName())}`;
	}

	/** Whether this Obsidian version has secret storage (1.11.4+). */
	get available(): boolean {
		return typeof this.app.secretStorage?.getSecret === "function";
	}

	private get session(): Session | null {
		if (this._session === undefined) this._session = this.load();
		return this._session;
	}

	get signedIn(): boolean {
		return Boolean(this.session?.refresh_token && this.matchesSettings());
	}

	get account(): string {
		return this.session?.account ?? "";
	}

	tenant(): string {
		return this.settings().tenant.trim() || "common";
	}

	matchesSettings(): boolean {
		const s = this.session;
		return Boolean(s) && s!.client_id === this.settings().clientId.trim() && s!.tenant === this.tenant();
	}

	private endpoint(path: string, tenant = this.tenant()): string {
		return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/${path}`;
	}

	private scope(): string {
		return [...SCOPES, "offline_access", "openid", "profile"].join(" ");
	}

	private load(): Session | null {
		if (!this.available) return null;
		try {
			const raw = this.app.secretStorage.getSecret(this.secretId);
			return raw ? (JSON.parse(raw) as Session) : null;
		} catch {
			return null;
		}
	}

	private save(session: Session | null): void {
		this._session = session;
		// Secret storage has no delete; an empty value reads back as signed out.
		if (this.available) this.app.secretStorage.setSecret(this.secretId, session ? JSON.stringify(session) : "");
	}

	signOut(): void {
		this.epoch++;
		this.save(null);
	}

	async startDeviceFlow(): Promise<DeviceFlow> {
		const ctx = { clientId: this.settings().clientId.trim(), tenant: this.tenant() };
		const res = await postForm(this.endpoint("devicecode", ctx.tenant), { client_id: ctx.clientId, scope: this.scope() });
		if (res.status !== 200) throw new Error(describeAuthError(res.json));
		return { ...(res.json as Omit<DeviceFlow, "ctx">), ctx };
	}

	/** Polls until the user finishes signing in in the browser, or `cancel.cancelled` is set. */
	async completeDeviceFlow(flow: DeviceFlow, cancel: { cancelled: boolean }): Promise<void> {
		const { ctx } = flow;
		let interval = Math.max(1, Number(flow.interval) || 5);
		const deadline = Date.now() + (Number(flow.expires_in) || 900) * 1000;
		while (!cancel.cancelled) {
			await sleep(interval * 1000);
			if (cancel.cancelled) break;
			if (Date.now() > deadline) throw new Error("The sign-in code expired. Start again.");
			const res = await postForm(this.endpoint("token", ctx.tenant), {
				grant_type: "urn:ietf:params:oauth:grant-type:device_code",
				client_id: ctx.clientId,
				device_code: flow.device_code,
			});
			if (cancel.cancelled) break;
			if (res.status === 200) {
				this.epoch++;
				this.storeTokens(res.json, ctx);
				return;
			}
			const error = res.json?.error;
			if (error === "authorization_pending") continue;
			if (error === "slow_down") {
				interval += 5;
				continue;
			}
			if (error === "authorization_declined") throw new Error("Sign-in was declined.");
			if (error === "expired_token") throw new Error("The sign-in code expired. Start again.");
			throw new Error(describeAuthError(res.json));
		}
		throw new Error("Sign-in cancelled.");
	}

	private storeTokens(tokens: TokenResponse, ctx: { clientId: string; tenant: string }): void {
		const previous = this.session;
		const claims = jwtClaims(tokens.id_token);
		this.save({
			client_id: ctx.clientId,
			tenant: ctx.tenant,
			access_token: tokens.access_token ?? "",
			// Microsoft rotates refresh tokens; keep the newest one.
			refresh_token: tokens.refresh_token || previous?.refresh_token || "",
			expires_at: Date.now() + (Number(tokens.expires_in) || 3600) * 1000,
			account: String(claims?.preferred_username || claims?.email || claims?.name || previous?.account || ""),
		});
	}

	async accessToken(force = false): Promise<string> {
		if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
		const s = this.session!;
		if (!force && s.access_token && s.expires_at - 120_000 > Date.now()) return s.access_token;
		// Concurrent callers share one refresh, so a rotated refresh token is never spent twice.
		this.refreshing ??= this.refresh().finally(() => (this.refreshing = null));
		await this.refreshing;
		if (!this.signedIn) throw new NotSignedInError("Not signed in to Microsoft.");
		return this.session!.access_token;
	}

	private async refresh(): Promise<void> {
		const s = this.session!;
		const ctx = { clientId: s.client_id, tenant: s.tenant };
		const epoch = this.epoch;
		const res = await postForm(this.endpoint("token", ctx.tenant), {
			grant_type: "refresh_token",
			client_id: ctx.clientId,
			refresh_token: s.refresh_token,
			scope: this.scope(),
		});
		if (epoch !== this.epoch) throw new NotSignedInError("Not signed in to Microsoft.");
		if (res.status === 200) {
			this.storeTokens(res.json, ctx);
			return;
		}
		const error = res.json?.error;
		if (error === "invalid_grant" || error === "interaction_required" || error === "invalid_client") {
			this.signOut();
			this.onSessionEnded();
			const expired = new NotSignedInError(`Your Microsoft sign-in has expired. Sign in again. (${describeAuthError(res.json)})`);
			expired.expired = true;
			throw expired;
		}
		throw new Error(describeAuthError(res.json));
	}
}

interface TokenResponse {
	access_token?: string;
	refresh_token?: string;
	expires_in?: number;
	id_token?: string;
	error?: string;
	error_description?: string;
}

async function postForm(url: string, fields: Record<string, string>): Promise<{ status: number; json: TokenResponse }> {
	const res = await requestUrl({
		url,
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

const AUTH_HINTS: Record<string, string> = {
	AADSTS50059: "Set Tenant to your directory (tenant) ID: this app accepts only accounts in its own organisation.",
	AADSTS700016: "Check the client ID and the tenant: Microsoft cannot find that app there.",
	AADSTS7000218: "In the app registration, turn on Authentication → Allow public client flows.",
	AADSTS65001: "The permission has not been granted yet. An administrator may need to approve the app.",
	AADSTS90002: "Check the tenant setting: Microsoft does not know that directory.",
};

export function describeAuthError(body: TokenResponse | null | undefined): string {
	if (!body) return "No response from Microsoft.";
	const text = String(body.error_description || body.error || "Unknown error.");
	const message = text.split(/\r?\n/)[0].replace(/\s*Trace ID:.*$/, "");
	const code = message.match(/AADSTS\d+/)?.[0];
	return code && AUTH_HINTS[code] ? `${message} ${AUTH_HINTS[code]}` : message;
}

function jwtClaims(token: string | undefined): Record<string, unknown> | null {
	try {
		const part = token!.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
		const bytes = Uint8Array.from(atob(part), (c) => c.charCodeAt(0));
		return JSON.parse(new TextDecoder().decode(bytes));
	} catch {
		return null;
	}
}

export function hashString(text: string): string {
	let hash = 0x811c9dc5;
	for (let i = 0; i < text.length; i++) {
		hash ^= text.charCodeAt(i);
		hash = Math.imul(hash, 0x01000193);
	}
	return (hash >>> 0).toString(16).padStart(8, "0");
}

export function sleep(ms: number): Promise<void> {
	return new Promise((resolve) => window.setTimeout(resolve, ms));
}
