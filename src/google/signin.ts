import { App, Modal, Notice, Setting } from "obsidian";
import type { GoogleAuth } from "./auth";

export interface GoogleSignInHost {
	google: GoogleAuth;
	settings: { googleClientId: string };
	saveSettings(): Promise<void>;
	googleSignedIn(): void;
}

/** Confirm the OAuth client, send the user to Google in the browser, wait for it to come back. */
export class GoogleSignInModal extends Modal {
	private cancel: { cancelled: boolean; onCancel?: () => void } = { cancelled: false };

	constructor(
		app: App,
		private host: GoogleSignInHost,
	) {
		super(app);
	}

	onOpen(): void {
		this.setTitle("Sign in to Google");
		// Signing in means leaving for a browser; a click on the dimmed background on return must not close this.
		this.containerEl.querySelector(".modal-bg")?.addEventListener("click", (e) => e.stopImmediatePropagation(), true);
		this.askForClient();
	}

	private askForClient(): void {
		const { contentEl } = this;
		const { google } = this.host;
		const draft = { clientId: this.host.settings.googleClientId.trim(), secret: google.clientSecret };
		contentEl.createDiv({
			cls: "contact-links-status",
			text: "Your OAuth client from Google Cloud (type Desktop app), with the People API turned on. The client secret is kept in this device's secret storage, not in the vault.",
		});
		new Setting(contentEl)
			.setName("Client ID")
			.addText((t) => t.setPlaceholder("…apps.googleusercontent.com").setValue(draft.clientId).onChange((v) => (draft.clientId = v.trim())));
		new Setting(contentEl).setName("Client secret").addText((t) => {
			t.inputEl.type = "password";
			t.setValue(draft.secret).onChange((v) => (draft.secret = v.trim()));
		});
		let started = false;
		new Setting(contentEl).addButton((b) =>
			b
				.setButtonText("Continue")
				.setCta()
				.onClick(async () => {
					if (started) return;
					if (!draft.clientId) {
						new Notice("Enter the client ID.");
						return;
					}
					started = true;
					this.host.settings.googleClientId = draft.clientId;
					google.clientSecret = draft.secret;
					await this.host.saveSettings();
					contentEl.empty();
					void this.startSignIn();
				}),
		);
	}

	private async startSignIn(): Promise<void> {
		const { contentEl } = this;
		const status = contentEl.createDiv({
			cls: "contact-links-status",
			text: "Sign in to Google in your browser and allow read access to your contacts. This window closes by itself when you are done.",
		});
		const actions = contentEl.createDiv({ cls: "contact-links-actions" });
		let authUrl = "";
		const again = actions.createEl("button", { text: "Open the sign-in page again" });
		again.addEventListener("click", () => authUrl && window.open(authUrl, "_blank"));
		try {
			await this.host.google.signIn(this.cancel, (url) => {
				authUrl = url;
				window.open(url, "_blank");
			});
			if (this.cancel.cancelled) return;
			this.close();
			this.host.googleSignedIn();
		} catch (e) {
			if (this.cancel.cancelled) return;
			actions.remove();
			status.setText((e as Error).message);
			status.addClass("contact-links-error");
			new Setting(contentEl).addButton((b) =>
				b.setButtonText("Back").onClick(() => {
					this.contentEl.empty();
					this.cancel = { cancelled: false };
					this.askForClient();
				}),
			);
		}
	}

	onClose(): void {
		this.cancel.cancelled = true;
		this.cancel.onCancel?.();
		this.contentEl.empty();
	}
}
