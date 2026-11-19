import { App, Modal, Notice, Setting } from "obsidian";
import type { MicrosoftAuth } from "./auth";

export interface SignInHost {
	auth: MicrosoftAuth;
	settings: { clientId: string; tenant: string };
	saveSettings(): Promise<void>;
	signedIn(): void;
}

/** Device-code sign-in: confirm the app IDs, show a code, wait for the browser sign-in to finish. */
export class SignInModal extends Modal {
	private cancel = { cancelled: false };

	constructor(
		app: App,
		private host: SignInHost,
	) {
		super(app);
	}

	onOpen(): void {
		this.setTitle("Sign in to Microsoft 365");
		// Signing in means leaving for a browser; a click on the dimmed background on return must not close this.
		this.containerEl.querySelector(".modal-bg")?.addEventListener("click", (e) => e.stopImmediatePropagation(), true);
		this.askForApp();
	}

	private askForApp(): void {
		const { contentEl } = this;
		const s = this.host.settings;
		const draft = { clientId: s.clientId.trim(), tenant: s.tenant.trim() || "common" };
		contentEl.createDiv({
			cls: "contact-links-status",
			text: "The app registration in Microsoft Entra. You can reuse the one from Outlook Calendar Panel or Microsoft To Do Panel; it needs the Contacts.Read and People.Read permissions.",
		});
		new Setting(contentEl)
			.setName("Client ID")
			.setDesc("Application (client) ID")
			.addText((t) => t.setPlaceholder("00000000-0000-0000-0000-000000000000").setValue(draft.clientId).onChange((v) => (draft.clientId = v.trim())));
		new Setting(contentEl)
			.setName("Tenant")
			.setDesc("Directory (tenant) ID for an app limited to your organisation; otherwise common.")
			.addText((t) => t.setPlaceholder("common").setValue(draft.tenant).onChange((v) => (draft.tenant = v.trim())));
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
					s.clientId = draft.clientId;
					s.tenant = draft.tenant || "common";
					await this.host.saveSettings();
					contentEl.empty();
					void this.startSignIn();
				}),
		);
	}

	private async startSignIn(): Promise<void> {
		const { contentEl } = this;
		const status = contentEl.createDiv({ cls: "contact-links-status", text: "Asking Microsoft for a sign-in code…" });
		let flow;
		try {
			flow = await this.host.auth.startDeviceFlow();
		} catch (e) {
			if (!this.cancel.cancelled) this.failed(status, `Could not start sign-in: ${(e as Error).message}`);
			return;
		}
		if (this.cancel.cancelled) return;

		status.setText("Open the Microsoft sign-in page, enter this code and sign in. This window closes by itself when you are done.");
		const codeRow = contentEl.createDiv({ cls: "contact-links-code" });
		codeRow.createEl("code", { text: flow.user_code });
		const actions = contentEl.createDiv({ cls: "contact-links-actions" });
		const copy = actions.createEl("button", { text: "Copy code" });
		copy.addEventListener("click", async () => {
			try {
				await navigator.clipboard.writeText(flow.user_code);
				copy.setText("Copied");
			} catch {
				new Notice("Could not copy; select the code instead.");
			}
		});
		const open = actions.createEl("button", { cls: "mod-cta", text: "Open sign-in page" });
		open.addEventListener("click", () => window.open(flow.verification_uri, "_blank"));

		try {
			await this.host.auth.completeDeviceFlow(flow, this.cancel);
			if (this.cancel.cancelled) return;
			this.close();
			this.host.signedIn();
		} catch (e) {
			if (this.cancel.cancelled) return;
			codeRow.remove();
			actions.remove();
			this.failed(status, (e as Error).message);
		}
	}

	private failed(status: HTMLElement, message: string): void {
		status.setText(message);
		status.addClass("contact-links-error");
		new Setting(this.contentEl).addButton((b) =>
			b.setButtonText("Back").onClick(() => {
				this.contentEl.empty();
				this.askForApp();
			}),
		);
	}

	onClose(): void {
		this.cancel.cancelled = true;
		this.contentEl.empty();
	}
}
