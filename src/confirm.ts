import { App, Modal, Setting } from "obsidian";

export interface ConfirmOptions {
	title: string;
	text: string;
	/**
	 * Shown as a list under the text, for example the files an action will touch. With `selectable`,
	 * each item has a checkbox (all ticked) and only the ticked ones are passed to `onConfirm`.
	 */
	items?: string[];
	selectable?: boolean;
	confirmText: string;
	onConfirm: (selected: string[]) => void;
}

/** A plain Obsidian confirmation dialog. Enter confirms, Esc cancels. */
export class ConfirmModal extends Modal {
	constructor(
		app: App,
		private opts: ConfirmOptions,
	) {
		super(app);
	}

	onOpen(): void {
		const { opts, contentEl } = this;
		this.setTitle(opts.title);
		contentEl.createEl("p", { text: opts.text });
		const items = opts.items ?? [];
		const selected = new Set(items);
		if (items.length) {
			const list = contentEl.createEl("ul", { cls: "contact-links-confirm-list" });
			for (const item of items) {
				const li = list.createEl("li");
				if (!opts.selectable) {
					li.setText(item);
					continue;
				}
				li.addClass("mod-selectable");
				const label = li.createEl("label");
				const box = label.createEl("input", { type: "checkbox" });
				box.checked = true;
				box.addEventListener("change", () => {
					if (box.checked) selected.add(item);
					else selected.delete(item);
					confirmButton.disabled = selected.size === 0;
				});
				label.appendText(` ${item}`);
			}
		}
		let confirmButton!: HTMLButtonElement;
		const confirm = () => {
			if (opts.selectable && !selected.size) return;
			this.close();
			opts.onConfirm(items.filter((i) => selected.has(i)));
		};
		this.scope.register([], "Enter", (evt) => {
			// Enter on a focused Cancel (or any other button) presses that button, not this one.
			const focused = activeDocument.activeElement;
			if (focused?.instanceOf(HTMLButtonElement) && focused !== confirmButton) return true;
			evt.preventDefault();
			confirm();
			return false;
		});
		new Setting(contentEl)
			.addButton((b) => b.setButtonText("Cancel").onClick(() => this.close()))
			.addButton((b) => {
				b.setButtonText(opts.confirmText).setCta().onClick(confirm);
				confirmButton = b.buttonEl;
			});
	}

	onClose(): void {
		this.contentEl.empty();
	}
}
