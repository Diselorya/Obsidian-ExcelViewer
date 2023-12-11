import { TextFileView, WorkspaceLeaf } from "obsidian";

import { createAppId } from "../../utils/appid";

export const EXCEL_VIEW = "excel";

export default class PreviewRenderPlugin extends TextFileView {
	private appId: string;
	private pluginId: string;
	private pluginVersion: string;

	data: string;

	constructor(leaf: WorkspaceLeaf, pluginId: string, pluginVersion: string) {
		console.log("预览模式.constructor");
		super(leaf);
		this.pluginId = pluginId;
		this.pluginVersion = pluginVersion;
		this.data = "";
		this.appId = createAppId();
	}

	async onOpen() {
		console.log("预览模式.onOpen");
		//Add offset to the container to account for the mobile action bar
		this.containerEl.style.paddingBottom = "48px";

		//Add settings button to action bar
		this.addAction("settings", "Settings", () => {
			//Open settings tab
			(this.app as any).setting.open();
			//Navigate to plugin settings
			(this.app as any).setting.openTabById(this.pluginId);
		});
	}

	async onClose() {
		console.log("预览模式.onClose");
	}

	setViewData(data: string, clear: boolean): void {
		console.log("预览模式.setViewData", data);
		this.data = data;

		//This is only called when the view is initially opened
		if (clear) {
			const container = this.containerEl.children[1];
		}
	}

	clear(): void {
		console.log("预览模式.clear");
		this.data = "{}";
	}

	getViewData(): string {
		console.log("预览模式.getViewData");
		return this.data;
	}

	getViewType() {
		console.log("预览模式.getViewType");
		return EXCEL_VIEW;
	}

	getDisplayText() {
		console.log("预览模式.getDisplayText");
		if (!this.file) return "";

		const fileName = this.file.name;
		const extensionIndex = fileName.lastIndexOf(".");
		return fileName.substring(0, extensionIndex);
	}
}
