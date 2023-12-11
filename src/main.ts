import { Plugin, } from "obsidian";

import * as C from "./constants";
import * as xlutils from "./excel/excelutils"

import * as COMMANDS from "./commands/commands"
import regEvents from "./events/registerevent"

import EditingViewPlugin from "./render/embedded/editing-view-plugin";
import ExcelView, { EXCEL_VIEW } from "./render/embedded/preview-excel-view";

export default class ExcelViewerPlugin extends Plugin {

	/**
	 * Called on plugin load.
	 * This can be when the plugin is enabled or Obsidian is first opened.
	 */
	async onload() {
		console.log("Loading ExcelViewer plugin...")


		// 打开文件的渲染？未生效
		this.registerView(
			EXCEL_VIEW,
			(leaf) =>
				new ExcelView(leaf, this.manifest.id, this.manifest.version)
		);

		// 也会影响嵌入渲染，扩展名要加点
		this.registerExtensions(C.EXCEL_EXTENSIONS, EXCEL_VIEW);

		// 这句影响 md 文档内嵌入 excel 的渲染
		this.registerEditorExtension(
			EditingViewPlugin(this.app, this.manifest.version)
		);

		this.addRibbonIcon("table", "Create Excel File", async () => {
			await xlutils.newExcelFile(null);
		});

		COMMANDS.registerCommands(this);

		regEvents(this);
	}

	/**
	 * Called on plugin unload.
	 * This can be when the plugin is disabled or Obsidian is closed.
	 */
	async onunload() {
		this.app.workspace.detachLeavesOfType(EXCEL_VIEW);
	}
}
