import { App, MarkdownView, TFile, WorkspaceLeaf } from "obsidian";

import {
	getEmbeddedExcelLinkEls,
	hasLoadedEmbeddedExcel,
} from "./embed-utils";
import * as C from "../../constants";

import * as xlembed from "../../excel/excelembed";

interface EmbeddedApp {
	id: string;
	containerEl: HTMLElement;
	leaf: WorkspaceLeaf;
	leafFilePath: string; //Leafs and views are reused, so we need to store a value that won't change
	file: TFile;
	mode: "source" | "preview";
}

//Stores all embedded apps
let embeddedApps: EmbeddedApp[] = [];

/**
 * Iterates through all open markdown leaves and then iterates through all embedded loom links
 * for each leaf and renders a loom for each one.
 * This is intended to be used in the `on("layout-change")` callback function
 * @param markdownLeaves - The open markdown leaves
 */
export const loadPreviewModeApps = (
	app: App,
	markdownLeaves: WorkspaceLeaf[],
	pluginVersion: string
) => {
	for (let i = 0; i < markdownLeaves.length; i++) {
		const leaf = markdownLeaves[i];

		const view = leaf.view;

		let mode = "";
		if (view instanceof MarkdownView) {
			mode = view.getMode();
		}

		if (mode === "preview")
			loadEmbeddedViews(app, pluginVersion, leaf, "preview");
	}
};


/**
 * Iterates through all embedded loom links and renders a Loom for each one.
 * Since a leaf can have an editing and reading view, we specify which child
 * to look in
 * @param markdownLeaf - The leaf that contains the markdown view
 * @param mode - The mode of the markdown view (source or preview)
 */
export const loadEmbeddedViews = (
	app: App,
	pluginVersion: string,
	markdownLeaf: WorkspaceLeaf,
	mode: "source" | "preview"
) => {
	console.log("loadEmbeddedExcelViews", markdownLeaf, mode);
	
	const view = markdownLeaf.view as MarkdownView;
	const linkEls = getEmbeddedExcelLinkEls(view, mode);
	console.log("EmbeddedExcelLinkEls", linkEls);

	lock = lock.then(() => {
		return Promise.all(linkEls.map((linkEl) => {
			return processLinkEl(app, pluginVersion, markdownLeaf, linkEl, mode);
		})).then(() => {});
	});
};

/**
 * Removes embedded apps that are found in leaves that are no longer open
 * @param leaves - The open markdown leaves
 */
export const purgeEmbeddedExcel = (leaves: WorkspaceLeaf[]) => {
	embeddedApps = embeddedApps.filter((app) =>
		leaves.find(
			(l) => (l.view as MarkdownView).file?.path === app.leafFilePath
		)
	);
};

let lock = Promise.resolve();

/**
 * Processes an embedded loom link
 * @param linkEl - The link element that contains the embedded loom
 * @param leaf - The leaf that contains the markdown view
 * @returns
 */
const processLinkEl = async (
	app: App,
	pluginVersion: string,
	leaf: WorkspaceLeaf,
	linkEl: HTMLElement,
	mode: "source" | "preview"
) => {
	console.log("processLinkEl.leaf", leaf);
	console.log("processLinkEl.linkEl", linkEl);
	console.log("processLinkEl.mode", mode);
	console.log("processLinkEl.pluginVersion", pluginVersion);
	
	// 判断是否已嵌入了 Excel 表格，如果已经嵌入了，就不再处理，否则会因为异步导致重复嵌入
	if (hasLoadedEmbeddedExcel(linkEl)) return;

	// 设置滚动条
	linkEl.style.overflow = "auto";

	// Excel 数据处理
	xlembed.excelEmbedding(app, leaf, linkEl);

};
