import { App, MarkdownView } from "obsidian";
import { numToPx } from "src/utils/conversion";
import * as C from "../../constants";
import xlinfo from "../../excel/excelembedinfo";


export const getEmbeddedExcelLinkEls = (
	view: MarkdownView,
	mode: "source" | "preview"
) => {
	console.log("getEmbeddedExcelLinkEls.mode", mode);
	const linkEls: HTMLElement[] = [];

	const { contentEl } = view;
	const el =
		mode === "source"
			? contentEl.querySelector(".markdown-source-view")
			: contentEl.querySelector(".markdown-reading-view");

	if (el) {
		const embeddedLinkEls = el.querySelectorAll(".internal-embed");
		console.log("getEmbeddedExcelLinkEls.embeddedLinkEls", embeddedLinkEls);
		for (let i = 0; i < embeddedLinkEls.length; i++) {
			const linkEl = embeddedLinkEls[i];
			let info = new xlinfo();
			info.loadFormLinkEl(linkEl as HTMLElement);
			console.log("getEmbeddedExcelLinkEls.info", info);
			if (info.isExcel)
				linkEls.push(linkEl as HTMLElement);
			// if (src?.endsWith(EXCEL_EXTENSION))
			// 	linkEls.push(linkEl as HTMLElement);
		}
	}
	return linkEls;
};

export const hasLoadedEmbeddedExcel = (linkEl: HTMLElement) => {
	console.log("hasLoadedEmbeddedExcel.linkEl", linkEl);
	if (linkEl.children.length > 0) {
		const judge = Array.from(linkEl.children).some(child => child.classList.contains(C.EXCEL_VIEW_EL_NAME));			
		console.log("hasLoadedEmbeddedExcel", judge);
		if (judge) return true;
	}

	if (linkEl.parentNode) {
		const judge = Array.from(linkEl.parentNode.children).some(
			child => child.classList.contains(C.EXCEL_VIEW_EL_NAME));	
		if (judge) return true;		
	}
	return false;
};

export const findEmbeddedExcelFile = (
	app: App,
	linkEl: HTMLElement,
	sourcePath: string
) => {
	const src = linkEl.getAttribute("src");
	if (!src) return null;

	//We use the getFirstLinkpathDest to handle absolute links, relative links, and short links
	//in Obsidian.
	return app.metadataCache.getFirstLinkpathDest(src, sourcePath);
};

export const getLinkWidth = (linkEl: HTMLElement, defaultWidth: string) => {
	const width = linkEl.getAttribute("width");
	if (width === null || width === "0") return defaultWidth;
	return numToPx(width);
};

export const getLinkHeight = (linkEl: HTMLElement, defaultHeight: string) => {
	const height = linkEl.getAttribute("height");
	if (height === null || height === "0") return defaultHeight;
	return numToPx(height);
};
