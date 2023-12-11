import { TFile, } from "obsidian";
import PLUGIN from "../main"
import * as xlutils from "../excel/excelutils"

import { getBasename } from "../utils/file-path-utils";

export function registerCommands(plugin: PLUGIN) {
    plugin.addCommand({
        id: "create",
        name: "Create Excel Workbook",
        callback: async () => {
            await xlutils.newExcelFile(null);
        },
    });

    plugin.addCommand({
        id: "create-and-embed",
        name: "Create loom and embed it into current file",
        hotkeys: [{ modifiers: ["Mod", "Shift"], key: "+" }],
        editorCallback: async (editor) => {
            const filePath = await xlutils.newExcelFile(null);
            if (!filePath) return;
            const file = plugin.app.vault.getAbstractFileByPath(filePath) as TFile;
            await plugin.app.workspace.getLeaf(true).openFile(file);

            const useMarkdownLinks = (plugin.app.vault as any).getConfig(
                "useMarkdownLinks"
            );

            // Use basename rather than whole name when using Markdownlink like ![abcd](abcd.loom) instead of ![abcd.loom](abcd.loom)
            // It will replace `.loom` to "" in abcd.loom
            const linkText = useMarkdownLinks
                ? `![${getBasename(filePath)}](${encodeURI(filePath)})`
                : `![[${filePath}]]`;

            editor.replaceRange(linkText, editor.getCursor());
            editor.setCursor(
                editor.getCursor().line,
                editor.getCursor().ch + linkText.length
            );
        },
    });

}