import { PluginValue, ViewPlugin, ViewUpdate } from "@codemirror/view";

import { loadEmbeddedViews } from "./embedded-app-manager";
import { App } from "obsidian";


export default function EditingViewPlugin(app: App, pluginVersion: string) {
	console.log("EditingViewPlugin", pluginVersion);
	return ViewPlugin.fromClass(
		// 此插件负责在实时预览模式下渲染 ExcelViewer
		// 它是为每个打开的叶实例化的
		class EditingViewPlugin implements PluginValue {
			/**
			 * Called whenever the markdown of the current leaf is updated.
			 */
			update(update: ViewUpdate) {
				console.log("EditingViewPlugin.update", update);
				const markdownLeaves =
					app.workspace.getLeavesOfType("markdown");
				const activeLeaf = markdownLeaves.find(
					//@ts-expect-error - private property
					(leaf) => leaf.view.editor.cm === update.view
				);
				if (!activeLeaf) return;

				loadEmbeddedViews(app, pluginVersion, activeLeaf, "source");
			}
		}
	);
}