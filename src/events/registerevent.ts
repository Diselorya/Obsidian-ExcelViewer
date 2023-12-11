import { Plugin, } from "obsidian";
import { loadPreviewModeApps, purgeEmbeddedExcel, } from "src/render/embedded/embedded-app-manager";


export default function registerEvents(plugin: Plugin) {

    //This event is fired whenever a leaf is opened, close, moved,
    //or the user switches between editing and preview mode
    plugin.registerEvent(
        plugin.app.workspace.on("layout-change", () => {
            const leaves = plugin.app.workspace.getLeavesOfType("markdown");
            purgeEmbeddedExcel(leaves);

            //TODO find a better way to do this
            //Wait for the DOM to update before loading the preview mode apps
            //2ms should be enough time
            setTimeout(() => {
                loadPreviewModeApps(
                    plugin.app,
                    leaves,
                    plugin.manifest.version
                );
            }, 2);
        })
    );
}