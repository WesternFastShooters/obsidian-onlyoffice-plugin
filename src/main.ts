import { Plugin } from "obsidian";
import { OfficeFileView } from "./OfficeFileView";
import { VIEW_TYPE_OFFICE, ALL_EXTENSIONS } from "./types";
import { LocalServer } from "./LocalServer";
import * as path from "path";
import "./styles.css";

export default class OnlyOfficePlugin extends Plugin {
  private server: LocalServer | null = null;
  private editorUrl = "";

  async onload() {
    await this.startServer();

    this.registerView(VIEW_TYPE_OFFICE, (leaf) => new OfficeFileView(leaf, this));

    for (const ext of ALL_EXTENSIONS) {
      this.registerExtensions([ext], VIEW_TYPE_OFFICE);
    }

    this.addCommand({
      id: "office-save",
      name: "OnlyOffice: 保存",
      checkCallback: (checking) => {
        const view = this.app.workspace.getActiveViewOfType(OfficeFileView);
        if (!view) return false;
        if (!checking) {
          // Ctrl+S inside the editor iframe triggers the editor's own save flow.
        }
        return true;
      },
    });
  }

  onunload() {
    if (this.server) {
      this.server.stop();
      this.server = null;
    }
  }

  private async startServer() {
    const pluginDir = this.getPluginDir();
    const assetsDir = path.join(pluginDir, "assets", "onlyoffice");
    this.server = new LocalServer(assetsDir);
    const port = await this.server.start();
    this.editorUrl = `http://127.0.0.1:${port}/editor.html`;
  }

  private getPluginDir(): string {
    const adapter = this.app.vault.adapter as any;
    if (typeof adapter.getBasePath !== "function") {
      throw new Error("Cannot determine vault base path (not desktop?)");
    }
    const basePath: string = adapter.getBasePath();
    return path.join(basePath, ".obsidian", "plugins", "obsidian-onlyoffice");
  }

  getEditorUrl(): string {
    return this.editorUrl;
  }
}
