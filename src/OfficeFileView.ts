import { FileView, TFile, WorkspaceLeaf, Notice, Platform } from "obsidian";
import { VIEW_TYPE_OFFICE, LEGACY_TO_MODERN } from "./types";
import { EditorBridge } from "./EditorBridge";
import { VaultIO } from "./VaultIO";
import type OnlyOfficePlugin from "./main";

export class OfficeFileView extends FileView {
  private bridge: EditorBridge | null = null;
  private editorContainer: HTMLElement | null = null;
  private loadingEl: HTMLElement | null = null;
  private plugin: OnlyOfficePlugin;
  private vaultIO: VaultIO;

  constructor(leaf: WorkspaceLeaf, plugin: OnlyOfficePlugin) {
    super(leaf);
    this.plugin = plugin;
    this.vaultIO = new VaultIO(this.app);
  }

  getViewType(): string {
    return VIEW_TYPE_OFFICE;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "Office Document";
  }

  getIcon(): string {
    if (!this.file) return "file";
    const ext = this.file.extension;
    if (ext === "docx" || ext === "doc") return "file-text";
    if (ext === "xlsx" || ext === "xls") return "table";
    if (ext === "pptx" || ext === "ppt") return "presentation";
    return "file";
  }

  async onLoadFile(file: TFile): Promise<void> {
    if (Platform.isMobile) {
      this.contentEl.empty();
      this.contentEl.createEl("div", {
        text: "Office 编辑仅支持桌面端",
        cls: "office-unsupported",
      });
      return;
    }

    const ext = file.extension;
    if (ext in LEGACY_TO_MODERN) {
      new Notice(
        `${file.name} 为旧格式(.${ext})，暂不支持直接编辑。请先用其他工具转换为 .${LEGACY_TO_MODERN[ext]} 格式。`,
        8000,
      );
      return;
    }

    this.contentEl.empty();
    this.contentEl.style.padding = "0";
    this.contentEl.addClass("office-view-container");

    this.loadingEl = this.contentEl.createEl("div", {
      cls: "office-loading",
    });
    this.loadingEl.createEl("div", {
      text: "正在加载编辑器...",
      cls: "office-loading-text",
    });

    this.editorContainer = this.contentEl.createEl("div", {
      cls: "office-editor",
    });
    this.editorContainer.style.display = "none";

    try {
      const editorUrl = this.plugin.getEditorUrl();
      const fileData = await this.vaultIO.readBinary(file);

      this.bridge = new EditorBridge(this.editorContainer, editorUrl);

      this.bridge.onSave(async (data: ArrayBuffer, fileName: string) => {
        await this.handleSave(data, fileName);
      });

      await this.bridge.open(file.name, fileData, false);

      if (this.loadingEl) this.loadingEl.remove();
      if (this.editorContainer) {
        this.editorContainer.style.display = "flex";
        this.editorContainer.style.height = "100%";
      }
    } catch (err) {
      console.error("Failed to open office file:", err);
      if (this.loadingEl) {
        this.loadingEl.empty();
        this.loadingEl.createEl("div", {
          text: `加载失败: ${err instanceof Error ? err.message : String(err)}`,
          cls: "office-error",
        });
      }
    }
  }

  private async handleSave(data: ArrayBuffer, _fileName: string) {
    if (!this.file) return;
    try {
      await this.vaultIO.writeBinary(this.file, data);
    } catch (err) {
      console.error("Failed to save:", err);
      new Notice(
        `保存失败: ${err instanceof Error ? err.message : String(err)}`,
      );
    }
  }

  async onUnloadFile(): Promise<void> {
    if (this.bridge) {
      await this.bridge.triggerSaveAndWait(3000);
      this.bridge.destroy();
      this.bridge = null;
    }
    this.contentEl.empty();
  }

  canAcceptExtension(extension: string): boolean {
    return ["docx", "xlsx", "pptx", "doc", "xls", "ppt"].includes(extension);
  }
}
