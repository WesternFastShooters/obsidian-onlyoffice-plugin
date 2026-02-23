import { App, TFile } from "obsidian";

export class VaultIO {
  constructor(private app: App) {}

  async readBinary(file: TFile): Promise<ArrayBuffer> {
    return this.app.vault.readBinary(file);
  }

  async writeBinary(file: TFile, data: ArrayBuffer): Promise<void> {
    await this.app.vault.modifyBinary(file, data);
  }

  async createBinary(path: string, data: ArrayBuffer): Promise<TFile> {
    return this.app.vault.createBinary(path, data);
  }
}
