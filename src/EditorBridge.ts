import type {
  PluginToEditorMessage,
  EditorToPluginMessage,
} from "./types";

type PendingRequest = {
  resolve: (msg: EditorToPluginMessage) => void;
  reject: (err: Error) => void;
  timer: ReturnType<typeof setTimeout>;
};

export class EditorBridge {
  private iframe: HTMLIFrameElement;
  private ready = false;
  private readyPromise: Promise<void>;
  private readyResolve!: () => void;
  private pending = new Map<string, PendingRequest>();
  private messageHandler: (e: MessageEvent) => void;
  private requestCounter = 0;
  private onSaveCallback: ((data: ArrayBuffer, fileName: string) => void) | null = null;

  constructor(
    container: HTMLElement,
    editorUrl: string,
  ) {
    this.iframe = document.createElement("iframe");
    this.iframe.style.width = "100%";
    this.iframe.style.height = "100%";
    this.iframe.style.border = "none";
    this.iframe.setAttribute("sandbox", "allow-scripts allow-same-origin allow-popups allow-forms");
    this.iframe.src = editorUrl;
    container.appendChild(this.iframe);

    this.readyPromise = new Promise((resolve) => {
      this.readyResolve = resolve;
    });

    this.messageHandler = (e: MessageEvent) => {
      this.handleMessage(e.data);
    };
    window.addEventListener("message", this.messageHandler);
  }

  onSave(callback: (data: ArrayBuffer, fileName: string) => void) {
    this.onSaveCallback = callback;
  }

  private handleMessage(msg: EditorToPluginMessage) {
    if (!msg || !msg.type || !msg.type.startsWith("oo:")) return;

    if (msg.type === "oo:ready") {
      this.ready = true;
      this.readyResolve();
      return;
    }

    if (msg.type === "oo:saved" && !msg.requestId && this.onSaveCallback) {
      this.onSaveCallback(msg.payload.fileData, msg.payload.fileName);
      return;
    }

    if ("requestId" in msg && msg.requestId) {
      const req = this.pending.get(msg.requestId);
      if (req) {
        clearTimeout(req.timer);
        this.pending.delete(msg.requestId);
        if (msg.type === "oo:error") {
          req.reject(new Error(msg.payload.message));
        } else {
          req.resolve(msg);
        }
      }
    }
  }

  private nextId(): string {
    return `req_${++this.requestCounter}_${Date.now()}`;
  }

  private send(msg: PluginToEditorMessage, transfer?: Transferable[]) {
    this.iframe.contentWindow?.postMessage(msg, "*", transfer);
  }

  private request(
    msg: PluginToEditorMessage,
    timeoutMs = 60000,
    transfer?: Transferable[],
  ): Promise<EditorToPluginMessage> {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(msg.requestId);
        reject(new Error(`Request ${msg.type} timed out`));
      }, timeoutMs);

      this.pending.set(msg.requestId, { resolve, reject, timer });
      this.send(msg, transfer);
    });
  }

  async open(
    fileName: string,
    fileData: ArrayBuffer,
    readonly = false,
  ): Promise<void> {
    await this.readyPromise;
    const requestId = this.nextId();
    const dataCopy = fileData.slice(0);
    await this.request(
      {
        type: "oo:open",
        requestId,
        payload: { fileName, fileData: dataCopy, readonly },
      },
      120000,
      [dataCopy],
    );
  }

  async exportAs(
    targetExt: string,
  ): Promise<{ fileData: ArrayBuffer; fileName: string }> {
    await this.readyPromise;
    const requestId = this.nextId();
    const resp = await this.request(
      { type: "oo:export", requestId, payload: { targetExt } },
      120000,
    );
    if (resp.type === "oo:exported") {
      return resp.payload;
    }
    throw new Error("Unexpected response");
  }

  destroy() {
    window.removeEventListener("message", this.messageHandler);
    for (const [, req] of this.pending) {
      clearTimeout(req.timer);
      req.reject(new Error("EditorBridge destroyed"));
    }
    this.pending.clear();
    this.iframe.remove();
  }
}
