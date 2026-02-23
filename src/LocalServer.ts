import * as http from "http";
import * as fs from "fs";
import * as path from "path";

const MIME_TYPES: Record<string, string> = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".mjs": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".ttf": "font/ttf",
  ".eot": "application/vnd.ms-fontobject",
  ".xml": "application/xml",
  ".txt": "text/plain; charset=utf-8",
};

export class LocalServer {
  private server: http.Server | null = null;
  private port = 0;
  private rootDir: string;

  constructor(rootDir: string) {
    this.rootDir = rootDir;
  }

  async start(): Promise<number> {
    if (this.server) return this.port;

    return new Promise((resolve, reject) => {
      const server = http.createServer((req, res) => {
        this.handleRequest(req, res);
      });

      server.on("error", (err) => {
        reject(err);
      });

      server.listen(0, "127.0.0.1", () => {
        const addr = server.address();
        if (addr && typeof addr === "object") {
          this.port = addr.port;
          this.server = server;
          console.log(`[OnlyOffice] Local server started on port ${this.port}`);
          resolve(this.port);
        } else {
          reject(new Error("Failed to get server address"));
        }
      });
    });
  }

  stop() {
    if (this.server) {
      this.server.close();
      this.server = null;
      this.port = 0;
      console.log("[OnlyOffice] Local server stopped");
    }
  }

  getBaseUrl(): string {
    return `http://127.0.0.1:${this.port}`;
  }

  private handleRequest(req: http.IncomingMessage, res: http.ServerResponse) {
    let urlPath = decodeURIComponent(req.url || "/");
    const qIdx = urlPath.indexOf("?");
    if (qIdx !== -1) urlPath = urlPath.slice(0, qIdx);

    if (urlPath === "/") urlPath = "/editor.html";

    const filePath = path.join(this.rootDir, urlPath);

    const resolved = path.resolve(filePath);
    if (!resolved.startsWith(path.resolve(this.rootDir))) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }

    fs.stat(resolved, (err, stats) => {
      if (err || !stats.isFile()) {
        res.writeHead(404);
        res.end("Not Found");
        return;
      }

      const ext = path.extname(resolved).toLowerCase();
      const contentType = MIME_TYPES[ext] || "application/octet-stream";

      res.writeHead(200, {
        "Content-Type": contentType,
        "Content-Length": stats.size,
        "Access-Control-Allow-Origin": "*",
        "Cache-Control": "no-cache",
      });

      const stream = fs.createReadStream(resolved);
      stream.pipe(res);
      stream.on("error", () => {
        res.writeHead(500);
        res.end("Internal Server Error");
      });
    });
  }
}
