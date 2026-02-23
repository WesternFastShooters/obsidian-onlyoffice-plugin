export const VIEW_TYPE_OFFICE = "office-view";

export const OFFICE_EXTENSIONS = ["docx", "xlsx", "pptx"] as const;
export const LEGACY_EXTENSIONS = ["doc", "xls", "ppt"] as const;
export const ALL_EXTENSIONS = [...OFFICE_EXTENSIONS, ...LEGACY_EXTENSIONS] as const;

export const LEGACY_TO_MODERN: Record<string, string> = {
  doc: "docx",
  xls: "xlsx",
  ppt: "pptx",
};

export type PluginToEditorMessage =
  | {
      type: "oo:open";
      requestId: string;
      payload: {
        fileName: string;
        fileData: ArrayBuffer;
        readonly: boolean;
      };
    }
  | { type: "oo:save"; requestId: string }
  | {
      type: "oo:export";
      requestId: string;
      payload: { targetExt: string };
    };

export type EditorToPluginMessage =
  | { type: "oo:ready" }
  | { type: "oo:opened"; requestId: string }
  | {
      type: "oo:saved";
      requestId: string;
      payload: { fileData: ArrayBuffer; fileName: string };
    }
  | {
      type: "oo:exported";
      requestId: string;
      payload: { fileData: ArrayBuffer; fileName: string };
    }
  | {
      type: "oo:error";
      requestId: string;
      payload: { code: string; message: string };
    };
