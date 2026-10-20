"use client";

/**
 * XERP embed: the editor inside an XERP iframe.
 *
 * The file never travels by URL. The parent page (an allowed origin) posts its
 * bytes; "save" asks the editor for the converted file and posts it back. The
 * parent stores it (Files replace flow). Nothing leaves this page otherwise:
 * no plugins, no macros, no analytics.
 *
 * Protocol (all messages are objects with `type`):
 *   editor → parent  xerp-office:ready    {version}
 *   parent → editor  xerp-office:open     {file: {name, data: ArrayBuffer}, mode: "view"|"edit",
 *                                          canDownload?, lang?, theme?, user?: {id, name}}
 *   editor → parent  xerp-office:loaded   {ms}
 *   editor → parent  xerp-office:dirty    {dirty}
 *   parent → editor  xerp-office:save     {requestId}
 *   editor → parent  xerp-office:saved    {requestId, name, data: ArrayBuffer}
 *   editor → parent  xerp-office:error    {requestId?, code, message}
 */

import { useEffect, useRef, useState } from "react";
import { API_JS, APP_ROOT, PRELOAD_HTML, getDocumentType, getFileExt } from "@/utils/editor/utils";
import io, { MockSocket } from "@/utils/editor/socket";
import { createFetchProxy } from "@/utils/editor/fetch";
import { createXHRProxy } from "@/utils/editor/xhr";
import { EditorServer } from "@/utils/editor/server";
import { DocEditor } from "@/utils/editor/types";

const PROTOCOL_VERSION = 1;
const MAX_BYTES = 100 * 1024 * 1024;
const SAVE_TIMEOUT_MS = 60_000;

// Build-time list, e.g. "https://app.example.vn,http://192.168.88.99".
const ALLOWED_PARENTS = (process.env.NEXT_PUBLIC_XERP_PARENT_ORIGINS || "")
  .split(",")
  .map((o) => o.trim())
  .filter(Boolean);

type OpenMessage = {
  type: "xerp-office:open";
  file: { name: string; data: ArrayBuffer };
  mode: "view" | "edit";
  /** The editor's own Download menu. Saving works either way. Default true. */
  canDownload?: boolean;
  lang?: string;
  theme?: string;
  user?: { id: string; name: string };
};

const EDITABLE = new Set(["docx", "xlsx", "pptx"]);

export default function Page() {
  const [status, setStatus] = useState("Đang chờ tài liệu…");
  const parentOrigin = useRef<string | null>(null);
  const editorRef = useRef<DocEditor | null>(null);
  const pendingSave = useRef<{ requestId: string; timer: ReturnType<typeof setTimeout> } | null>(null);
  const canDownload = useRef(false);

  useEffect(() => {
    const embedded = window.parent && window.parent !== window;
    if (!embedded || ALLOWED_PARENTS.length === 0) {
      setStatus("Trang này chỉ mở được bên trong XERP.");
      return;
    }

    const post = (msg: Record<string, unknown>, transfer: Transferable[] = []) => {
      if (parentOrigin.current) window.parent.postMessage(msg, parentOrigin.current, transfer);
    };
    const fail = (code: string, message: string, requestId?: string) =>
      post({ type: "xerp-office:error", code, message, requestId });

    // The converted file arrives here. Only a save the parent asked for goes
    // back to it; the editor's own Download menu downloads (or is swallowed).
    const onExport = (data: Uint8Array, title: string) => {
      const pending = pendingSave.current;
      if (pending) {
        clearTimeout(pending.timer);
        pendingSave.current = null;
        const buf = data.slice().buffer;
        post({ type: "xerp-office:saved", requestId: pending.requestId, name: title, data: buf }, [buf]);
        return true;
      }
      return !canDownload.current;
    };

    const server = new EditorServer({ getState: () => ({ plugins: "none" }), onExport });
    const startedAt = performance.now();

    const patchFrame = () => {
      const iframe = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
      const win = iframe?.contentWindow as (typeof window) | null;
      if (!win) throw new Error("Iframe not loaded");
      const xhr = createXHRProxy(win.XMLHttpRequest);
      const fetchProxy = createFetchProxy(win);
      const _Worker = win.Worker;
      xhr.use((request: Request) => server.handleRequest(request));
      fetchProxy.use((request: Request) => server.handleRequest(request));
      Object.assign(win, {
        io,
        XMLHttpRequest: xhr,
        fetch: fetchProxy,
        Worker: function (url: string, options?: WorkerOptions) {
          const u = new URL(url, location.origin);
          return new _Worker(u.href.replace(u.origin, location.origin), options);
        },
      });
    };

    const createEditor = (msg: OpenMessage) => {
      const doc = server.getDocument();
      const editing = msg.mode === "edit" && EDITABLE.has(doc.fileType);
      server.setClient({ buildVersion: window.DocsAPI!.DocEditor.version() });
      editorRef.current = new window.DocsAPI!.DocEditor("placeholder", {
        document: {
          fileType: doc.fileType,
          key: doc.key,
          title: doc.title,
          url: doc.url,
          permissions: {
            edit: editing,
            // Saving goes through the editor's download command, so it needs
            // this even when the menu's Download is not offered (onExport).
            download: true,
            print: false,
            chat: false,
            rename: false,
            protect: editing,
            review: false,
          },
        },
        documentType: getDocumentType(doc.fileType),
        editorConfig: {
          lang: msg.lang || "vi",
          mode: editing ? "edit" : "view",
          coEditing: { mode: "fast", change: false },
          user: server.getUser(),
          customization: {
            uiTheme: msg.theme || "theme-white",
            macros: false,
            plugins: false,
            features: { spellcheck: { change: false } },
          },
        },
        events: {
          onAppReady: () => patchFrame(),
          onDocumentReady: () => {
            setStatus("");
            post({ type: "xerp-office:loaded", ms: Math.round(performance.now() - startedAt) });
          },
          onDocumentStateChange: (e: { data: boolean }) => post({ type: "xerp-office:dirty", dirty: !!e.data }),
          onError: (e: { data?: { errorCode?: number; errorDescription?: string } }) =>
            fail("EDITOR_ERROR", e?.data?.errorDescription || String(e?.data?.errorCode ?? "")),
        },
        type: "desktop",
        width: "100%",
        height: "100%",
      });
    };

    const loadEditor = (msg: OpenMessage) => {
      if (window.DocsAPI?.DocEditor) return createEditor(msg);
      const script = document.createElement("script");
      script.src = APP_ROOT + API_JS;
      script.onload = () => createEditor(msg);
      script.onerror = () => fail("EDITOR_LOAD_FAILED", "Không tải được trình soạn thảo");
      document.head.appendChild(script);
    };

    const open = async (msg: OpenMessage) => {
      const { name, data } = msg.file || ({} as OpenMessage["file"]);
      if (typeof name !== "string" || !(data instanceof ArrayBuffer)) return fail("BAD_MESSAGE", "Thiếu tệp");
      if (data.byteLength > MAX_BYTES) return fail("FILE_TOO_LARGE", "Tệp quá lớn để mở trong trình soạn thảo");
      if (editorRef.current) return fail("ALREADY_OPEN", "Đã mở một tài liệu");
      canDownload.current = msg.canDownload !== false;
      if (msg.user?.id) server.setUser({ id: String(msg.user.id), name: String(msg.user.name || "") });
      setStatus("Đang mở tài liệu…");
      await server.open(new File([data], name), { fileType: getFileExt(name) });
      loadEditor(msg);
    };

    const save = (requestId: string) => {
      const editor = editorRef.current;
      if (!editor) return fail("NOT_OPEN", "Chưa mở tài liệu", requestId);
      if (pendingSave.current) return fail("SAVE_IN_PROGRESS", "Đang lưu", requestId);
      const timer = setTimeout(() => {
        pendingSave.current = null;
        fail("SAVE_TIMEOUT", "Lưu quá lâu, thử lại", requestId);
      }, SAVE_TIMEOUT_MS);
      pendingSave.current = { requestId, timer };
      editor.downloadAs(server.getDocument().fileType);
    };

    const onMessage = (event: MessageEvent) => {
      if (event.source !== window.parent || !ALLOWED_PARENTS.includes(event.origin)) return;
      const msg = event.data;
      if (!msg || typeof msg !== "object" || typeof msg.type !== "string") return;
      // The first allowed parent to open the document owns this editor.
      if (parentOrigin.current && parentOrigin.current !== event.origin) return;
      if (msg.type === "xerp-office:open") {
        parentOrigin.current = event.origin;
        open(msg as OpenMessage).catch(() => fail("OPEN_FAILED", "Không mở được tài liệu"));
      } else if (msg.type === "xerp-office:save" && parentOrigin.current) {
        save(String(msg.requestId || ""));
      }
    };

    window.addEventListener("message", onMessage);
    MockSocket.on("connect", server.handleConnect);
    MockSocket.on("disconnect", server.handleDisconnect);
    // Tell the embedding page we are ready. Chrome names it (ancestorOrigins);
    // elsewhere try each allowed origin, the browser drops the others.
    const embedder = window.location.ancestorOrigins?.[0];
    const readyTo = embedder ? ALLOWED_PARENTS.filter((o) => o === embedder) : ALLOWED_PARENTS;
    for (const origin of readyTo) {
      window.parent.postMessage({ type: "xerp-office:ready", version: PROTOCOL_VERSION }, origin);
    }

    return () => {
      window.removeEventListener("message", onMessage);
      MockSocket.off("connect", server.handleConnect);
      MockSocket.off("disconnect", server.handleDisconnect);
      editorRef.current?.destroyEditor?.();
    };
  }, []);

  return (
    <div className="w-screen h-screen">
      {status && <div className="p-6 text-sm text-gray-600">{status}</div>}
      <div id="placeholder">
        <iframe className="w-0 h-0 hidden" src={APP_ROOT + PRELOAD_HTML}></iframe>
      </div>
    </div>
  );
}
