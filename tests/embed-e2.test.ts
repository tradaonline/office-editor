import assert from "node:assert";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import {
  EDITABLE_EXTENSIONS,
  getTargetOoxmlType,
  getSuggestedName,
  createSavedPayload,
  shouldBlockExport,
} from "../utils/editor/utils";

test("legacy and modern office formats editable", () => {
  const requiredEditable = [
    "docx",
    "xlsx",
    "pptx",
    "doc",
    "xls",
    "ppt",
    "odt",
    "ods",
    "odp",
    "rtf",
  ];
  for (const ext of requiredEditable) {
    assert.strictEqual(
      EDITABLE_EXTENSIONS.has(ext),
      true,
      `Extension ${ext} must be editable`
    );
  }
  assert.strictEqual(EDITABLE_EXTENSIONS.has("pdf"), false, "PDF must not be editable");
});

test("target OOXML types mapped correctly", () => {
  assert.strictEqual(getTargetOoxmlType("doc"), "docx");
  assert.strictEqual(getTargetOoxmlType("docx"), "docx");
  assert.strictEqual(getTargetOoxmlType("odt"), "docx");
  assert.strictEqual(getTargetOoxmlType("rtf"), "docx");

  assert.strictEqual(getTargetOoxmlType("xls"), "xlsx");
  assert.strictEqual(getTargetOoxmlType("xlsx"), "xlsx");
  assert.strictEqual(getTargetOoxmlType("ods"), "xlsx");

  assert.strictEqual(getTargetOoxmlType("ppt"), "pptx");
  assert.strictEqual(getTargetOoxmlType("pptx"), "pptx");
  assert.strictEqual(getTargetOoxmlType("odp"), "pptx");
});

test("suggestedName replaces original extension with target OOXML", () => {
  assert.strictEqual(getSuggestedName("report.doc", "docx"), "report.docx");
  assert.strictEqual(getSuggestedName("data.xls", "xlsx"), "data.xlsx");
  assert.strictEqual(getSuggestedName("slides.ppt", "pptx"), "slides.pptx");
  assert.strictEqual(getSuggestedName("contract.odt", "docx"), "contract.docx");
  assert.strictEqual(getSuggestedName("sheet.ods", "xlsx"), "sheet.xlsx");
  assert.strictEqual(getSuggestedName("presentation.odp", "pptx"), "presentation.pptx");
  assert.strictEqual(getSuggestedName("memo.rtf", "docx"), "memo.docx");
  assert.strictEqual(getSuggestedName("already.docx", "docx"), "already.docx");
  assert.strictEqual(getSuggestedName("plan.2026.final.doc", "docx"), "plan.2026.final.docx");
});

test("saved payload preserves v1 and adds title, fileType, suggestedName", () => {
  const dummyBuffer = new ArrayBuffer(8);
  const savedPayload = createSavedPayload({
    requestId: "req-123",
    name: "export.docx",
    originalTitle: "contract.doc",
    data: dummyBuffer,
  });

  assert.strictEqual(savedPayload.type, "xerp-office:saved");
  assert.strictEqual(savedPayload.requestId, "req-123");
  assert.strictEqual(savedPayload.name, "export.docx"); // v1 field preserved
  assert.strictEqual(savedPayload.title, "contract.doc"); // original title added
  assert.strictEqual(savedPayload.fileType, "docx"); // target OOXML extension added
  assert.strictEqual(savedPayload.suggestedName, "contract.docx"); // extension replaced
  assert.strictEqual(savedPayload.data, dummyBuffer);
});

test("canDownload blocking on export", () => {
  assert.strictEqual(shouldBlockExport(true, false), true, "Parent save must handle export");
  assert.strictEqual(shouldBlockExport(true, true), true, "Parent save must handle export");
  assert.strictEqual(shouldBlockExport(false, false), true, "canDownload=false must block user export");
  assert.strictEqual(shouldBlockExport(false, true), false, "canDownload=true must allow user export");
});

test("save guard fails READ_ONLY when document is not in edit mode", () => {
  const pageSrc = readFileSync(new URL("../app/xerp/page.tsx", import.meta.url), "utf8");
  const match = pageSrc.match(/const save\s*=\s*\([^)]*\)\s*=>\s*\{([\s\S]*?\n\s{4}\};)/);
  assert(match, "save function must exist in app/xerp/page.tsx");

  const saveFactory = new Function(
    "editorRef",
    "isEditingRef",
    "pendingSave",
    "fail",
    "getTargetOoxmlType",
    "server",
    "SAVE_TIMEOUT_MS",
    `${match[0].replace(/:\s*string/g, "")}; return save;`
  );

  let failed: { code: string; message: string; requestId?: string } | null = null;
  let downloaded: string | null = null;
  const isEditingRef = { current: false };
  const pendingSave = { current: null as { requestId: string; timer?: ReturnType<typeof setTimeout> } | null };

  const save = saveFactory(
    { current: { downloadAs: (type: string) => { downloaded = type; } } },
    isEditingRef,
    pendingSave,
    (code: string, message: string, requestId?: string) => {
      failed = { code, message, requestId };
    },
    getTargetOoxmlType,
    { getDocument: () => ({ fileType: "docx" }) },
    5000
  );

  save("req-readonly");
  assert.deepStrictEqual(failed, {
    code: "READ_ONLY",
    message: "Tài liệu ở chế độ chỉ xem",
    requestId: "req-readonly",
  });
  assert.strictEqual(downloaded, null, "must not download when read-only");
  assert.strictEqual(pendingSave.current, null, "pendingSave must remain null when read-only");

  failed = null;
  isEditingRef.current = true;
  save("req-edit");
  assert.strictEqual(failed, null, "must not fail when editing is true");
  assert.strictEqual(downloaded, "docx", "must trigger download when editing");
  const saved = pendingSave.current as { timer?: ReturnType<typeof setTimeout> } | null;
  if (saved?.timer) clearTimeout(saved.timer);
});
