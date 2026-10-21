import assert from "node:assert";
import { test } from "node:test";
import { emptyDocx, emptyXlsx, emptyPptx } from "../utils/editor/empty";
import { EditorServer } from "../utils/editor/server";
import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import zlib from "node:zlib";

function containsUtf16le(buffer: Uint8Array, needle: string): boolean {
  const needleBuf = Buffer.from(needle, "utf16le");
  return Buffer.from(buffer).includes(needleBuf);
}

function containsUtf8(buffer: Uint8Array, needle: string): boolean {
  const needleBuf = Buffer.from(needle, "utf8");
  return Buffer.from(buffer).includes(needleBuf);
}

test("empty templates must be Uint8Array binaries with clean Arial defaults and NO DengXian", () => {
  assert(emptyDocx instanceof Uint8Array, "emptyDocx must be Uint8Array");
  assert(emptyXlsx instanceof Uint8Array, "emptyXlsx must be Uint8Array");
  assert(emptyPptx instanceof Uint8Array, "emptyPptx must be Uint8Array");

  // Verify headers match ONLYOFFICE v10 binary format
  assert(
    Buffer.from(emptyDocx.slice(0, 11)).toString("utf8") === "DOCY;v10;0;",
    "emptyDocx must start with DOCY;v10;0;"
  );
  assert(
    Buffer.from(emptyXlsx.slice(0, 11)).toString("utf8") === "XLSY;v10;0;",
    "emptyXlsx must start with XLSY;v10;0;"
  );
  assert(
    Buffer.from(emptyPptx.slice(0, 11)).toString("utf8") === "PPTY;v10;0;",
    "emptyPptx must start with PPTY;v10;0;"
  );

  // Chinese DengXian (等线) must NOT be present in any template
  const forbiddenFonts = ["等线", "DengXian"];
  for (const forbidden of forbiddenFonts) {
    assert(
      !containsUtf16le(emptyDocx, forbidden) && !containsUtf8(emptyDocx, forbidden),
      `emptyDocx must not contain ${forbidden}`
    );
    assert(
      !containsUtf16le(emptyXlsx, forbidden) && !containsUtf8(emptyXlsx, forbidden),
      `emptyXlsx must not contain ${forbidden}`
    );
    assert(
      !containsUtf16le(emptyPptx, forbidden) && !containsUtf8(emptyPptx, forbidden),
      `emptyPptx must not contain ${forbidden}`
    );
  }

  // Arial must be present as the default font
  assert(
    containsUtf16le(emptyDocx, "Arial"),
    "emptyDocx must contain Arial as default font"
  );
  assert(
    containsUtf16le(emptyXlsx, "Arial"),
    "emptyXlsx must contain Arial as default font"
  );
  assert(
    containsUtf16le(emptyPptx, "Arial"),
    "emptyPptx must contain Arial as default font"
  );
});

test("EditorServer.openNew initializes clean Arial Editor.bin for word, cell, slide", () => {
  const server = new EditorServer();

  const wordDoc = server.openNew("docx");
  assert.strictEqual(wordDoc.documentType, "word");
  const wordBin = (server as any).fsMap.get("Editor.bin");
  assert(wordBin instanceof Uint8Array, "word Editor.bin must exist");
  assert(containsUtf16le(wordBin, "Arial"), "word Editor.bin must have Arial");
  assert(!containsUtf16le(wordBin, "等线"), "word Editor.bin must NOT have 等线");

  const cellDoc = server.openNew("xlsx");
  assert.strictEqual(cellDoc.documentType, "cell");
  const cellBin = (server as any).fsMap.get("Editor.bin");
  assert(cellBin instanceof Uint8Array, "cell Editor.bin must exist");
  assert(containsUtf16le(cellBin, "Arial"), "cell Editor.bin must have Arial");
  assert(!containsUtf16le(cellBin, "等线"), "cell Editor.bin must NOT have 等线");

  const slideDoc = server.openNew("pptx");
  assert.strictEqual(slideDoc.documentType, "slide");
  const slideBin = (server as any).fsMap.get("Editor.bin");
  assert(slideBin instanceof Uint8Array, "slide Editor.bin must exist");
  assert(containsUtf16le(slideBin, "Arial"), "slide Editor.bin must have Arial");
  assert(!containsUtf16le(slideBin, "等线"), "slide Editor.bin must NOT have 等线");
});

test("EditorServer.open with 0-byte file defaults to clean template without corrupting x2t", async () => {
  const server = new EditorServer();

  const emptyDocxFile = new File([new ArrayBuffer(0)], "New_Word.docx");
  await server.open(emptyDocxFile);
  await (server as any).loadPromise;

  const bin = (server as any).fsMap.get("Editor.bin");
  assert(bin instanceof Uint8Array, "0-byte docx must load valid Editor.bin");
  assert(bin.byteLength > 1000, "0-byte docx must use full empty template, not corrupt 11-byte stub");
  assert(containsUtf16le(bin, "Arial"), "0-byte docx Editor.bin must have Arial default");
  assert(!containsUtf16le(bin, "等线"), "0-byte docx Editor.bin must NOT have 等线");
});

test("SDK API limitation documented: DocsAPI has no runtime default font config", () => {
  // DocsAPI config customization.font only styles the loader branding and triggers
  // the commercial license modal in open-source ONLYOFFICE DocumentServer.
  // Native SDK API does not expose a defaultTextFont / defaultRunFont / defaultCellFont option.
  // Default font resolution is entirely governed by OOXML styles/theme in empty.ts templates.
  const apiJsPath = resolve(process.cwd(), "public/v9.3.0.24-1/web-apps/apps/api/documents/api.js");
  if (existsSync(apiJsPath)) {
    const apiJs = readFileSync(apiJsPath, "utf8");
    // Verify customization.font exists in api.js template as branding font
    assert(apiJs.includes("font: {"), "DocsAPI documents customization.font branding");
    // Verify no document.defaultFont or editorConfig.defaultFont exists in api.js
    assert(!apiJs.includes("defaultFont"), "DocsAPI has no defaultFont config property");
    assert(!apiJs.includes("defaultTextFont"), "DocsAPI has no defaultTextFont config property");
  }
});
