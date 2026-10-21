import { DocumentType } from "./types";

export function getFileExt(name: string) {
  const type = name.split(".").pop() || "";
  return type.toLowerCase();
}

export enum AppType {
  word = 1,
  slide = 3,
  cell = 2,
  draw = 4,
  pdf = 5,
}

export const docTypeMap = {
  // Document
  docx: AppType.word,
  doc: AppType.word,
  odt: AppType.word,
  rtf: AppType.word,
  txt: AppType.word,
  html: AppType.word,
  mht: AppType.word,
  epub: AppType.word,
  fb2: AppType.word,
  mobi: AppType.word,
  docm: AppType.word,
  dotx: AppType.word,
  dotm: AppType.word,
  oform: AppType.word,
  docxf: AppType.word,

  // Presentation
  pptx: AppType.slide,
  ppt: AppType.slide,
  odp: AppType.slide,
  ppsx: AppType.slide,
  pptm: AppType.slide,
  ppsm: AppType.slide,
  potx: AppType.slide,
  potm: AppType.slide,
  otp: AppType.slide,
  odg: AppType.slide,

  // Spreadsheet
  xlsx: AppType.cell,
  xls: AppType.cell,
  ods: AppType.cell,
  csv: AppType.cell,
  xlsm: AppType.cell,
  xltx: AppType.cell,
  xltm: AppType.cell,
  xlsb: AppType.cell,
  ots: AppType.cell,

  // Draw
  vsdx: AppType.draw,
  vssx: AppType.draw,
  vstx: AppType.draw,
  vsdm: AppType.draw,
  vssm: AppType.draw,
  vstm: AppType.draw,

  // PDF
  pdf: AppType.pdf,
};

export function getDocumentType(ext: string) {
  const code = docTypeMap[ext.toLowerCase() as keyof typeof docTypeMap];
  const type = AppType[code] as DocumentType;
  return type || DocumentType.Word;
}

/** Returns the URL for creating a new document of the given type. */
export function getNewUrl(type: string) {
  return `/editor?new=${type}`;
}

export const APP_ROOT = process.env.NEXT_PUBLIC_APP_ROOT || "/v9.3.0.24-1";
export const PRELOAD_HTML = "/web-apps/apps/api/documents/preload.html";
export const API_JS = "/web-apps/apps/api/documents/api.js";

export const EDITABLE_EXTENSIONS = new Set([
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
]);

export function getTargetOoxmlType(ext: string): "docx" | "xlsx" | "pptx" {
  const docType = getDocumentType(ext);
  if (docType === "cell") return "xlsx";
  if (docType === "slide") return "pptx";
  return "docx";
}

export function getSuggestedName(originalName: string, targetExt: string): string {
  const lastDot = originalName.lastIndexOf(".");
  const base = lastDot > 0 ? originalName.slice(0, lastDot) : originalName;
  return `${base}.${targetExt}`;
}

export function createSavedPayload(params: {
  requestId: string;
  name: string;
  originalTitle: string;
  data: ArrayBuffer;
}) {
  const origName = params.originalTitle || params.name;
  const targetType = getTargetOoxmlType(getFileExt(origName));
  const suggested = getSuggestedName(origName, targetType);
  return {
    type: "xerp-office:saved" as const,
    requestId: params.requestId,
    name: params.name,
    title: origName,
    fileType: targetType,
    suggestedName: suggested,
    data: params.data,
  };
}

export function shouldBlockExport(hasPendingSave: boolean, canDownload: boolean): boolean {
  if (hasPendingSave) return true;
  return !canDownload;
}
