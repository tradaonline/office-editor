// Drives xerp/harness in headless Chrome (macOS Chrome path) and captures every
// frame's console: open → type → save → reopen, with screenshots.
//   HARNESS=http://192.168.88.99:8091/ SAMPLE=<docx url> NAME="a.docx" MODE=edit|view \
//   TEXT="..." OUT=./out node xerp/e2e/harness-cdp.mjs
import { spawn } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const HARNESS = process.env.HARNESS || "http://192.168.88.99:8091/";
const SAMPLE = process.env.SAMPLE || "https://raw.githubusercontent.com/baotlake/office-website/main/public/files/essay_templates_Essay%20Outline%20Doc.docx";
const NAME = process.env.NAME || "Essay Outline.docx";
const OUT = process.env.OUT || ".";
const port = 9300 + Math.floor(Math.random() * 500);
const profile = mkdtempSync(join(tmpdir(), "p0-chrome-"));
const chrome = spawn("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", [
  "--headless=new", `--remote-debugging-port=${port}`, `--user-data-dir=${profile}`,
  "--no-first-run", "--window-size=1440,900", "about:blank",
], { stdio: "ignore" });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let targets;
for (let i = 0; i < 50; i++) {
  try { targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json(); if (targets.length) break; } catch {}
  await sleep(200);
}
const page = targets.find((t) => t.type === "page");
const ws = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((r) => ws.addEventListener("open", r));
let id = 0;
const pending = new Map();
const consoleLines = [];
ws.addEventListener("message", (ev) => {
  const m = JSON.parse(ev.data);
  if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); return; }
  const p = m.params || {};
  if (m.method === "Runtime.consoleAPICalled") {
    const text = (p.args || []).map((a) => a.value ?? a.description ?? a.type).join(" ");
    if (p.type === "error" || p.type === "warning" || /xerp|x2t|Error|error/.test(text)) consoleLines.push(`[${m.sessionId ? "frame" : "top"}:${p.type}] ${text}`.slice(0, 300));
  } else if (m.method === "Log.entryAdded") {
    consoleLines.push(`[${m.sessionId ? "frame" : "top"}:log-${p.entry.level}] ${p.entry.text} ${p.entry.url || ""}`.slice(0, 300));
  } else if (m.method === "Target.attachedToTarget") {
    const sid = p.sessionId;
    for (const method of ["Runtime.enable", "Log.enable"]) send(method, {}, sid);
    send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, sid);
    send("Runtime.runIfWaitingForDebugger", {}, sid);
  }
});
function send(method, params = {}, sessionId) {
  const msg = { id: ++id, method, params };
  if (sessionId) msg.sessionId = sessionId;
  ws.send(JSON.stringify(msg));
  return new Promise((r) => pending.set(msg.id, r));
}
const evaluate = async (expression) => {
  const r = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
  if (r.result?.exceptionDetails) throw new Error(JSON.stringify(r.result.exceptionDetails).slice(0, 400));
  return r.result?.result?.value;
};
const shot = async (file) => {
  const r = await send("Page.captureScreenshot", { format: "png" });
  writeFileSync(join(OUT, file), Buffer.from(r.result.data, "base64"));
};
const logText = () => evaluate(`document.getElementById('log').innerText`);
const waitFor = async (re, ms) => {
  const end = Date.now() + ms;
  while (Date.now() < end) { const t = await logText(); if (re.test(t)) return t; await sleep(500); }
  return logText();
};

await send("Runtime.enable"); await send("Log.enable"); await send("Page.enable");
await send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true });
await send("Page.navigate", { url: HARNESS });
await sleep(1500);

const result = {};
// 1. Open the sample in edit mode.
await evaluate(`(async () => { const process_mode = ${JSON.stringify(process.env.MODE || "edit")};
  const b = await (await fetch(${JSON.stringify(SAMPLE)})).blob();
  const dt = new DataTransfer(); dt.items.add(new File([b], ${JSON.stringify(NAME)}));
  document.getElementById('file').files = dt.files;
  document.getElementById('mode').value = process_mode;
  document.getElementById('open').click();
  return b.size;
})()`).then((s) => (result.sampleBytes = s));
result.afterOpen = await waitFor(/loaded:|error /, 90000);
await sleep(2000);
await shot("p0-1-open.png");

// 2. Type into the document, then save back to the parent.
if (/loaded:/.test(result.afterOpen)) {
  await send("Input.dispatchMouseEvent", { type: "mousePressed", x: 720, y: 450, button: "left", clickCount: 1 });
  await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: 720, y: 450, button: "left", clickCount: 1 });
  await sleep(500);
  await send("Input.insertText", { text: process.env.TEXT || "XERP P0 da sua tai lieu nay. " });
  await sleep(1500);
  await evaluate(`document.getElementById('save').click()`);
  result.afterSave = await waitFor(/saved:|error /, 60000);
  await shot("p0-2-saved.png");
  // 3. Reopen what was saved: proves the bytes are a valid document with the edit.
  if (/saved:/.test(result.afterSave)) {
    await evaluate(`document.getElementById('reopen').click()`);
    await sleep(1000);
    result.afterReopen = (await waitFor(/loaded:[\s\S]*loaded:|error /, 90000)).split("\n").slice(0, 3).join("\n");
    await sleep(2000);
    await shot("p0-3-reopened.png");
  }
}
result.console = consoleLines.slice(0, 40);
console.log(JSON.stringify(result, null, 2));
ws.close();
chrome.kill();
