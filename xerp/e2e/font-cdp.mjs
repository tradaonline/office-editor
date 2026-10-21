// XERP Office Editor — Font Selector & Blank OOXML CDP Acceptance
// Verifies 0-byte docx, xlsx, pptx opened via harness:
// 1. Routes to Arial empty templates in deployed editor
// 2. Inner editor DOM font selector strictly shows Arial (not empty, not DengXian)
// 3. Saves blank document via actual UI/protocol
// 4. Unzips saved OOXML and proves Arial defaults in XML structures
// 5. Captures screenshots and writes JSON report
//
// Usage:
//   node xerp/e2e/font-cdp.mjs
//   HARNESS=http://192.168.88.99:8091/ EDITOR=http://192.168.88.99:8090 node xerp/e2e/font-cdp.mjs

import { execFileSync, spawn } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));

const args = process.argv.slice(2);
const getArg = (name, fallback) => {
  const prefix = `--${name}=`;
  const hit = args.find((a) => a.startsWith(prefix));
  return hit ? hit.slice(prefix.length) : (process.env[name.toUpperCase()] || fallback);
};

const HARNESS = getArg("harness", "http://192.168.88.99:8091/");
const EDITOR = getArg("editor", "http://192.168.88.99:8090").replace(/\/$/, "");
const OUT_DIR = resolve(getArg("out", join(__dirname, "out")));
const SCRATCH_DIR = resolve(getArg("scratch", join(tmpdir(), "xerp-office-e2e")));
const TIMEOUT_MS = Number(getArg("timeout", "60000"));

mkdirSync(OUT_DIR, { recursive: true });
mkdirSync(SCRATCH_DIR, { recursive: true });

function getChromePath() {
  if (process.env.CHROME_PATH && existsSync(process.env.CHROME_PATH)) return process.env.CHROME_PATH;
  const candidates = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
  ];
  for (const c of candidates) {
    if (existsSync(c)) return c;
  }
  return "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
}

function getHarnessUrl() {
  const u = new URL(HARNESS);
  u.searchParams.set("editor", EDITOR);
  return u.toString();
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const report = {
    timestamp: new Date().toISOString(),
    harness: HARNESS,
    editor: EDITOR,
    summary: { total: 3, passed: 0, failed: 0 },
    formats: {},
    screenshots: [],
  };

  const port = 9400 + Math.floor(Math.random() * 500);
  const profile = mkdtempSync(join(tmpdir(), "xerp-font-cdp-"));
  const chromeBin = getChromePath();
  const chrome = spawn(chromeBin, [
    "--headless=new",
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profile}`,
    "--no-first-run",
    "--window-size=1440,900",
    "about:blank",
  ], { stdio: "ignore" });

  let ws = null;
  let msgId = 0;
  const pending = new Map();
  const contexts = [];

  function send(method, params = {}, sessionId, timeoutMs = 30000) {
    const id = ++msgId;
    const msg = { id, method, params };
    if (sessionId) msg.sessionId = sessionId;
    ws.send(JSON.stringify(msg));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error(`CDP command timed out: ${method}`));
        }
      }, timeoutMs);
      pending.set(id, {
        resolve: (val) => { clearTimeout(timer); resolve(val); },
        reject: (err) => { clearTimeout(timer); reject(err); },
      });
    });
  }

  const evaluate = async (expression, contextId, sessionId) => {
    const params = { expression, awaitPromise: true, returnByValue: true };
    if (contextId) params.contextId = contextId;
    const r = await send("Runtime.evaluate", params, sessionId);
    if (r.exceptionDetails) {
      throw new Error(r.exceptionDetails.text || JSON.stringify(r.exceptionDetails));
    }
    return r.result?.value;
  };

  const captureScreenshot = async (name) => {
    const filePath = join(OUT_DIR, name);
    try {
      const r = await send("Page.captureScreenshot", { format: "png" });
      writeFileSync(filePath, Buffer.from(r.data, "base64"));
      const scratchPath = join(SCRATCH_DIR, name);
      copyFileSync(filePath, scratchPath);
      report.screenshots.push({ name, path: filePath, scratchPath });
      return scratchPath;
    } catch {
      return null;
    }
  };

  const waitForLog = async (pattern, timeoutMs = TIMEOUT_MS) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const text = await evaluate(`document.getElementById('log')?.innerText || ''`).catch(() => "");
      if (pattern.test(text)) return { ok: true, text };
      await sleep(300);
    }
    const finalLog = await evaluate(`document.getElementById('log')?.innerText || ''`).catch(() => "");
    return { ok: false, text: finalLog };
  };

  const resetHarness = async () => {
    await evaluate(`(() => {
      if (window.__harness?.reset) window.__harness.reset();
      const log = document.getElementById('log');
      if (log) log.innerHTML = '';
      if (window.__harnessMessages) window.__harnessMessages = [];
      window.__lastSaved = null;
    })()`).catch(() => {});
    await sleep(600);
  };

  try {
    let targets = [];
    for (let i = 0; i < 40; i++) {
      try {
        targets = await (await fetch(`http://127.0.0.1:${port}/json`)).json();
        if (targets.length) break;
      } catch {}
      await sleep(200);
    }
    const page = targets.find((t) => t.type === "page");
    if (!page) throw new Error("Chrome page target not found");

    ws = new WebSocket(page.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
      ws.addEventListener("open", resolve, { once: true });
      ws.addEventListener("error", reject, { once: true });
    });

    ws.addEventListener("message", (ev) => {
      const m = JSON.parse(ev.data);
      if (m.id && pending.has(m.id)) {
        const { resolve, reject } = pending.get(m.id);
        pending.delete(m.id);
        if (m.error) reject(new Error(m.error.message || JSON.stringify(m.error)));
        else resolve(m.result || {});
        return;
      }
      if (m.method === "Runtime.executionContextCreated") {
        contexts.push({ sessionId: m.sessionId, ...m.params.context });
      }
    });

    await send("Runtime.enable");
    await send("Page.enable");
    await send("Page.navigate", { url: getHarnessUrl() });
    await sleep(2000);

    const formats = [
      {
        ext: "docx",
        name: "blank.docx",
        requiresFocus: false,
        keyXmlFiles: ["word/styles.xml", "word/fontTable.xml", "word/theme/theme1.xml"],
      },
      {
        ext: "xlsx",
        name: "blank.xlsx",
        requiresFocus: false,
        keyXmlFiles: ["xl/styles.xml", "xl/theme/theme1.xml"],
      },
      {
        ext: "pptx",
        name: "blank.pptx",
        requiresFocus: true,
        focusPos: { x: 600, y: 400 },
        keyXmlFiles: ["ppt/theme/theme1.xml", "ppt/slideMasters/slideMaster1.xml"],
      },
    ];

    for (const fmt of formats) {
      await resetHarness();
      const res = {
        format: fmt.ext,
        name: fmt.name,
        pass: false,
        loaded: false,
        domFontSelector: null,
        saved: false,
        savedBytes: 0,
        ooxmlVerification: null,
        screenshot: null,
      };

      console.log(`[font-cdp] Testing 0-byte ${fmt.name}...`);
      const t0 = Date.now();
      await evaluate(`(() => {
        const emptyBytes = new Uint8Array(0);
        window.__harness.openBuffer(${JSON.stringify(fmt.name)}, emptyBytes.buffer, "edit");
      })()`);

      const openLog = await waitForLog(/loaded:|error /);
      res.loaded = /loaded:/.test(openLog.text) && !/error /.test(openLog.text);
      res.openMs = Date.now() - t0;
      if (!res.loaded) {
        res.error = "Document failed to load: " + openLog.text;
        report.formats[fmt.ext] = res;
        report.summary.failed++;
        continue;
      }

      await sleep(1800);

      // Locate inner editor execution context
      let editorContextId = null;
      for (let i = contexts.length - 1; i >= 0; i--) {
        const ctx = contexts[i];
        if (ctx.sessionId) continue;
        try {
          const u = await evaluate(`location.href`, ctx.id);
          if (u && u.includes("web-apps/apps/")) {
            editorContextId = ctx.id;
            break;
          }
        } catch {}
      }

      if (!editorContextId) {
        res.error = "Could not locate inner editor execution context";
        report.formats[fmt.ext] = res;
        report.summary.failed++;
        continue;
      }

      // If PPTX, double click to activate textbox
      if (fmt.requiresFocus && fmt.focusPos) {
        await send("Input.dispatchMouseEvent", { type: "mousePressed", x: fmt.focusPos.x, y: fmt.focusPos.y, button: "left", clickCount: 1 });
        await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: fmt.focusPos.x, y: fmt.focusPos.y, button: "left", clickCount: 1 });
        await sleep(150);
        await send("Input.dispatchMouseEvent", { type: "mousePressed", x: fmt.focusPos.x, y: fmt.focusPos.y, button: "left", clickCount: 2 });
        await send("Input.dispatchMouseEvent", { type: "mouseReleased", x: fmt.focusPos.x, y: fmt.focusPos.y, button: "left", clickCount: 2 });
        await sleep(1000);
      }

      // Inspect inner editor DOM font selector
      const domInspection = await evaluate(`(() => {
        const slotInput = document.querySelector('#slot-field-fontname input');
        const comboInput = document.querySelector('.combobox.fonts input');
        const ariaInput = document.querySelector('input[aria-label="Phông chữ"], input[aria-label="Font"]');
        const target = slotInput || comboInput || ariaInput;
        if (!target) return null;
        return {
          selector: slotInput ? '#slot-field-fontname input' : (comboInput ? '.combobox.fonts input' : 'input[aria-label="Phông chữ"]'),
          tagName: target.tagName,
          id: target.id || null,
          className: target.className || null,
          ariaLabel: target.getAttribute('aria-label'),
          value: target.value,
          disabled: target.disabled,
        };
      })()`, editorContextId);

      res.domFontSelector = domInspection;
      const fontValue = domInspection?.value || "";
      const fontIsArial = fontValue.trim().toLowerCase() === "arial";
      const fontHasDengXian = /dengxian|等线/i.test(fontValue);

      // Capture screenshot
      const shotPath = await captureScreenshot(`office-e2-font-${fmt.ext}.png`);
      res.screenshot = shotPath;

      // Trigger Save
      const tSave = Date.now();
      await evaluate(`document.getElementById('save').click()`);
      const saveLog = await waitForLog(/saved:|error /, 45000);
      res.saved = /saved:/.test(saveLog.text) && !/error /.test(saveLog.text);
      res.saveMs = Date.now() - tSave;

      if (!res.saved) {
        res.error = "Document failed to save: " + saveLog.text;
        report.formats[fmt.ext] = res;
        report.summary.failed++;
        continue;
      }

      // Extract saved bytes
      const b64 = await evaluate(`window.__harness.getLastSavedBase64()`);
      if (!b64) {
        res.error = "getLastSavedBase64 returned null";
        report.formats[fmt.ext] = res;
        report.summary.failed++;
        continue;
      }

      const buffer = Buffer.from(b64, "base64");
      res.savedBytes = buffer.byteLength;

      // Unzip & OOXML Verification
      const tmpZip = join(tmpdir(), `xerp-font-cdp-${fmt.ext}-${Date.now()}.zip`);
      writeFileSync(tmpZip, buffer);

      try {
        const fullXml = execFileSync("unzip", ["-p", tmpZip], {
          maxBuffer: 32 * 1024 * 1024,
          encoding: "utf8",
        });

        const ooxmlHasArial = fullXml.includes("Arial");
        const ooxmlHasDengXian = fullXml.includes("DengXian") || fullXml.includes("等线");

        const xmlSnippets = {};
        for (const kf of fmt.keyXmlFiles) {
          try {
            const xmlContent = execFileSync("unzip", ["-p", tmpZip, kf], { encoding: "utf8" });
            const snippets = [];
            for (const chunk of xmlContent.split(/(?=<[^>]+>)/)) {
              if (/font|typeface|Arial/i.test(chunk)) {
                snippets.push(chunk.trim().slice(0, 120));
              }
            }
            xmlSnippets[kf] = snippets.slice(0, 4);
          } catch {}
        }

        res.ooxmlVerification = {
          pass: ooxmlHasArial && !ooxmlHasDengXian,
          hasArial: ooxmlHasArial,
          hasDengXian: ooxmlHasDengXian,
          keyXmlSnippets: xmlSnippets,
        };

        res.pass = res.loaded && fontIsArial && !fontHasDengXian && res.saved && res.ooxmlVerification.pass;
      } catch (err) {
        res.ooxmlVerification = { pass: false, error: err.message };
        res.pass = false;
      } finally {
        try { rmSync(tmpZip, { force: true }); } catch {}
      }

      report.formats[fmt.ext] = res;
      if (res.pass) {
        report.summary.passed++;
        console.log(`[font-cdp] PASS: ${fmt.name} (font=${fontValue}, bytes=${res.savedBytes}, ooxmlArial=${res.ooxmlVerification?.hasArial})`);
      } else {
        report.summary.failed++;
        console.error(`[font-cdp] FAIL: ${fmt.name}`, res.error || "Verification failed");
      }
    }

    report.passAll = report.summary.passed === report.summary.total;

    // Write reports
    const outJsonPath = join(OUT_DIR, "font-selector-report.json");
    const scratchJsonPath = join(SCRATCH_DIR, "office-e2-font-selector-verification.json");
    const reportText = JSON.stringify(report, null, 2);
    writeFileSync(outJsonPath, reportText);
    writeFileSync(scratchJsonPath, reportText);

    console.log(`\n[font-cdp] Verification finished. ${report.summary.passed}/${report.summary.total} passed.`);
    console.log(`Report saved to: ${scratchJsonPath}`);

    return report;
  } finally {
    chrome.kill("SIGKILL");
    try { rmSync(profile, { recursive: true, force: true }); } catch {}
  }
}

main()
  .then((rep) => {
    if (!rep?.passAll) process.exit(1);
  })
  .catch((err) => {
    console.error("[font-cdp] Fatal error:", err);
    process.exit(1);
  });
