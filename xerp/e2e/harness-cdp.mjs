// E2 CDP harness for XERP Office Editor.
// Drives xerp/harness in headless Chrome, testing:
// - OE-T10: Network requests (editor targets restricted strictly to editor origin, blob:, data:)
// - OE-T11: Disallowed frame embedding & foreign postMessage rejection
// - OE-T13: Vietnamese typing, save, reopen across docx, xlsx, pptx
// - OE-T15: Public endpoints (/, /editor, /editor404) return 404; about/license notice
// - OE-T16: Legacy .doc opens in view/edit, modern export (.docx) on save
// - xlsx Enter cell commit verification & OOXML unzip XML verification
// - 404 gate for missing editor resources (except harness favicon)
// - View vs Edit matrix across all formats
//
// Usage:
//   node xerp/e2e/harness-cdp.mjs
//   HARNESS=http://192.168.88.99:8091/ EDITOR=http://192.168.88.99:8090 SUITE=matrix node xerp/e2e/harness-cdp.mjs

import { execFileSync, spawn } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(__dirname, "../..");

// CLI args parsing
const args = process.argv.slice(2);
const getArg = (name, fallback) => {
  const prefix = `--${name}=`;
  const hit = args.find((a) => a.startsWith(prefix));
  return hit ? hit.slice(prefix.length) : (process.env[name.toUpperCase()] || fallback);
};

const HARNESS = getArg("harness", "http://192.168.88.99:8091/");
const EDITOR = getArg("editor", "http://192.168.88.99:8090").replace(/\/$/, "");
const SUITE = getArg("suite", "all").toLowerCase();
const OUT_DIR = resolve(getArg("out", join(__dirname, "out")));
const TIMEOUT_MS = Number(getArg("timeout", "90000"));
const CDP_CMD_TIMEOUT_MS = 25000;

// Pass EDITOR query on every harness navigation per review finding
function getHarnessUrl() {
  const u = new URL(HARNESS);
  u.searchParams.set("editor", EDITOR);
  return u.toString();
}

mkdirSync(OUT_DIR, { recursive: true });

// Chrome binary discovery
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

// Fixture discovery - require explicit SAMPLE_DOC (no private Downloads fallback, no RTF counterfeit)
function discoverFixtures() {
  return {
    docx: {
      name: "Essay Outline.docx",
      path: process.env.SAMPLE_DOCX || join(REPO_ROOT, "public/files/essay_templates_Essay Outline Doc.docx"),
      url: "https://raw.githubusercontent.com/baotlake/office-website/main/public/files/essay_templates_Essay%20Outline%20Doc.docx",
      vietnameseText: "Cộng hòa Xã hội Chủ nghĩa Việt Nam - Độc lập Tự do Hạnh phúc. Thử nghiệm XERP E2 gõ tiếng Việt có dấu: á à ả ã ạ đ ê ô ơ ư. Hợp đồng kinh tế số 108/2026/HĐ-XERP.",
      clickPos: { x: 720, y: 450 },
    },
    xlsx: {
      name: "swot-analysis.xlsx",
      path: process.env.SAMPLE_XLSX || join(REPO_ROOT, "public/files/swot-analysis.xlsx"),
      url: "https://raw.githubusercontent.com/baotlake/office-website/main/public/files/swot-analysis.xlsx",
      vietnameseText: "Báo cáo tài chính XERP 2026: Doanh thu quý 4 đạt 15.000.000.000 đ (tiếng Việt: á à ả ã ạ đ ê ô ơ ư). Đã duyệt.",
      clickPos: { x: 300, y: 300 },
    },
    pptx: {
      name: "Daily Workflow Infographic.pptx",
      path: process.env.SAMPLE_PPTX || join(REPO_ROOT, "public/files/Daily Workflow Infographic.pptx"),
      url: "https://raw.githubusercontent.com/baotlake/office-website/main/public/files/Daily%20Workflow%20Infographic.pptx",
      vietnameseText: "Chiến lược sản phẩm XERP 2026: Tích hợp trình soạn thảo Office nội bộ an toàn (tiếng Việt: á à ả ã ạ đ ê ô ơ ư).",
      clickPos: { x: 600, y: 400 },
    },
    doc: {
      name: "legacy-sample.doc",
      path: process.env.SAMPLE_DOC || null,
      url: null,
      vietnameseText: "Văn bản hành chính XERP (.doc) chuyển đổi sang định dạng hiện đại (.docx) có dấu tiếng Việt đầy đủ.",
      clickPos: { x: 720, y: 450 },
    },
  };
}

async function loadFixtureBuffer(fixture) {
  if (fixture.path && existsSync(fixture.path)) {
    return { buffer: readFileSync(fixture.path), source: `file:${fixture.path}` };
  }
  if (fixture.url) {
    const res = await fetch(fixture.url);
    if (!res.ok) throw new Error(`HTTP ${res.status} fetching fixture from ${res.url}`);
    const ab = await res.arrayBuffer();
    return { buffer: Buffer.from(ab), source: `url:${fixture.url}` };
  }
  if (fixture.name?.endsWith(".doc")) {
    throw new Error(
      "SAMPLE_DOC environment variable required for .doc test (no private Downloads fallback; RTF payload is not valid .doc)"
    );
  }
  throw new Error(`Fixture not found locally (${fixture.path}) and no fallback URL configured.`);
}

// OOXML verification: extract saved archive and check XML text directly
function verifyOoxmlContainsText(buffer, expectedText) {
  const tmpZip = join(tmpdir(), `xerp-ooxml-${Date.now()}-${Math.random().toString(36).slice(2)}.zip`);
  try {
    writeFileSync(tmpZip, buffer);
    // Use unzip from PATH per harness review finding
    const stdout = execFileSync("unzip", ["-p", tmpZip], {
      maxBuffer: 32 * 1024 * 1024,
      encoding: "utf8",
    });

    // Extract text preserving runs: block boundaries get spaces, inline run tags get removed
    const textPreservingRuns = stdout
      .replace(/<\/(?:[a-zA-Z0-9_]+:)?(p|row|si|tc|tr|sp)>/gi, " ")
      .replace(/<(?:[a-zA-Z0-9_]+:)?(tab|br)\b[^>]*\/?>/gi, " ")
      .replace(/<[^>]+>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));

    const collapsed = textPreservingRuns.replace(/\s+/g, " ").trim();
    const normExpected = expectedText.replace(/\s+/g, " ").trim();

    const exactMatch = stdout.includes(normExpected) || collapsed.includes(normExpected);
    const segment = normExpected.split(/[:. -]/).map((s) => s.trim()).filter((s) => s.length >= 10)[0] || normExpected.slice(0, 20);
    const segmentMatch = collapsed.includes(segment);
    const insertedTextProven = exactMatch || segmentMatch;

    return {
      pass: insertedTextProven,
      insertedTextProven,
      exactMatch,
      segmentMatch,
      matchedSegment: segment,
      bytes: buffer.byteLength,
    };
  } catch (err) {
    return { pass: false, insertedTextProven: false, error: `unzip verification error: ${err.message}` };
  } finally {
    try { rmSync(tmpZip, { force: true }); } catch {}
  }
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const report = {
    timestamp: new Date().toISOString(),
    harness: HARNESS,
    editor: EDITOR,
    suite: SUITE,
    passAll: false,
    summary: { total: 0, passed: 0, failed: 0 },
    checks: {},
    screenshots: [],
    consoleErrors: [],
  };

  const port = 9300 + Math.floor(Math.random() * 500);
  const profile = mkdtempSync(join(tmpdir(), "xerp-cdp-"));
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
  const networkRequests = [];
  const networkResponses = [];
  const networkFailures = [];
  const logEntries = [];
  const topTestOrigins = new Set();
  const targetSessions = new Map();

  function send(method, params = {}, sessionId, timeoutMs = CDP_CMD_TIMEOUT_MS) {
    const id = ++msgId;
    const msg = { id, method, params };
    if (sessionId) msg.sessionId = sessionId;
    ws.send(JSON.stringify(msg));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error(`CDP command timed out after ${timeoutMs}ms: ${method}`));
        }
      }, timeoutMs);
      pending.set(id, {
        resolve: (val) => { clearTimeout(timer); resolve(val); },
        reject: (err) => { clearTimeout(timer); reject(err); },
      });
    });
  }

  const evaluate = async (expression, sessionId) => {
    const r = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true }, sessionId);
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
      report.screenshots.push({ name, path: filePath });
      return filePath;
    } catch {
      return null;
    }
  };

  const clickAt = async (x, y) => {
    await send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
    await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
  };

  const doubleClickAt = async (x, y) => {
    await send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 });
    await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 });
    await sleep(60);
    await send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 2 });
    await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 2 });
  };

  const typeText = async (text) => {
    await send("Input.insertText", { text });
  };

  // Dispatch standard Enter keydown + keyup to commit xlsx active cell into spreadsheet data model
  const pressEnter = async () => {
    await send("Input.dispatchKeyEvent", {
      type: "rawKeyDown",
      key: "Enter",
      code: "Enter",
      windowsVirtualKeyCode: 13,
      nativeVirtualKeyCode: 13,
      macCharCode: 13,
      text: "\r",
      unmodifiedText: "\r",
    });
    await send("Input.dispatchKeyEvent", {
      type: "keyUp",
      key: "Enter",
      code: "Enter",
      windowsVirtualKeyCode: 13,
      nativeVirtualKeyCode: 13,
      macCharCode: 13,
    });
  };

  const waitForLog = async (pattern, timeoutMs = TIMEOUT_MS) => {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      const text = await evaluate(`document.getElementById('log')?.innerText || ''`).catch(() => "");
      if (pattern.test(text)) return { ok: true, text };
      await sleep(400);
    }
    const finalLog = await evaluate(`document.getElementById('log')?.innerText || ''`).catch(() => "");
    return { ok: false, text: finalLog };
  };

  const resetHarness = async () => {
    await evaluate(`(() => {
      if (window.__harness?.reset) {
        window.__harness.reset();
      } else {
        const stage = document.getElementById('stage');
        if (stage) stage.innerHTML = '<div id="empty">Chọn một tệp rồi bấm Mở.</div>';
        const saveBtn = document.getElementById('save');
        if (saveBtn) saveBtn.disabled = true;
        const reopenBtn = document.getElementById('reopen');
        if (reopenBtn) reopenBtn.disabled = true;
        const dirty = document.getElementById('dirty');
        if (dirty) dirty.textContent = '';
      }
      const log = document.getElementById('log');
      if (log) log.innerHTML = '';
      if (window.__harnessMessages) window.__harnessMessages = [];
      window.__lastSaved = null;
    })()`).catch(() => {});
    await sleep(500);
  };

  const recordCheck = (key, res) => {
    report.checks[key] = res;
    report.summary.total++;
    if (res.pass) report.summary.passed++;
    else report.summary.failed++;
  };

  try {
    // 1. Connect Chrome CDP
    let targets = [];
    for (let i = 0; i < 50; i++) {
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
        if (m.error) reject(new Error(`CDP error: ${m.error.message || JSON.stringify(m.error)}`));
        else resolve(m.result || {});
        return;
      }

      const p = m.params || {};
      if (m.method === "Runtime.exceptionThrown") {
        const details = p.exceptionDetails || {};
        const text = details.exception?.description || details.text || "Unknown runtime exception";
        report.consoleErrors.push({
          target: m.sessionId ? "frame" : "top",
          type: "exception",
          text: text.slice(0, 500),
          url: details.url,
          lineNumber: details.lineNumber,
        });
      } else if (m.method === "Runtime.consoleAPICalled") {
        const text = (p.args || []).map((a) => a.value ?? a.description ?? a.type).join(" ");
        if (p.type === "error" || p.type === "warning" || /xerp|x2t|Error|error/i.test(text)) {
          report.consoleErrors.push({
            target: m.sessionId ? "frame" : "top",
            type: p.type,
            text: text.slice(0, 300),
          });
        }
      } else if (m.method === "Log.entryAdded") {
        const entry = p.entry || {};
        logEntries.push({
          source: entry.source,
          level: entry.level,
          text: entry.text || "",
          timestamp: entry.timestamp,
        });
      } else if (m.method === "Network.requestWillBeSent") {
        if (p.request?.url) {
          networkRequests.push({
            url: p.request.url,
            method: p.request.method,
            type: p.type,
            sessionId: m.sessionId,
          });
        }
      } else if (m.method === "Network.responseReceived") {
        if (p.response?.url) {
          networkResponses.push({
            url: p.response.url,
            status: p.response.status,
            statusText: p.response.statusText,
            sessionId: m.sessionId,
          });
        }
      } else if (m.method === "Network.loadingFailed") {
        networkFailures.push({
          requestId: p.requestId,
          errorText: p.errorText,
          canceled: p.canceled,
          blockedReason: p.blockedReason,
          type: p.type,
          sessionId: m.sessionId,
        });
      } else if (m.method === "Target.attachedToTarget") {
        const sid = p.sessionId;
        targetSessions.set(sid, p.targetInfo);
        for (const method of ["Runtime.enable", "Log.enable", "Network.enable"]) {
          send(method, {}, sid).catch(() => {});
        }
        send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true }, sid).catch(() => {});
        send("Runtime.runIfWaitingForDebugger", {}, sid).catch(() => {});
      } else if (m.method === "Target.detachedFromTarget") {
        targetSessions.delete(p.sessionId);
      }
    });

    await send("Runtime.enable");
    await send("Log.enable");
    await send("Page.enable");
    await send("Network.enable");
    await send("Target.setAutoAttach", { autoAttach: true, waitForDebuggerOnStart: false, flatten: true });

    // --- CHECK: OE-T15 Public endpoints 404 & About Notice ---
    if (SUITE === "all" || SUITE === "t15") {
      const check404 = { pass: true, endpoints: {} };
      for (const path of ["/", "/editor", "/editor404"]) {
        try {
          const res = await fetch(`${EDITOR}${path}`, { redirect: "manual" });
          check404.endpoints[path] = { status: res.status, ok: res.status === 404 };
          if (res.status !== 404) check404.pass = false;
        } catch (err) {
          check404.endpoints[path] = { error: err.message, ok: false };
          check404.pass = false;
        }
      }
      recordCheck("oe_t15_public_404", check404);

      let aboutRes = { pass: false };
      try {
        const res = await fetch(`${EDITOR}/xerp/about`);
        const text = await res.text();
        const hasOnlyOffice = /ONLYOFFICE/i.test(text);
        const hasAscensio = /Ascensio/i.test(text);
        const hasAgpl = /AGPL/i.test(text);
        const hasSourceUrl = /https?:\/\/[^\s<"']+/i.test(text);
        const hasPendingNotice = /Đang cập nhật/i.test(text);
        const hasSourceOrNotice = hasSourceUrl || hasPendingNotice;
        aboutRes = {
          pass: res.status === 200 && hasOnlyOffice && hasAscensio && hasAgpl && hasSourceOrNotice,
          status: res.status,
          hasOnlyOffice,
          hasAscensio,
          hasAgpl,
          hasSourceUrl,
          hasPendingNotice,
          hasSourceOrNotice,
          snippet: text.slice(0, 200),
        };
      } catch (err) {
        aboutRes = { pass: false, error: err.message };
      }
      recordCheck("oe_t15_about_license", aboutRes);
    }

    const fixtures = discoverFixtures();

    // --- CHECK: OE-T11 Real Unallowed-Origin Embedding CSP & Foreign postMessage ---
    if (SUITE === "all" || SUITE === "t11") {
      // Real unallowed-origin embedding test and foreign-origin postMessage probe server
      const unallowedEmbedCheck = { pass: false };
      let unallowedServer = null;
      let unallowedOrigin = null;
      try {
        const serverPort = 9800 + Math.floor(Math.random() * 150);
        await new Promise((res, rej) => {
          unallowedServer = createServer((req, resHttp) => {
            resHttp.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
            if (req.url === "/foreign-probe.html") {
              resHttp.end(`<!doctype html><html><body><script>
window.onload = () => {
  for (let i = 0; i < window.parent.length; i++) {
    try {
      window.parent[i].postMessage({ type: "xerp-office:save", requestId: "foreign-origin-probe" }, "*");
    } catch (e) {}
  }
};
</script></body></html>`);
            } else {
              resHttp.end(`<!doctype html><html><body><iframe id="unallowed_frame" src="${EDITOR}/xerp" width="600" height="400"></iframe></body></html>`);
            }
          });
          unallowedServer.listen(serverPort, "127.0.0.1", () => res());
          unallowedServer.on("error", rej);
        });

        unallowedOrigin = `http://127.0.0.1:${serverPort}`;
        topTestOrigins.add(unallowedOrigin);
        const prevConsoleCount = report.consoleErrors.length;
        const prevLogCount = logEntries.length;
        const prevFailureCount = networkFailures.length;

        await send("Page.navigate", { url: unallowedOrigin });
        await sleep(2500);

        const newConsoleErrors = report.consoleErrors.slice(prevConsoleCount);
        const newLogs = logEntries.slice(prevLogCount);
        const newFailures = networkFailures.slice(prevFailureCount);

        const hasCspLogViolation = newLogs.some(
          (e) => /frame-ancestors|Refused to frame|Content Security Policy/i.test(e.text)
        );
        const hasCspConsole = newConsoleErrors.some(
          (e) => /frame-ancestors|Refused to frame|Content Security Policy/i.test(e.text)
        );
        const hasBlockedFrameNav = newFailures.some(
          (f) =>
            f.blockedReason === "csp" ||
            f.blockedReason === "subresource-filter" ||
            /ERR_BLOCKED_BY_RESPONSE|blocked/i.test(f.errorText || "")
        );
        const hasCspViolation = hasCspLogViolation || hasCspConsole || hasBlockedFrameNav;

        // Verify CSP response header contains frame-ancestors
        const probeRes = await fetch(`${EDITOR}/xerp`);
        const cspHeader = probeRes.headers.get("content-security-policy") || "";
        const hasFrameAncestorsCsp = /frame-ancestors\s+'self'/i.test(cspHeader);

        await captureScreenshot("oe-t11-unallowed-csp.png");

        // Fail-closed without SOP exception
        unallowedEmbedCheck.pass = hasFrameAncestorsCsp && hasCspViolation;
        unallowedEmbedCheck.hasFrameAncestorsCsp = hasFrameAncestorsCsp;
        unallowedEmbedCheck.hasCspViolation = hasCspViolation;
        unallowedEmbedCheck.hasCspLogViolation = hasCspLogViolation;
        unallowedEmbedCheck.hasCspConsole = hasCspConsole;
        unallowedEmbedCheck.hasBlockedFrameNav = hasBlockedFrameNav;
        unallowedEmbedCheck.cspHeader = cspHeader;
      } catch (err) {
        unallowedEmbedCheck.error = err.message;
      }
      recordCheck("oe_t11_unallowed_origin_csp", unallowedEmbedCheck);

      // 2. Foreign-origin and foreign-source postMessage rejection
      await send("Page.navigate", { url: getHarnessUrl() });
      await sleep(1500);

      // Open a real fixture first; require active frame (no skipped pass)
      const probeFixture = fixtures.docx;
      let probeFixtureLoaded = false;
      try {
        const loadedFixture = await loadFixtureBuffer(probeFixture);
        const b64 = loadedFixture.buffer.toString("base64");
        await resetHarness();
        await evaluate(`(async () => {
          const binStr = atob(${JSON.stringify(b64)});
          const bytes = new Uint8Array(binStr.length);
          for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
          if (window.__harness?.openBuffer) {
            window.__harness.openBuffer(${JSON.stringify(probeFixture.name)}, bytes.buffer, "edit");
          }
        })()`);
        const openLog = await waitForLog(/loaded:|error /, TIMEOUT_MS);
        probeFixtureLoaded = /loaded:/.test(openLog.text) && !/error /.test(openLog.text);
      } catch (err) {
        probeFixtureLoaded = false;
      }

      let foreignRes;
      if (!probeFixtureLoaded) {
        foreignRes = {
          pass: false,
          error: "probe_fixture_failed_to_load",
          frameActive: false,
        };
      } else {
        foreignRes = await evaluate(`(async () => {
          const frame = document.getElementById("frame");
          if (!frame || !frame.contentWindow) {
            return { pass: false, error: "frame_not_active", frameActive: false };
          }
          const initialCount = (window.__harnessMessages || []).length;

          // Foreign source probe: sibling iframe (about:blank, same-origin with harness)
          // posts to editor frame. Sender origin is allowed parent, but source is not window.parent.
          const foreignSourceChild = document.createElement("iframe");
          document.body.appendChild(foreignSourceChild);
          try {
            foreignSourceChild.contentWindow.eval(
              'window.parent.document.getElementById("frame").contentWindow.postMessage({ type: "xerp-office:save", requestId: "foreign-source-probe" }, "*")'
            );
          } catch (e) {}

          // Foreign origin probe: real HTTP origin iframe (${unallowedOrigin}) posts to editor frame
          let foreignOriginAttempted = false;
          if (${JSON.stringify(unallowedOrigin)}) {
            const foreignOriginChild = document.createElement("iframe");
            foreignOriginChild.src = "${unallowedOrigin}/foreign-probe.html";
            document.body.appendChild(foreignOriginChild);
            foreignOriginAttempted = true;
          }

          await new Promise((r) => setTimeout(r, 2000));
          try { document.body.removeChild(foreignSourceChild); } catch {}
          try {
            const fo = document.querySelector('iframe[src*="foreign-probe"]');
            if (fo) document.body.removeChild(fo);
          } catch {}

          const afterCount = (window.__harnessMessages || []).length;
          const unauthorizedMessages = (window.__harnessMessages || []).filter(
            (m) => m.requestId === "foreign-source-probe" || m.requestId === "foreign-origin-probe"
          );
          const noUnauthorizedSaved = unauthorizedMessages.length === 0;
          return {
            pass: noUnauthorizedSaved && foreignOriginAttempted,
            frameActive: true,
            foreignOriginAttempted,
            foreignSourceAttempted: true,
            initialCount,
            afterCount,
            unauthorizedCount: unauthorizedMessages.length,
          };
        })()`).catch((e) => ({ pass: false, error: e.message, frameActive: false }));
      }
      if (unallowedServer) {
        try { unallowedServer.close(); } catch {}
      }
      recordCheck("oe_t11_foreign_postmessage", foreignRes);
    }

    // --- CHECK: View / Edit Matrix with Meaningful Vietnamese & Enter commit ---
    const formatsToTest = [];
    if (SUITE === "all" || SUITE === "matrix") {
      formatsToTest.push("docx", "xlsx", "pptx", "doc");
    } else if (["docx", "xlsx", "pptx", "doc"].includes(SUITE)) {
      formatsToTest.push(SUITE);
    }

    // Ensure we are on the harness page
    await send("Page.navigate", { url: getHarnessUrl() });
    await sleep(1500);

    for (const fmt of formatsToTest) {
      const fixture = fixtures[fmt];
      let loadedFixture;
      try {
        loadedFixture = await loadFixtureBuffer(fixture);
      } catch (err) {
        recordCheck(`matrix_${fmt}_view`, { pass: false, error: `Fixture load error: ${err.message}` });
        recordCheck(`matrix_${fmt}_edit`, { pass: false, error: `Fixture load error: ${err.message}` });
        continue;
      }

      const b64 = loadedFixture.buffer.toString("base64");

      // 1. View Mode Test
      await resetHarness();
      const viewResult = { pass: false, format: fmt, mode: "view", fixtureSource: loadedFixture.source };
      try {
        await evaluate(`(async () => {
          const binStr = atob(${JSON.stringify(b64)});
          const bytes = new Uint8Array(binStr.length);
          for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
          if (window.__harness?.openBuffer) {
            window.__harness.openBuffer(${JSON.stringify(fixture.name)}, bytes.buffer, "view");
          } else {
            const dt = new DataTransfer();
            dt.items.add(new File([new Blob([bytes])], ${JSON.stringify(fixture.name)}));
            document.getElementById('file').files = dt.files;
            document.getElementById('mode').value = "view";
            document.getElementById('open').click();
          }
        })()`);

        const openLog = await waitForLog(/loaded:|error /);
        viewResult.loaded = /loaded:/.test(openLog.text) && !/error /.test(openLog.text);
        await sleep(1000);
        viewResult.saveDisabled = await evaluate(`document.getElementById('save')?.disabled === true`);
        await captureScreenshot(`${fmt}-view.png`);
        viewResult.pass = viewResult.loaded && viewResult.saveDisabled;
      } catch (err) {
        viewResult.error = err.message;
      }
      recordCheck(`matrix_${fmt}_view`, viewResult);

      // 2. Edit Mode Test
      await resetHarness();
      const editResult = {
        pass: false,
        format: fmt,
        mode: "edit",
        typing: fixture.vietnameseText,
        fixtureSource: loadedFixture.source,
      };
      try {
        const t0 = Date.now();
        await evaluate(`(async () => {
          const binStr = atob(${JSON.stringify(b64)});
          const bytes = new Uint8Array(binStr.length);
          for (let i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i);
          if (window.__harness?.openBuffer) {
            window.__harness.openBuffer(${JSON.stringify(fixture.name)}, bytes.buffer, "edit");
          } else {
            const dt = new DataTransfer();
            dt.items.add(new File([new Blob([bytes])], ${JSON.stringify(fixture.name)}));
            document.getElementById('file').files = dt.files;
            document.getElementById('mode').value = "edit";
            document.getElementById('open').click();
          }
        })()`);

        const openLog = await waitForLog(/loaded:|error /);
        editResult.loaded = /loaded:/.test(openLog.text) && !/error /.test(openLog.text);
        editResult.openMs = Date.now() - t0;
        await sleep(1500);
        await captureScreenshot(`${fmt}-edit-open.png`);

        if (editResult.loaded) {
          // Focus and type Vietnamese text
          if (fmt === "pptx") {
            await doubleClickAt(fixture.clickPos.x, fixture.clickPos.y);
            await sleep(600);
          } else {
            await clickAt(fixture.clickPos.x, fixture.clickPos.y);
            await sleep(500);
          }
          await typeText(fixture.vietnameseText);
          await sleep(500);

          // For xlsx: MUST press Enter to commit the active cell into spreadsheet data model
          if (fmt === "xlsx") {
            await pressEnter();
            editResult.cellCommittedWithEnter = true;
            await sleep(800);
          }

          // Trigger save
          const tSave = Date.now();
          await evaluate(`document.getElementById('save').click()`);
          const saveLog = await waitForLog(/saved:|error /, 60000);
          editResult.saved = /saved:/.test(saveLog.text) && !/error /.test(saveLog.text);
          editResult.saveMs = Date.now() - tSave;
          await captureScreenshot(`${fmt}-edit-saved.png`);

          // Inspect saved metadata (evaluate scalar in page)
          const lastSaved = await evaluate(`(() => {
            const s = window.__harness?.getLastSaved?.() || window.__lastSaved || null;
            if (!s) return null;
            const byteLength = (s.buffer && typeof s.buffer.byteLength === "number")
              ? s.buffer.byteLength
              : (typeof s.byteLength === "number" ? s.byteLength : 0);
            return {
              name: s.name || "",
              byteLength,
              fileType: s.fileType || null,
              suggestedName: s.suggestedName || null,
              title: s.title || null,
            };
          })()`);
          editResult.savedMetadata = lastSaved;

          // OOXML verification: extract saved bytes from harness and verify via unzip XML
          const savedBase64 = await evaluate(`window.__harness?.getLastSavedBase64?.() || null`);
          if (editResult.saved && savedBase64) {
            const savedBytes = Buffer.from(savedBase64, "base64");
            const ooxmlCheck = verifyOoxmlContainsText(savedBytes, fixture.vietnameseText);
            editResult.ooxmlVerification = ooxmlCheck;
            if (!ooxmlCheck.pass) {
              editResult.pass = false;
              editResult.ooxmlVerificationFailed = true;
            }
          } else if (editResult.saved && !savedBase64) {
            editResult.ooxmlVerification = { pass: false, error: "could_not_extract_saved_bytes" };
            editResult.pass = false;
          }

          // For legacy .doc: check modern export (.docx)
          if (fmt === "doc") {
            const isModern = lastSaved?.fileType === "docx" || lastSaved?.suggestedName?.endsWith(".docx");
            editResult.modernExport = {
              expectedFileType: "docx",
              actualFileType: lastSaved?.fileType,
              suggestedName: lastSaved?.suggestedName,
              pass: !!isModern,
            };
          }

          // Reopen saved document
          if (editResult.saved) {
            const tReopen = Date.now();
            await evaluate(`document.getElementById('reopen').click()`);
            const reopenLog = await waitForLog(/loaded:[\s\S]*loaded:|error /, TIMEOUT_MS);
            editResult.reopened = /loaded:[\s\S]*loaded:/.test(reopenLog.text) && !/error /.test(reopenLog.text);
            editResult.reopenMs = Date.now() - tReopen;
            await sleep(1500);
            await captureScreenshot(`${fmt}-edit-reopened.png`);
          }

          editResult.pass =
            editResult.loaded &&
            editResult.saved &&
            editResult.reopened &&
            Boolean(editResult.ooxmlVerification?.pass && editResult.ooxmlVerification?.insertedTextProven);

          if (fmt === "doc" && editResult.modernExport && !editResult.modernExport.pass) {
            editResult.pass = false;
          }
        }
      } catch (err) {
        editResult.error = err.message;
      }
      recordCheck(`matrix_${fmt}_edit`, editResult);
    }

    // --- CHECK: OE-T10 Network Requests across all frames & workers ---
    if (SUITE === "all" || SUITE === "t10") {
      const editorHost = new URL(EDITOR).host;
      const harnessHost = new URL(HARNESS).host;
      const editorHostsSeen = new Set();
      const harnessHostsSeen = new Set();
      const protocolsSeen = new Set();
      const externalRequests = [];
      const disallowedEditorRequests = [];
      const topTestNavigations = [];

      for (const req of networkRequests) {
        let u;
        try {
          u = new URL(req.url);
          protocolsSeen.add(u.protocol);
        } catch {
          if (req.url.startsWith("blob:")) protocolsSeen.add("blob:");
          else if (req.url.startsWith("data:")) protocolsSeen.add("data:");
          else if (req.url.startsWith("about:")) protocolsSeen.add("about:");
          continue;
        }

        if (u.protocol !== "http:" && u.protocol !== "https:") {
          continue;
        }

        // Classify target: top harness vs editor target (iframe, worker, etc.)
        const isTopHarness = !req.sessionId;

        if (isTopHarness) {
          // Scope top-page test navigation separately (e.g. OE-T11 127.0.0.1 probe), not editor whitelist
          if (topTestOrigins.has(u.origin)) {
            topTestNavigations.push({ url: req.url, host: u.host });
            continue;
          }
          harnessHostsSeen.add(u.host);
          if (u.host !== harnessHost && u.host !== editorHost) {
            externalRequests.push({ url: req.url, host: u.host, target: "top_harness" });
          }
        } else {
          // Editor targets (frames, workers, service workers)
          editorHostsSeen.add(u.host);
          // Ponytail rule: DO NOT whitelist harness requests for editor frames!
          // Editor targets must ONLY request editorHost, never harnessHost or external.
          if (u.host !== editorHost) {
            disallowedEditorRequests.push({
              url: req.url,
              host: u.host,
              target: targetSessions.get(req.sessionId)?.type || "editor_frame_or_worker",
              reason: u.host === harnessHost ? "editor_frame_requested_harness" : "external_host",
            });
          }
        }
      }

      const editorMounted = editorHostsSeen.has(editorHost);
      recordCheck("oe_t10_network", {
        pass: editorMounted && externalRequests.length === 0 && disallowedEditorRequests.length === 0,
        editorMounted,
        editorHostsSeen: Array.from(editorHostsSeen),
        harnessHostsSeen: Array.from(harnessHostsSeen),
        topTestNavigations,
        allProtocols: Array.from(protocolsSeen),
        disallowedEditorRequests,
        externalRequests,
        totalRequestsCaptured: networkRequests.length,
        ...(editorMounted ? {} : { error: "editor_never_mounted" }),
      });
    }

    // --- CHECK: 404 gate for missing editor resources ---
    if (SUITE === "all" || SUITE === "no404") {
      const editorHost = new URL(EDITOR).host;
      const harnessHost = new URL(HARNESS).host;
      const missingEditorResources = [];
      const expectedT15Urls = new Set([
        `${EDITOR}/`,
        `${EDITOR}/editor`,
        `${EDITOR}/editor404`,
        `${EDITOR}/index.html`,
        `${EDITOR}/editor.html`,
      ]);

      for (const res of networkResponses) {
        if (res.status === 404) {
          // Harness favicon is the ONLY allowed 404
          const isHarnessFavicon = res.url.includes("favicon.ico") && res.url.includes(harnessHost);
          // Also skip the explicit OE-T15 negative probes (/, /editor, /editor404)
          const isExpectedProbe = expectedT15Urls.has(res.url) || expectedT15Urls.has(res.url.replace(/\/$/, ""));
          if (!isHarnessFavicon && !isExpectedProbe) {
            missingEditorResources.push({
              url: res.url,
              status: res.status,
              statusText: res.statusText,
              sessionId: res.sessionId,
            });
          }
        }
      }

      const editorMounted = networkRequests.some((r) => r.sessionId && r.url.includes(editorHost));
      recordCheck("oe_editor_no_missing_404_resources", {
        pass: editorMounted && missingEditorResources.length === 0,
        editorMounted,
        missingCount: missingEditorResources.length,
        missingResources: missingEditorResources,
        ...(editorMounted ? {} : { error: "editor_never_mounted" }),
      });
    }

    report.passAll = report.summary.total > 0 && report.summary.failed === 0;
  } catch (fatalErr) {
    report.fatalError = fatalErr.message;
    report.passAll = false;
  } finally {
    if (ws) {
      try { ws.close(); } catch {}
    }
    chrome.kill();
    try { rmSync(profile, { recursive: true, force: true }); } catch {}
  }

  // Fail-closed machine readable JSON stdout
  console.log(JSON.stringify(report, null, 2));
  process.exit(report.passAll ? 0 : 1);
}

main();