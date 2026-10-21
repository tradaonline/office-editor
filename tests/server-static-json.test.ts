import assert from "node:assert";
import { test } from "node:test";
import {
  EditorServer,
  STATIC_JSON_ROUTES,
  FORMULA_FALLBACK_ROUTES,
} from "../utils/editor/server";
import { APP_ROOT } from "../utils/editor/utils";

test("STATIC_JSON_ROUTES maps all bare and runtime static resources to exact proven paths", () => {
  const expectedRoutes: Record<string, string> = {
    // Bare filenames
    "/themes.json": "/web-apps/apps/common/main/resources/themes/themes.json",
    "/alphabetletters.json": "/web-apps/apps/common/main/resources/alphabetletters/alphabetletters.json",
    "/qwertyletters.json": "/web-apps/apps/common/main/resources/alphabetletters/qwertyletters.json",
    "/numbering-lists.json": "/web-apps/apps/documenteditor/main/resources/numbering/numbering-lists.json",
    "/multilevel-lists.json": "/web-apps/apps/documenteditor/main/resources/numbering/multilevel-lists.json",

    // Runtime request paths from ONLYOFFICE editors
    "/common/main/resources/alphabetletters/alphabetletters.json": "/web-apps/apps/common/main/resources/alphabetletters/alphabetletters.json",
    "/common/main/resources/alphabetletters/qwertyletters.json": "/web-apps/apps/common/main/resources/alphabetletters/qwertyletters.json",
    "/resources/numbering/numbering-lists.json": "/web-apps/apps/documenteditor/main/resources/numbering/numbering-lists.json",
    "/resources/numbering/multilevel-lists.json": "/web-apps/apps/documenteditor/main/resources/numbering/multilevel-lists.json",
    "/resources/formula-lang/vi.json": "/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/vi.json",
    "/resources/formula-lang/vi_desc.json": "/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/vi_desc.json",
    "/resources/formula-lang/en_desc.json": "/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/en_desc.json",
    "/resources/formula-lang/en.json": "/web-apps/apps/spreadsheeteditor/main/resources/formula-lang/en.json",
  };

  assert.deepStrictEqual(STATIC_JSON_ROUTES, expectedRoutes);
});

test("FORMULA_FALLBACK_ROUTES maps Vietnamese formula files to exact English fallbacks", () => {
  const expectedFallbacks: Record<string, string> = {
    "/resources/formula-lang/vi.json": "/resources/formula-lang/en.json",
    "/resources/formula-lang/vi_desc.json": "/resources/formula-lang/en_desc.json",
  };

  assert.deepStrictEqual(FORMULA_FALLBACK_ROUTES, expectedFallbacks);
});

test("server.handleRequest routes static JSON files under APP_ROOT without 404", async () => {
  const server = new EditorServer();

  // Mock global fetch to verify target URL
  const originalFetch = globalThis.fetch;
  const fetchedUrls: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = typeof input === "string" ? input : input.toString();
    fetchedUrls.push(url);
    return new Response(JSON.stringify({ ok: true }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }) as typeof fetch;

  try {
    for (const [route, targetSubpath] of Object.entries(STATIC_JSON_ROUTES)) {
      const req = new Request(`http://localhost:8090${route}`);
      const res = await server.handleRequest(req);
      assert.ok(res, `Route ${route} must be handled by handleRequest`);
      assert.strictEqual(res.status, 200);
      const expectedUrl = `http://localhost:8090${APP_ROOT}${targetSubpath}`;
      assert.strictEqual(fetchedUrls[fetchedUrls.length - 1], expectedUrl);
    }
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("server.handleRequest preserves genuine upstream Vietnamese formula translations when present", async () => {
  const server = new EditorServer();

  const originalFetch = globalThis.fetch;
  const fetchedUrls: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = typeof input === "string" ? input : input.toString();
    fetchedUrls.push(url);
    return new Response(JSON.stringify({ lang: "vi", real: true }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  }) as typeof fetch;

  try {
    const req = new Request("http://localhost:8090/resources/formula-lang/vi.json");
    const res = await server.handleRequest(req);
    assert.ok(res);
    assert.strictEqual(res.status, 200);
    const data = await res.json();
    assert.strictEqual(data.lang, "vi");
    assert.strictEqual(fetchedUrls.length, 1);
    assert.ok(fetchedUrls[0].endsWith("/formula-lang/vi.json"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("server.handleRequest falls back to upstream en.json when vi.json is 404 upstream", async () => {
  const server = new EditorServer();

  const originalFetch = globalThis.fetch;
  const fetchedUrls: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = typeof input === "string" ? input : input.toString();
    fetchedUrls.push(url);
    if (url.endsWith("/formula-lang/vi.json")) {
      return new Response("Not Found", { status: 404, statusText: "Not Found" });
    }
    if (url.endsWith("/formula-lang/en.json")) {
      return new Response(JSON.stringify({ lang: "en", fallback: true }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }
    return new Response("Not Found", { status: 404 });
  }) as typeof fetch;

  try {
    const req = new Request("http://localhost:8090/resources/formula-lang/vi.json");
    const res = await server.handleRequest(req);
    assert.ok(res, "Fallback must return a response");
    assert.strictEqual(res.status, 200);
    const data = await res.json();
    assert.strictEqual(data.lang, "en");
    assert.strictEqual(data.fallback, true);
    assert.strictEqual(fetchedUrls.length, 2);
    assert.ok(fetchedUrls[0].endsWith("/formula-lang/vi.json"));
    assert.ok(fetchedUrls[1].endsWith("/formula-lang/en.json"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("server.handleRequest falls back to upstream en_desc.json when vi_desc.json is 404 upstream", async () => {
  const server = new EditorServer();

  const originalFetch = globalThis.fetch;
  const fetchedUrls: string[] = [];
  globalThis.fetch = (async (input: RequestInfo | URL) => {
    const url = typeof input === "string" ? input : input.toString();
    fetchedUrls.push(url);
    if (url.endsWith("/formula-lang/vi_desc.json")) {
      return new Response("Not Found", { status: 404, statusText: "Not Found" });
    }
    if (url.endsWith("/formula-lang/en_desc.json")) {
      return new Response(JSON.stringify({ descLang: "en_desc", fallback: true }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }
    return new Response("Not Found", { status: 404 });
  }) as typeof fetch;

  try {
    const req = new Request("http://localhost:8090/resources/formula-lang/vi_desc.json");
    const res = await server.handleRequest(req);
    assert.ok(res, "Fallback must return a response");
    assert.strictEqual(res.status, 200);
    const data = await res.json();
    assert.strictEqual(data.descLang, "en_desc");
    assert.strictEqual(data.fallback, true);
    assert.strictEqual(fetchedUrls.length, 2);
    assert.ok(fetchedUrls[0].endsWith("/formula-lang/vi_desc.json"));
    assert.ok(fetchedUrls[1].endsWith("/formula-lang/en_desc.json"));
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test("server.handleRequest leaves non-mapped requests to fall through", async () => {
  const server = new EditorServer();
  const res = await server.handleRequest(new Request("http://localhost:8090/unmapped.json"));
  assert.strictEqual(res, null);
});
