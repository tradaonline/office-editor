import assert from "node:assert";
import { test } from "node:test";
import { isMobileViewport, shouldHideLeftMenu } from "../utils/editor/utils";

test("shouldHideLeftMenu hides rail on mobile view only", () => {
  // Mobile view: should hide
  assert.strictEqual(shouldHideLeftMenu(false, 375, "Mozilla/5.0 (iPhone)"), true);
  assert.strictEqual(shouldHideLeftMenu(false, 412, "Mozilla/5.0 (Linux; Android)"), true);
  assert.strictEqual(shouldHideLeftMenu(false, 767), true);

  // Mobile edit: unchanged (must not hide)
  assert.strictEqual(shouldHideLeftMenu(true, 375, "Mozilla/5.0 (iPhone)"), false);
  assert.strictEqual(shouldHideLeftMenu(true, 412, "Mozilla/5.0 (Linux; Android)"), false);

  // Desktop view: unchanged (must not hide)
  assert.strictEqual(shouldHideLeftMenu(false, 1440, "Mozilla/5.0 (Macintosh)"), false);
  assert.strictEqual(shouldHideLeftMenu(false, 1024, "Mozilla/5.0 (Windows NT)"), false);
  assert.strictEqual(shouldHideLeftMenu(false, 768), false);

  // Desktop edit: unchanged (must not hide)
  assert.strictEqual(shouldHideLeftMenu(true, 1440, "Mozilla/5.0 (Macintosh)"), false);
});

test("isMobileViewport detects width < 768 and mobile user agents", () => {
  // Desktop
  assert.strictEqual(isMobileViewport(1440, "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"), false);
  assert.strictEqual(isMobileViewport(1024, "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"), false);
  assert.strictEqual(isMobileViewport(768, "Mozilla/5.0 (Macintosh)"), false);

  // Mobile widths
  assert.strictEqual(isMobileViewport(375, "Mozilla/5.0 (Macintosh)"), true);
  assert.strictEqual(isMobileViewport(412, "Mozilla/5.0 (Macintosh)"), true);
  assert.strictEqual(isMobileViewport(767, "Mozilla/5.0 (Macintosh)"), true);

  // Mobile User Agents
  assert.strictEqual(isMobileViewport(undefined, "Mozilla/5.0 (Linux; Android 14; Pixel 8) Mobile Safari/537.36"), true);
  assert.strictEqual(isMobileViewport(undefined, "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X)"), true);
});
