# XERP Office — modified version

This is a **modified version** of [ZIZIYI Office](https://github.com/baotlake/office-website)
(AGPL-3.0), which runs the ONLYOFFICE editors (web-apps, sdkjs, x2t) in the browser.
ONLYOFFICE is developed by Ascensio System SIA; this build is based on ONLYOFFICE and is
not affiliated with or endorsed by Ascensio System SIA. The ONLYOFFICE notices, license
texts and attributions shipped with the editors are kept unchanged.

License: GNU AGPL v3 (see `LICENSE.txt`), with the ONLYOFFICE additional terms that come with
its components. The complete source of this modified version is published with it.

## Changes

| Date | Change |
|---|---|
| 2026-10-22 | `app/xerp/page.tsx`, `utils/editor/utils.ts`, `utils/editor/server.ts`, `tests/mobile-layout.test.ts`, `xerp/e2e/mobile-cdp.mjs`: restore non-commercial `branding: false` in EditorServer license handshake (no fabricated licensing capabilities); hide ONLYOFFICE left icon rail search/outline/info for mobile VIEW only (<768px or mobile UA) via existing non-gated LeftMenu controller hide and NotificationCenter `layout:changed` main notification; expand canvas to full viewport; keep mobile edit, desktop view, and desktop edit rails unchanged; preserve visible ONLYOFFICE notices and About link. |
| 2026-10-21 | `utils/editor/empty.ts`, `utils/editor/server.ts`, `tests/font-defaults.test.ts`: replace Chinese 等线 (DengXian) theme and fallback font definitions in empty document templates with standard Arial font scheme; route 0-byte docx/xlsx/pptx files to clean empty templates; preserve imported document formatting without global font overwrites. |
| 2026-10-21 | `utils/editor/server.ts`, `tests/cell-lock-protocol.test.ts`: fix cell lock WebSocket protocol identity to echo user.id + indexUser, matching ONLYOFFICE cell SDK expected identity and eliminating uncaught cell lock indexOf TypeError. |
| 2026-10-21 | `Caddyfile`, `app/xerp/page.tsx`, `xerp/e2e/harness-cdp.mjs`: restore 'unsafe-eval' to script-src CSP (runtime CDP proof: ONLYOFFICE sdk-all-min and app.js template compiler require new Function); drop customization.font which triggered commercial loader license warning modal; add Runtime.exceptionThrown capture to CDP harness. |
| 2026-10-21 | `xerp/e2e/harness-cdp.mjs`, `xerp/harness/index.html`: extend CDP matrix, commit spreadsheet cells before save, inspect saved OOXML Vietnamese text, track frame/worker requests and resource 404s, check forbidden embedding / messages and license page; require explicit SAMPLE_DOC for legacy fixture. |
| 2026-10-21 | `app/xerp/page.tsx`, `utils/editor/utils.ts`, `tests/embed-e2.test.ts`: retain protocol v1, add saved original title / OOXML fileType / suggestedName and close-request; enable legacy-format editing with modern export; disable collaboration/review menus, keep original About, expose XERP license link; set documented Arial 11px editor font. |
| 2026-10-21 | `utils/editor/server.ts`, `tests/server-static-json.test.ts`: route five root JSON requests to exact versioned ONLYOFFICE resource paths under APP_ROOT, preserving unrelated request handling. |
| 2026-10-21 | `Dockerfile`, `Caddyfile`, `Caddyfile.harness`, `app/xerp/about/page.tsx`: add build-time `XERP_SOURCE_URL`, Vietnamese ONLYOFFICE / Ascensio System SIA attribution, AGPL and modified-source links; default `XERP_HARNESS=0` excludes harness content and listener from production; opt-in test site on :81; add script and worker CSP directives (runtime validation pending). |
| 2026-10-20 | `app/xerp`: embed route for XERP. The parent page sends the file bytes by `postMessage` from an allowed origin (`NEXT_PUBLIC_XERP_PARENT_ORIGINS`); "save" returns the converted file to the parent instead of downloading it. Plugins and macros off. |
| 2026-10-20 | `utils/editor/server.ts`, `types.ts`: `onExport` hook and `setUser`. |
| 2026-10-20 | `app/layout.tsx`: Google Analytics removed. |
| 2026-10-20 | `Dockerfile`, `Caddyfile`: build arg for allowed parents, CSP (`frame-ancestors`, `connect-src 'self'`), upstream public pages not served, sample templates and sdkjs-plugins not shipped, pnpm pinned to 10, x2t served plain with precompressed br/gz (brotli-only failed over HTTP), on-the-fly compression for the editor assets, `no-cache` on pages (versioned assets stay immutable via `path_regexp` — the upstream `path /v*-*/*` never matched; bump `HASH` when headers change); `xerp/harness` test page on :81 (test box only). |
