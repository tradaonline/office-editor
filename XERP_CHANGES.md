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
| 2026-10-20 | `app/xerp`: embed route for XERP. The parent page sends the file bytes by `postMessage` from an allowed origin (`NEXT_PUBLIC_XERP_PARENT_ORIGINS`); "save" returns the converted file to the parent instead of downloading it. Plugins and macros off. |
| 2026-10-20 | `utils/editor/server.ts`, `types.ts`: `onExport` hook and `setUser`. |
| 2026-10-20 | `app/layout.tsx`: Google Analytics removed. |
| 2026-10-20 | `Dockerfile`, `Caddyfile`: build arg for allowed parents, CSP (`frame-ancestors`, `connect-src 'self'`), upstream public pages not served, sample templates and sdkjs-plugins not shipped, pnpm pinned to 10, x2t served plain with precompressed br/gz (brotli-only failed over HTTP), on-the-fly compression for the editor assets; `xerp/harness` test page on :81 (test box only). |
