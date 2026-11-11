# office-wasm-sdk

Host-side part of office-wasm. The viewer itself (AGPL-3.0-or-later) is in the sibling repo `office-wasm-viewer`; this repo talks to it only through
`postMessage` and contains none of its code.

| Path | What |
|---|---|
| `sdk/office-viewer.mjs` | `OfficeViewer`: embeds the viewer iframe; `open`, `exportPdf`, `save`, `convert`, local fonts (`enableLocalFonts`) |
| `src/extract/` | pure-JS text/structure extraction for docx, xlsx, pptx (no wasm) |
| `demo/` | demo host page + `server.cjs` (starts the viewer's `tools/serve.cjs` on port+1) + `e2e.cjs` (Playwright) |
| `samples/` | test documents (`samples/local/` is not versioned: real documents) |
| `tools/` | `extract-bench.mjs`, `wmf-fonts.mjs` |

```
npm install
npm run serve          # http://127.0.0.1:8765/ (viewer on 8766); needs ../office-wasm-viewer with its artifacts built
npm run e2e            # embedded + desktop; -- --modes edit | --local-fonts; npm run e2e:coi
```

The viewer folder is `../office-wasm-viewer` or the path in `OFFICE_WASM_VIEWER`. Licence: not decided yet (`UNLICENSED`).
