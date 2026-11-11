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

The viewer folder is `../office-wasm-viewer` or the path in `OFFICE_WASM_VIEWER`.

## Licence and the boundary with the AGPL viewer

This repository is **Apache-2.0** (`LICENSE`, `NOTICE`). An application that uses it (for example mschool) does not have to publish its own source.
That holds as long as the application and the viewer stay two separate programs:

- Load the viewer only as an `<iframe>` from its own origin and talk to it through the SDK (`window.postMessage`). Do not import `viewer.mjs`,
  `x2t/`, `fonts/*.mjs` or the wasm modules into your page or bundle: that would make your application part of the AGPL work.
- The viewer (`office-wasm-viewer`) and `metafile-wasm` stay AGPL: whoever runs them for users over a network must offer those users the source of
  the exact version running (set the `office-wasm-viewer:source` meta tag; tag releases) and publish any changes made to them.
- The same applies when x2t or metafile run as a separate process on a server: they are AGPL programs, our builds are modified versions.
- `src/extract` and `sdk/office-viewer.mjs` are ours (Apache-2.0) and can be bundled freely. The code of the SDK must not import code of the viewer.

### The postMessage contract

Host to viewer: `{target: 'office-wasm-viewer', id, method, params}`; answers `{source: 'office-wasm-viewer', id, ok, result | error}`; events
`{source, type: 'event', event, data}`. Methods: `open`, `exportPdf`, `save`, `convert`, `fontNames`, `close`, `getState`. The full table is in
`office-wasm-viewer/README.md` (section Messages); the contract is the public interface between the two projects.

