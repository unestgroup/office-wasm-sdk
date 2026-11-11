# office-wasm-sdk

Phía host của office-wasm: `sdk/office-viewer.mjs` (nhúng viewer qua iframe + postMessage, **không chứa mã AGPL**), `src/extract` (trích văn bản bằng JS thuần),
`demo/` (trang demo 2 origin: 8765 host, 8766 viewer + `e2e.cjs`), `samples/`, `tools/`.
Repo anh em cùng thư mục cha: `office-wasm-viewer` (AGPL; bắt buộc có ngay cạnh để chạy demo/e2e, hoặc đặt `OFFICE_WASM_VIEWER`), `metafile-wasm`.

- Không commit/push khi chủ chưa yêu cầu; không ghi tên model / Co-Authored-By vào commit.
- Giấy phép: Apache-2.0 (chủ đồng ý 12/11/2026; tên chủ bản quyền trong `NOTICE` còn là "the office-wasm-sdk authors", thay bằng pháp nhân khi có). Không đưa mã AGPL của viewer vào đây, không import mã viewer; mschool chỉ dùng SDK + iframe.
- e2e: `npm run e2e` (mặc định embedded+desktop), `-- --modes edit`, `-- --local-fonts`, `npm run e2e:coi` (mẫu thật `samples/local/de-mau.docx`, không vào git).
- Chỉ kill process của chính mình theo PID (cổng 8765/8766).
