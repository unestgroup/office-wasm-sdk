# office-wasm-sdk

Phía host của office-wasm: `sdk/office-viewer.mjs` (nhúng viewer qua iframe + postMessage, **không chứa mã AGPL**), `src/extract` (trích văn bản bằng JS thuần),
`demo/` (trang demo 2 origin: 8765 host, 8766 viewer + `e2e.cjs`), `samples/`, `tools/`.
Repo anh em cùng thư mục cha: `office-wasm-viewer` (AGPL; bắt buộc có ngay cạnh để chạy demo/e2e, hoặc đặt `OFFICE_WASM_VIEWER`), `metafile-wasm`.

- Không commit/push khi chủ chưa yêu cầu; không ghi tên model / Co-Authored-By vào commit.
- Giấy phép: Apache-2.0 (chủ đồng ý 12/11/2026; chủ bản quyền trong `NOTICE`: Datnv9 <datnv9@unestgroup.com>). Không đưa mã AGPL của viewer vào đây, không import mã viewer; mschool chỉ dùng SDK + iframe.
- Chỉ cho phép site của mình: domain nhúng nằm trong config của viewer (`serve.cjs --config`); token ký `sdk/sign.mjs` (`signViewerUrl`, Node); kiểm bằng `npm run e2e:access` (`-- --site <site đã đóng gói>`).
- Đóng gói: `npm run pack` + `pack:smoke`; `tools/trace-requests.cjs` ghi file viewer phục vụ (`--viewer-url` để thử site đã đóng gói, `--baseline` so với bản nguồn); `npm run e2e -- --viewer-url <url>` chạy e2e trên viewer đã đóng gói.
- e2e: `npm run e2e` (mặc định embedded+desktop), `-- --modes edit`, `-- --local-fonts`, `npm run e2e:coi` (mẫu thật `samples/local/de-mau.docx`, không vào git).
- Chỉ kill process của chính mình theo PID (cổng 8765/8766).
