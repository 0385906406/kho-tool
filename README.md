# Kho Tool Đối Soát

Trang tải các tool cho khâu giao nhận phụ tùng. Hiện có **Tool Đối Soát NCC ↔ MB51**, gồm 2 bản:

- **Dùng cho Microsoft Excel** (macro VBA): gói `ToolDoiSoat_DungCho_MicrosoftExcel_*.zip` — file `dist/ToolDoiSoat_MicrosoftExcel.xlsm`, mã nguồn `dist/MaNguon_MicrosoftExcel_VBA.bas`
- **Dùng cho WPS Office** (macro JS): gói `ToolDoiSoat_DungCho_WPSOffice_*.zip` — file `dist/ToolDoiSoat_WPSOffice.xlsm`, mã nguồn `dist/MaNguon_WPSOffice_JS.txt`

Mỗi lần push lên nhánh `main`, GitHub Actions chạy `node build-public.js` và đăng thư mục `public/` lên GitHub Pages.

## Cấu trúc

| File / thư mục | Vai trò |
|---|---|
| `tools.config.js` | Danh sách tool: tên, mô tả, các bước dùng, file tải về. **Thêm tool mới ở đây.** |
| `public-index.html` | Giao diện trang web công khai |
| `build-public.js` | Tạo `public/` (trang + file tải) từ `tools.config.js` |
| `dist/`, `*.zip` | File tool và file mẫu để tải về (đóng gói trên máy Windows có WPS) |
| `claude-artifact.html`, `build-tools.js` | Bản trang dành cho claude.ai (không dùng cho GitHub Pages) |
| `.github/workflows/pages.yml` | Tự build và đăng lên GitHub Pages |

## Cập nhật tool

1. Đóng gói lại tool trên máy có WPS, chép file mới vào `dist/` và file `.zip`.
2. Sửa phiên bản / mô tả trong `tools.config.js` nếu cần.
3. Commit và push lên `main`. Trang tự cập nhật sau khoảng 1 phút.

Xem thử trên máy trước khi push: `node build-public.js`, rồi mở `public/index.html`.

## Lưu ý

Repo công khai: không đưa file số liệu thật (PO, MB51 thật) vào đây. File mẫu trong `dist/` là số liệu giả lập.
