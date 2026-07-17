# PVEP OECD Governance HUB

Cổng tri thức & điều hành OECD của Ban KSNB — xây cho HĐTV/TGĐ, phong cách OECD.org (navy/cobalt) + KPI lớn kiểu McKinsey.

## Cấu trúc

- `index.html` — toàn bộ HUB (1 file, tĩnh, không backend)
- `assets/pvep-logo.png` — logo PVEP (tách từ Command Center)

## 5 module

| # | Module | Nguồn dữ liệu (thư mục OECD trên M365) |
|---|--------|------------------------------------------|
| 01 | Tài liệu gốc OECD (Original + Đào tạo) | `1. Tài liệu từ tổ chức OECD`, `2. Tài liệu tham khảo` |
| 02 | Chỉ đạo các cấp (CP/TW → PVN → Đảng ủy → HĐTV) | `3. Tap hop VB chi dao OECD cac cap`, `8. KH trien khai` |
| 03 | Kế hoạch 1017 (link OECD Command Center) | Tracker `9. Theo dõi triển khai` |
| 04 | Sản phẩm KSNB & các Ban (kèm kết quả ACGS 89,25% Hạng A) | `4. Memo Ban KSNB`, `5. Bo cau hoi khao sat`, `7. Hội nghị NĐD` |
| 05 | Báo cáo định kỳ tháng (Giao ban HĐTV–BTGĐ 2026) | `04_BÁO CÁO/.../Giao ban HĐTV-BTGĐ/2026/Thang N` |

## Cách link tài liệu

Mọi tài liệu mở qua SharePoint: `https://pvepcorp.sharepoint.com/sites/m365-ksnb/Shared Documents/<đường dẫn>`.
Người xem cần tài khoản PVEP (M365) — đây là lớp bảo mật, file gốc không rời SharePoint.

## Cách cập nhật

Mở `index.html`, tìm khối `/* ===== DATA ===== */`:

- Thêm văn bản chỉ đạo → mảng `DIRECTIVES` (lane: `tw` | `pvn` | `du` | `hdtv`)
- Thêm tài liệu → `DOCS_ORIGINAL` / `DOCS_TRAINING`
- Thêm sản phẩm → `PRODUCTS`
- Sang tháng mới → sửa `MONTHS_DONE` (số tháng đã có báo cáo)
- Số liệu KPI/tiến độ → khối `KPI` và `GROUPS`
- Link Command Center → hằng `COMMAND_CENTER_URL` (đầu khối data)

Sau khi sửa: kéo thả thư mục lên Netlify (hoặc `netlify deploy`) như Command Center.

## Số liệu gốc (snapshot 14/07/2026)

- ACGS: 407/456 = **89,25% — Hạng A** (A: 100% · B: 71,2% · C: 87,3% · D: 88,9%), bản gửi PVN 20/05/2026
- KH 1017: 21 nhiệm vụ / 8 nhóm / 63 đầu việc, tiến độ chung 11,4% (dữ liệu 10/07/2026)
- Chỉ đạo: 13 văn bản (TW 1 · PVN 4 · Đảng ủy PVEP 2 · HĐTV PVEP 6)
