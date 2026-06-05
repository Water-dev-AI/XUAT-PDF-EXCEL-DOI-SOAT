# Web Đối Soát Đơn Hàng & Theo Dõi Bán Hàng

Công cụ web tĩnh, chạy 100% trong trình duyệt. Đọc trực tiếp Google Sheet, cho
kiểm tra/sửa, rồi xuất **PDF** và **Excel** cho từng tháng (mỗi tháng gồm tab
SHOPEE và ZALO). Không có server, dữ liệu không gửi đi đâu.

## Có gì trong này
- `index.html` + `app.js` — toàn bộ ứng dụng (2 file, không cần build).

---

## A. Đưa lên GitHub Pages (làm 1 lần, ~5 phút)

1. Vào https://github.com → **New repository** → đặt tên ví dụ `doi-soat` →
   chọn **Public** → **Create repository**.
2. Bấm **uploading an existing file** → kéo thả 2 file `index.html` và `app.js`
   → **Commit changes**.
3. Trong repo → **Settings** → menu trái **Pages** → mục **Branch** chọn
   `main` / `(root)` → **Save**.
4. Đợi ~1 phút, GitHub hiện link dạng:
   `https://<tên-github>.github.io/doi-soat/` — đó là web của bạn.

Sau này sửa code chỉ cần upload đè file là web tự cập nhật.

---

## B. Lấy Google API Key (miễn phí, làm 1 lần)

1. https://console.cloud.google.com → đăng nhập.
2. Tạo **Project** mới (đặt tên tùy ý).
3. **APIs & Services → Library** → tìm **Google Sheets API** → **Enable**.
4. **APIs & Services → Credentials → Create Credentials → API key** → copy chuỗi
   `AIza...`.
5. (Nên làm) Bấm vào key → **API restrictions** → **Restrict key** → tick
   **Google Sheets API** → Save. Có thể thêm **Website restrictions** =
   địa chỉ GitHub Pages của bạn để khóa chặt hơn.
6. Mở Google Sheet → **Chia sẻ** → **Bất kỳ ai có đường liên kết → Người xem**.

> API key chỉ đọc được sheet đã ở chế độ "ai có link đều xem". Vì sheet đã công
> khai-có-link nên để key trong web không làm lộ thêm gì. Bạn vẫn nên giới hạn
> key như bước 5.

---

## C. Dùng hằng ngày

1. Mở web → dán **link Google Sheet** + **API key** → tick *Ghi nhớ API key*
   (lưu trên máy bạn, không gửi đi) → **Đọc dữ liệu**.
   - Hoặc: kéo-thả **file Excel đã xuất trước đó** vào ô "Mở lại file Excel" để
     xem/sửa nhanh mà không cần đọc lại Google Sheet.
2. Web tự nhận các tab `T5 SHOPEE`, `T5 ZALO`... (bỏ qua tab ghi chú), gom theo
   tháng. Mỗi tháng một mục.
3. **Kiểm tra trước khi xuất** — mỗi tab hiện:
   - ⚠ **Ô có chữ ghi chú lẫn trong mã** (vd `HỦY ĐƠN`, `CHỜ XỬ LÝ 091...`):
     web tách sẵn phần mã, bạn sửa lại ô "mã xuất ra" cho đúng.
   - ● **Dòng trống ngày order** (chỉ những dòng *thật sự* thiếu — dòng con của
     đơn gộp tự thừa hưởng ngày nên không bị báo nhầm).
   - Bảng **xem trước** đúng như bản in.
4. Xuất: **Excel/PDF (tháng đang xem)** hoặc **(tất cả)**.
   - Excel: số là **số thật** (cộng/tính được), có **màu nền** (ngắt ngày,
     đơn hủy, dòng hóa đơn), merge cột mã giữ nguyên.
   - PDF: khổ **ngang A4**, font hẹp, mở qua cửa sổ in của trình duyệt → chọn
     *Save as PDF*. Nhớ bật **Background graphics** trong hộp thoại in để có màu.

---

## D. Quy tắc xử lý (đã cài sẵn)

- **Nhận cột theo từ khóa ở hàng 2**, không theo vị trí → đổi chỗ cột / thêm cột
  vẫn chạy đúng.
- **Tiêu đề file**: *Bảng đối soát đơn hàng, theo dõi bán hàng - Tháng X/20XX*
  (năm tự suy từ dữ liệu).
- **ZALO**: bỏ cột Mã Đơn Vận/Đơn Hàng (bán ngoài, không có mã bán ra).
- **SHOPEE**: ô Mã Đơn Vận merge → Mã Order ID gộp merge tương ứng; các mã con
  dồn về thành nhiều dòng trong ô. Ngày Order / Ngày DV **không** merge.
- **Đơn giá chưa VAT (sau giảm, trừ SIM trắng)**: nếu tab cũ thiếu cột này, web
  tự tính `(đơn giá có VAT sau giảm − tiền SIM trắng/số lượng) ÷ 1.1`.
- **Đơn hủy** (số lượng âm): giữ nguyên, tô nền nhận biết.
- **Hàng ngắt ngày** (trống + nền xanh): xuất ra thành dải xanh ngăn cách.
- **Thông tin xuất hóa đơn** (MST, tên, địa chỉ, email): đưa xuống **một dòng
  phụ** ngay dưới đơn, tránh thêm 4 cột dài.

## E. Khi cấu trúc sheet đổi trong tương lai
Web khớp cột theo *nội dung tiêu đề*. Nếu sau này bạn đổi tên cột nhiều, mở
`app.js`, tìm `FIELD_DEFS` ở đầu file và thêm từ khóa vào danh sách `match`
tương ứng. Không cần sửa chỗ nào khác.
