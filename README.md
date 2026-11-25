# Sổ Tiến Độ Thi Công

Ứng dụng web lập và theo dõi tiến độ thi công theo phương pháp đường găng (CPM).

- Bảng WBS sửa trực tiếp, liên kết FS/SS/FF/SF có độ trễ, tuần làm việc 5/6/7 ngày, ngày lễ
- Chọn nhiều dòng (Shift/Ctrl + bấm STT), cắt/chép/dán dòng, dán thẳng bảng từ Excel
- Kiểm tra dữ liệu: đầu/đuôi hở, vòng lặp, liên kết sai, % sai trình tự
- Nhập báo giá / BOQ Excel (mẫu Phúc Khang và bảng có cột Nội dung + Thành tiền): tạo tiến độ từ báo giá hoặc gắn giá trị vào tiến độ có sẵn; cảnh báo dòng tổng lệch dòng con
- Dòng tiền theo tháng (chi theo tiến độ, thu theo đợt thanh toán), vốn cần ứng tối đa, chỉ số SPI/CPI
- Gantt: việc găng, dự trữ, kế hoạch gốc, ngày báo cáo, hôm nay
- Biểu đồ nguồn lực, đường cong S, báo cáo nhanh gửi Zalo
- Nhập/xuất Excel, lưu/mở tệp `.tiendo.json`

## Chạy trên máy

```bash
npm install
npm run dev      # mở http://localhost:5173
npm test         # kiểm thử thuật toán CPM
npm run build    # tạo bản phát hành trong thư mục dist/
```

## Triển khai lên Cloudflare

Repo đã nối với Cloudflare Workers (Workers & Pages → `tien-do-thi-cong`).
Tệp `wrangler.jsonc` cho Cloudflare biết cách dựng trang: chạy `npm run build`
rồi phát hành thư mục `dist/` dưới dạng trang tĩnh. Mỗi lần đẩy code lên `main`,
Cloudflare tự dựng lại và cập nhật website.

Kiểm tra cấu hình trên máy:

```bash
npx wrangler deploy --dry-run
```

## Cấu trúc

- `src/core.js` — lõi tính CPM, lịch làm việc, dòng tiền, EVM, dữ liệu mẫu (không phụ thuộc giao diện)
- `src/boq.js` — đọc báo giá/BOQ và dựng tiến độ từ báo giá
- `src/main.js` — giao diện: bảng, Gantt, nguồn lực, báo cáo, hộp thoại
- `src/styles.css` — giao diện sáng/tối
- `tests/` — kiểm thử lõi CPM (`npm test`)
