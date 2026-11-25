# Sổ Tiến Độ Thi Công

Ứng dụng web lập và theo dõi tiến độ thi công theo phương pháp đường găng (CPM).

- Bảng WBS sửa trực tiếp, liên kết FS/SS/FF/SF có độ trễ, tuần làm việc 5/6/7 ngày, ngày lễ
- Chọn nhiều dòng (Shift/Ctrl + bấm STT), cắt/chép/dán dòng, dán thẳng bảng từ Excel
- Kiểm tra dữ liệu: đầu/đuôi hở, vòng lặp, liên kết sai, % sai trình tự
- Nhập báo giá / BOQ Excel (mẫu Phúc Khang và bảng có cột Nội dung + Thành tiền): tạo tiến độ từ báo giá hoặc gắn giá trị vào tiến độ có sẵn; cảnh báo dòng tổng lệch dòng con
- Dòng tiền theo tháng (chi theo tiến độ, thu theo đợt thanh toán), vốn cần ứng tối đa, chỉ số SPI/CPI
- Cập nhật tuần (% thực tế, ghi chú vướng mắc, chi phí thực tế), lưu lịch sử để vẽ đường thực tế trên đường cong S
- Look-ahead 3/4/6 tuần theo hạng mục, xuất Excel
- Chế độ Họp giao ban toàn màn hình: tổng quan, tuần qua, việc chậm, 2 tuần tới, dòng tiền, đường găng
- Mô hình 3D khối (three.js): tầng và diện tích lấy từ báo giá hoặc đoán từ tên công việc, tự gắn công việc theo tầng, xem theo kế hoạch tại một ngày (có phát tự động) hoặc thực tế tại ngày báo cáo, chụp ảnh PNG
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
- `src/model3d.js` — tầng, cấu kiện, gắn công việc, mức hoàn thành cho 3D
- `src/view3d.js` — hiển thị 3D bằng three.js (chỉ tải khi mở tab 3D)
- `src/main.js` — giao diện: bảng, Gantt, nguồn lực, báo cáo, hộp thoại
- `src/styles.css` — giao diện sáng/tối
- `tests/` — kiểm thử lõi CPM (`npm test`)
