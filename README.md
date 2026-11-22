# Sổ Tiến Độ Thi Công

Ứng dụng web lập và theo dõi tiến độ thi công theo phương pháp đường găng (CPM).

- Bảng WBS sửa trực tiếp, liên kết FS/SS/FF/SF có độ trễ, tuần làm việc 5/6/7 ngày, ngày lễ
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

## Triển khai lên Cloudflare Pages

Trong Cloudflare: **Workers & Pages → Create → Pages → Connect to Git**, chọn repo này, rồi đặt:

| Mục | Giá trị |
|---|---|
| Framework preset | None |
| Build command | `npm run build` |
| Build output directory | `dist` |

Mỗi lần đẩy code lên nhánh `main`, Cloudflare tự dựng lại trang.

## Cấu trúc

- `src/core.js` — lõi tính CPM, lịch làm việc, dữ liệu mẫu (không phụ thuộc giao diện)
- `src/main.js` — giao diện: bảng, Gantt, nguồn lực, báo cáo, hộp thoại
- `src/styles.css` — giao diện sáng/tối
- `tests/` — kiểm thử lõi CPM (`npm test`)
