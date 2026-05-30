# Hướng dẫn gọi báo cáo

## Báo cáo sản lượng (file In đơn)

| Lệnh | Kết quả |
|------|---------|
| `/tonghop` | Xem tổng hợp sản lượng hôm nay |
| `/tonghop tuan` | Xem tổng hợp sản lượng tuần này |
| `/baocao` | Báo cáo sản lượng tuần trước (Thứ 2 → Chủ nhật) |
| `/baocao thang` | Báo cáo sản lượng tháng trước |
| `/baocao 18/5 24/5` | Báo cáo sản lượng từ ngày 18/5 đến 24/5 |

## Báo cáo doanh thu (file Quản lý)

| Lệnh | Kết quả |
|------|---------|
| `/doanhthu 5` | Báo cáo doanh thu tháng 5 |
| `/doanhthu 4` | Báo cáo doanh thu tháng 4 |
| `/doanhthu 5 2026` | Báo cáo doanh thu tháng 5 năm 2026 |

## Lịch gọi báo cáo

- **Thứ 2 hàng tuần**: gõ `/baocao` → nhận báo cáo sản lượng tuần trước
- **Đầu tháng**: gõ `/doanhthu [số tháng]` → nhận báo cáo doanh thu tháng trước

## Lưu ý

- Bot chạy 24/7 trên cloud (Railway) — không cần mở máy tính
- File Excel gửi vào nhóm sẽ được bot lưu lại tự động
- Bot phân biệt 2 loại file: file sản lượng (dùng /baocao) và file doanh thu (dùng /doanhthu)
