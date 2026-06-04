# Hàm Excel Tra Cứu Mã Số Thuế Việt Nam

## 📖 Giới thiệu

**Ham_Excel_Tra_MST** là một tiện ích (Add-in/VBA Macro) dành cho Microsoft Excel, giúp người dùng tra cứu nhanh chóng và tự động đưa ra thông tin từ CCCD hay MST Doanh Nghiệp.

Hàm lấy dữ liệu trực tiếp từ hệ thống API của xinvoice thuộc VietQR và Casso, cho phép trả về các thông tin quan trọng như: Tên công ty, địa chỉ, trạng thái hoạt động, cơ quan quản lý,... trực tiếp ngay trên ô (cell) của Excel.

## ✨ Tính năng nổi bật

* **Tra cứu tự động:** Chỉ cần nhập Mã số thuế, hàm sẽ tự động tra và điền thông tin.
* **Tốc độ cao & Tối ưu:** Tích hợp bộ nhớ tạm (Cache) để lưu lại các kết quả đã tra cứu, giúp các lần gọi hàm sau với cùng một MST diễn ra ngay lập tức mà không cần tra lại trên hệ thống.
* **Chống lỗi định dạng:** Tự động loại bỏ dấu cách thừa, dấu nháy đơn `'` hoặc các ký tự khoảng trắng ẩn thường gặp khi copy từ phần mềm khác.
* **Thân thiện với máy chủ:** Có tính năng trễ (Sleep 250ms) giữa các lần tra mới để tránh bị lỗi do tra quá nhanh.

## ⚙️ Cài đặt

Bạn có thể sử dụng công cụ này theo 2 cách:

### Cách 1: Sử dụng file Add-in `.xlam` (Khuyên dùng)

1. Tải file `MST_Casso+VietQR_Hetcuu.xlam` về máy tính.
2. Mở Excel, truy cập **File** > **Options** > **Add-Ins**.
3. Ở mục *Manage*, chọn **Excel Add-ins** và nhấn **Go...**
4. Nhấn **Browse...**, tìm đến file `.xlam` vừa tải về và chọn **OK**.
5. Đảm bảo file đã được tích chọn trong danh sách Add-in.

### Cách 2: Chèn trực tiếp vào file excel đang sử dung

1. Tải file `MST_Casso+VietQR_Hetcuu.xlam` về máy tính.
2. Mở file Excel của bạn
3. Kéo file `MST_Casso+VietQR_Hetcuu.xlam` vào cửa sổ đang sài
4. Chọn **Enable Macros**
* **Lưu Ý:** Phải kéo mỗi khi sử dụng.

## 🚀 Hướng dẫn sử dụng

Sau khi cài đặt, bạn có thể sử dụng hàm trực tiếp trên bảng tính Excel với cú pháp sau:

```excel
=mst(taxCode, field)
```

**Trong đó:**

* `taxCode`: Là chuỗi văn bản chứa Mã số thuế cần tra cứu (Ví dụ: `"0101248141"` hoặc chọn một ô như `A2`).
* `field`: Là trường thông tin bạn muốn lấy.

### 📋 Danh sách các trường thông tin (`field`) hỗ trợ:

| Giá trị nhập vào hàm (`field`) | Thông tin trả về tương ứng |
| --- | --- |
| `"ten"` | Tên công ty / Doanh nghiệp |
| `"diachi"` | Địa chỉ đăng ký kinh doanh |
| `"trangthai"` | Trạng thái hoạt động (Ví dụ: Đang hoạt động) |
| `"mst"` | Mã số thuế (để xác nhận lại) |
| `"loaihinhdn"` | Loại hình doanh nghiệp |
| `"quanly"` | Cơ quan thuế quản lý trực tiếp |
| `"ngaycapnhat"` | Ngày cập nhật thông tin gần nhất |

*(Lưu ý: Các giá trị `field` không phân biệt viết hoa/viết thường).*

### 💡 Ví dụ minh họa

Giả sử ô **A2** chứa Mã số thuế: `0100109106`

* Lấy tên công ty: `=mst(A2, "ten")`
* Lấy địa chỉ: `=mst(A2, "diachi")`
* Kiểm tra tình trạng hoạt động: `=mst(A2, "trangthai")`
* Lấy cơ quan quản lý: `=mst(A2, "quanly")`

## ⚠️ Khắc phục sự cố (Troubleshooting)

Hàm đã được tích hợp sẵn các thông báo lỗi bằng tiếng Việt không dấu để bạn dễ dàng nhận biết:

* **"Chua dang ky MST"**: Không tìm thấy thông tin mã số thuế này trên hệ thống.
* **"Khong tim thay truong thong tin"**: Tên công ty/Doanh nghiệp này không có trường dữ liệu bạn yêu cầu.
* **"Loi: Mang cham hoac API qua tai"**: Máy tính mất kết nối Internet, mạng quá yếu, hoặc hệ thống API đang từ chối phản hồi. Vui lòng thử lại sau.

**Lưu ý quan trọng:**

* Thiết bị của bạn **bắt buộc phải có kết nối Internet** để hàm có thể hoạt động.
* Hãy đảm bảo bạn đã bật tính năng cho phép chạy Macro trong Excel (Enable Macros).

## ☕ Ủng hộ tác giả (Donate)

Nếu bạn thấy công cụ này hữu ích, giúp tiết kiệm thời gian tra cứu và xử lý công việc hiệu quả hơn, bạn có thể cân nhắc mời mình một ly cà phê nhé. Sự ủng hộ của bạn là động lực rất lớn để mình tiếp tục duy trì, cập nhật sửa lỗi và phát triển thêm nhiều dự án miễn phí khác! ❤️

**💳 Thông tin ủng hộ:**

*   **Ngân hàng:** MB Bank
*   **Số tài khoản:** ``01001200910010``
*   **Chủ tài khoản:** NGUYEN HOANG MINH TRI

*(Quét mã VietQR bên dưới để chuyển khoản nhanh)*

![QR Donate](https://qr.sepay.vn/img?acc=VQRQAJMGT9786&bank=MBBank)

Cảm ơn bạn rất nhiều vì đã sử dụng và ủng hộ công cụ này!

## 📬 Liên hệ & Hỗ trợ

Nếu bạn phát sinh lỗi, có góp ý để cải thiện công cụ hoặc đơn giản là muốn kết nối, đừng ngần ngại liên hệ với mình qua thông tin dưới đây nhé:

* **Trang thông tin (Bio):** [hetcuu.com/in4](https://hetcuu.com/in4)

Chúc bạn thao tác thành công và làm việc hiệu quả!
