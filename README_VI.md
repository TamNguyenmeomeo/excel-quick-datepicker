# Excel Quick Date Picker (Tiện ích Chọn Ngày Nhanh trên Excel) 📅

Tiện ích mở rộng Microsoft Excel Add-in (`.xlam`) siêu nhẹ, giao diện hiện đại giúp hiển thị bảng lịch chọn ngày trực quan. Tiện ích được viết hoàn toàn bằng VBA (Visual Basic for Applications), cho phép người dùng click đúp (double-click) vào ô bất kỳ để chọn ngày nhanh chóng từ cửa sổ lịch thay vì phải gõ thủ công.

---

## 🌟 Các tính năng nổi bật

*   **Giao diện lịch hiện đại:** Cửa sổ chọn ngày đẹp mắt, dễ dàng chuyển đổi nhanh giữa các ngày, tháng, năm.
*   **Kích hoạt bằng click đúp:** Tự động hiện bảng lịch khi bạn click đúp vào các ô nhập dữ liệu được chỉ định.
*   **Phím tắt toàn cục:** Nhấn phím tắt nhanh để mở lịch ở bất kỳ ô nào trên mọi bảng tính.
*   **Tự động nhận diện ngày cũ:** Nếu trong ô đang chọn đã có sẵn một ngày (ví dụ `15/06/2026`), lịch sẽ tự động mở đúng ngày đó để bạn dễ điều chỉnh.
*   **Tùy chỉnh giao diện không giới hạn:** Thay đổi màu sắc (màu nền, màu chữ chủ đạo, màu tô khi rê chuột, màu thứ Bảy/Chủ Nhật) bằng mã màu RGB để phù hợp với màu thương hiệu bảng tính của bạn.
*   **Hoạt động Offline & An toàn:** 100% offline, không gửi dữ liệu ra bên ngoài. Tích hợp trực tiếp vào Excel dưới dạng add-in hệ thống, không làm nặng bảng tính của bạn.

---

## 🎨 Minh họa Giao diện

![Bảng lịch chọn ngày trên Excel](excel_calendar_preview.png)

---

## 📘 Hướng dẫn Sử dụng cho Người dùng cuối

Tiện ích này giúp loại bỏ việc gõ ngày tháng thủ công — vốn rất dễ bị gõ sai định dạng hoặc nhầm lẫn năm.

### 1. Phím tắt mở nhanh lịch chọn ngày
Sau khi cài đặt add-in thành công:
*   Chọn bất kỳ ô nào trên Excel và nhấn tổ hợp phím **`Ctrl + \``** (Phím Control và phím dấu huyền nằm ngay dưới phím `Esc`, bên cạnh số `1`).
*   Bảng lịch sẽ hiện ra. Bạn chỉ cần chọn ngày và nhấn **OK** (hoặc click đúp vào ngày đó), ngày sẽ tự động được điền vào ô với định dạng `dd/mm/yyyy`.

### 2. Tự động mở lịch khi Click đúp (Dành cho nhà phát triển/kỹ sư bảng tính)
Bạn có thể lập trình để khi người dùng click đúp vào các ô nhập ngày (ví dụ cột "Ngày bắt đầu"), bảng lịch sẽ tự hiện lên:
1. Nhấn tổ hợp phím **`ALT + F11`** để mở cửa sổ VBA.
2. Click đúp vào tên trang tính (ví dụ: `Sheet1`) ở danh mục bên trái.
3. Dán đoạn mã sau vào:

```vba
' Hiện lịch chọn ngày khi click đúp vào ô H16
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Range("H16")) Is Nothing Then
        Cancel = True ' Ngăn Excel chuyển sang chế độ sửa chữ mặc định
        
        Dim dateVariable As Variant
        dateVariable = CalendarForm.GetDate
        
        If dateVariable <> 0 Then 
            Target.Value = dateVariable
        End If
    End If
End Sub
```

---

## 💻 Hướng dẫn Cài đặt Add-in vào Excel

Để sử dụng tiện ích này trên tất cả các file Excel của bạn:

1.  **Tải file add-in:** Tải file `QuickDataPicker.xlam` về máy tính và lưu ở một thư mục an toàn (không nên xóa hoặc di chuyển file này sau khi đã cài).
2.  **Mở cửa sổ Quản lý Add-ins trong Excel:**
    *   Mở Excel -> Chọn thẻ **Developer** (Nhà phát triển) -> Chọn **Excel Add-ins**.
    *   *(Nếu không thấy thẻ Developer, vào **File** -> **Options** -> **Add-ins** -> ở mục *Manage* chọn *Excel Add-ins* và nhấn **Go...**).*
3.  **Thêm Tiện ích:** Nhấn nút **Browse...** -> Tìm đến thư mục bạn đã lưu file `QuickDataPicker.xlam` ở Bước 1 -> Chọn file và nhấn **OK**.
4.  **Kích hoạt:** Đảm bảo ô trống bên cạnh chữ `QuickDataPicker` đã được tích chọn, sau đó nhấn **OK**. Tiện ích đã được kích hoạt thành công!

---

## 📦 Nội dung dự án

*   `QuickDataPicker.xlam`: File Add-in hoàn chỉnh, sẵn sàng cài đặt.
*   `CalendarForm v1.5.2.xlsm`: File Excel ví dụ chứa bảng kiểm thử và hướng dẫn tích hợp chi tiết.
*   `CalendarForm.frm` & `CalendarForm.frx`: Mã nguồn thô của Form Lịch dành cho nhà phát triển muốn nhập (import) thủ công vào dự án riêng.
*   `Module_Calender/`: Các mô-đun hỗ trợ tính toán lịch và màu sắc.

---

## 📄 Bản quyền
Dự án được cấp phép theo Giấy phép MIT - xem tệp [LICENSE](LICENSE) để biết thêm chi tiết.
