# So sánh IFC hai ngày trên Model Control Center

## Cách dùng bản v3.4

1. Trên máy anh, mở `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\ifc-diff\release\SoSanhIFC.exe`. Chọn IFC cũ và IFC mới, bấm **Bắt đầu so sánh**. Khi xong, mở thư mục kết quả và lấy file `diff_*.json` của đúng lần chạy.
2. Mở đồng thời **hai mốc IFC** trong cùng Trimble Viewer. Anh có thể thử hai phiên bản của một file/model như dự định; chỉ tiếp tục khi danh sách **Model đang mở** hiện hai `versionId` riêng. Nếu Viewer chỉ cho mở một phiên bản tại một thời điểm, tải hai IFC thành hai model riêng cùng tọa độ để xem so sánh đồng thời. File 500–700 MB được xử lý trên máy bằng ứng dụng Windows; GitHub Pages không nhận IFC gốc.
3. Trong Model Control Center, mở mục **So sánh IFC hai ngày trên 3D**, chọn `diff_*.json`, bấm **Model đang mở**, đối chiếu rõ mốc cũ/mới với tên file và ID phiên bản, rồi bấm **Tô so sánh**. Khi tên hiển thị khác tên IFC nguồn, chỉ tích ô xác nhận sau khi đã đối chiếu thủ công.
4. Model cũ thành xám. Kết quả có màu: xanh lá = chỉ có ở mốc mới; cam = khác thuộc tính/vị trí; đỏ = chỉ có ở mốc cũ; tím = tạo lại/dựng lại; xanh dương = đổi mã; xanh ngọc = bu lông mới. Bấm một dòng trong bảng để chọn cấu kiện trên model. **Khôi phục màu** trả các trạng thái so sánh về màu gốc của model.

## Cách đọc kết quả

- `Chỉ có mốc mới` là ứng viên thêm; `chỉ có mốc cũ` là ứng viên bỏ. Chưa gọi là sản lượng dựng mới hoặc xóa thật cho đến khi xác nhận hai lần xuất cùng phạm vi.
- Bảng tấn lấy từ báo cáo `diff_*.json` của bộ so sánh hiện có. Bảng không cộng khối lượng assembly với part con. Khi báo cáo có cảnh báo về cấu hình xuất, số tấn chỉ để đối chiếu.
- Cặp mẫu 15/09–16/09 khác `BaseQuantities`, `SurfaceTreatments`, `ViewColors`, và cả hai có `Export all: Off`; lớp `IfcCovering` do `SurfaceTreatments` được tách khỏi sản lượng kết cấu. Hai IFC không có căn cứ đủ để quy công việc cho một nhân viên cụ thể.
- Tool kiểm tra chéo mẫu GUID chỉ có ở một mốc để phát hiện chọn ngược. GUID khớp quá ít thì dừng trước khi tô. Nếu Viewer không trả đúng runtime ID, xem dòng trạng thái và LOG; không lấy số không khớp làm bằng chứng thay đổi.

## Giới hạn đã biết

Luồng hiện tại cần hai mốc IFC hiện đồng thời như **hai versionId được tải** trong Viewer. Khả năng Trimble thực sự tải song song hai phiên bản của cùng file phải xác minh trên dự án anh; nếu không, dùng hai model riêng cùng tọa độ. Kết quả so sánh chỉ được giữ trong phiên trình duyệt đang mở; file JSON do ứng dụng Windows tạo là nguồn để nạp lại lần sau. Chưa có bằng chứng thao tác thực tế trong Trimble cho v3.4.
