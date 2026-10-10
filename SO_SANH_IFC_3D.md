# So sánh IFC hai ngày trên Model Control Center

## Cách dùng bản v3.5

1. Tải IFC ngày cũ và IFC ngày mới thành **hai model riêng cùng tọa độ** trong Trimble Connect; mở cả hai như ảnh anh gửi. Chỉ chọn hai file của cùng hạng mục/phạm vi xuất, không lấy Zone 02 so với Zone 03.
2. Trong Model Control Center, mở mục **So sánh IFC hai ngày trên 3D**, bấm **Model đang mở**, chọn IFC cũ và IFC mới, xác nhận cùng phạm vi rồi bấm **So sánh trực tiếp**. Không cần chọn file JSON. Tool đọc GUID và thuộc tính từ Viewer theo từng lô, hiển thị tiến độ; có nút **Dừng**.
3. Hai model thành xám để màu khác biệt nổi bật. Xanh lá = GUID chỉ có ở mốc mới; đỏ = GUID chỉ có ở mốc cũ; cam = cùng GUID nhưng khác thuộc tính/vị trí mà Viewer cấp. Bấm dòng kết quả để chọn cấu kiện trên 3D. **Khôi phục màu** trả về màu gốc model, không lấy lại màu tô tay trước đó.
4. Khi cần phân loại sâu đổi mã, dựng lại, gộp assembly và bu lông, mở phần tùy chọn, chạy `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\ifc-diff\release\SoSanhIFC.exe` với hai IFC gốc rồi nạp `diff_*.json`. Đây là đường phân tích bổ sung, không bắt buộc cho nút so sánh trực tiếp.

## Cách đọc kết quả

- `Chỉ có mốc mới` là ứng viên thêm; `chỉ có mốc cũ` là ứng viên bỏ. Chưa gọi là sản lượng dựng mới hoặc xóa thật cho đến khi xác nhận hai lần xuất cùng phạm vi.
- Luồng trực tiếp bỏ qua assembly và `IfcCovering`, chỉ lấy class cấu kiện vật lý; vì vậy tránh cộng assembly với part con. Khối lượng ứng viên thêm/bỏ chỉ hiện khi toàn bộ cấu kiện liên quan có KL hoặc V để suy ra KL với thép 7850 kg/m³. Số tấn vẫn cần kiểm tra phạm vi/cấu hình xuất. Luồng JSON giữ số liệu từ bộ so sánh hiện có.
- Cặp mẫu 15/09–16/09 khác `BaseQuantities`, `SurfaceTreatments`, `ViewColors`, và cả hai có `Export all: Off`; lớp `IfcCovering` do `SurfaceTreatments` được tách khỏi sản lượng kết cấu. Hai IFC không có căn cứ đủ để quy công việc cho một nhân viên cụ thể.
- Luồng trực tiếp dừng khi thiếu/trùng nhiều GUID hoặc hai model gần như không có GUID chung. Nó chỉ nhận diện thay đổi của cùng GUID qua thuộc tính và vị trí do Viewer trả; thay đổi hình học bên trong cấu kiện mà các giá trị này không đổi có thể không được phát hiện. Luồng JSON kiểm tra chéo mẫu GUID chỉ có ở mỗi phía để phát hiện chọn ngược.

## Giới hạn đã biết

Luồng chính dùng hai model riêng cùng mở trong Viewer, đúng ảnh anh gửi. Với model lớn, đọc GUID và thuộc tính qua API có thể mất thời gian; nếu Viewer không cấp đủ dữ liệu hoặc dừng giữa chừng, tool báo lỗi và không gọi đó là kết quả hoàn chỉnh. Kết quả so sánh chỉ giữ trong phiên trình duyệt đang mở; nạp lại hai model và chạy lại khi mở phiên khác. Chưa có bằng chứng thao tác thực tế trong Trimble cho v3.5.
