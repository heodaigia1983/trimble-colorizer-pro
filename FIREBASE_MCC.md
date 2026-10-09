# Firebase cho Model Control Center

**Firebase project:** `trimble-model-control-center` (tách khỏi Dashboard DDC). **Gói:** Spark miễn phí. **Firestore:** `asia-southeast1` (Singapore), xác nhận trong Firebase Console. **Ứng dụng:** Model Control Center GitHub Pages. **Đăng nhập:** Google, miền `heodaigia1983.github.io` được phép. Firestore rules chỉ cho tài khoản Google đã xác minh `heodaigia1983@gmail.com` đọc/ghi dưới `mccProjects/{trimbleProjectId}`. Không lưu service account, mật khẩu hoặc token vào repo.

## Dữ liệu và cách dùng

- Mỗi Trimble project có nhánh `mccProjects/{trimbleProjectId}`. `views/{viewId}` chứa tên View và số liệu nhóm màu để dashboard đọc nhanh. `drafts/{uid}` chứa bản đang tô; `imports/{uuid}` lưu file nguồn đã nạp và số GUID.
- Ledger cấu kiện và file nhập được nén gzip nếu browser hỗ trợ, sau đó chia thành các document tối đa 180 kB dữ liệu; metadata chỉ trỏ đến revision hoàn chỉnh. Giữ revision trước làm bản khôi phục; revision cũ hơn được dọn sau khi bản mới ghi xong. Mỗi bản tối đa 500 phần (khoảng 90 MB sau nén).
- Tool tự lưu bản đang tô sau thao tác màu, thay đổi mật độ/ghi chú, chọn IFC nguồn, nạp hoặc Apply file. File nguồn `.xsr/.spm/.txt` ở Slot 1/2 được lưu nguyên byte sau khi đăng nhập. Nút **Lưu thêm file Excel/nguồn** lưu `.xlsx/.xls/.csv` và các file nguồn để tải lại; nút này không tự dùng file đó để tô màu. File IFC gốc chỉ đọc trên máy, không đưa lên Firebase.
- Dashboard trong tool liệt kê từng View với SL, m³, tấn theo màu và file đã nạp. Không cộng gộp nhiều View thành tổng dự án vì cùng cấu kiện có thể nằm ở nhiều View. Bấm tên View để mở, bấm tên file để tải bản nguồn.
- Trước khi áp lại màu từ Firebase, tool đổi external IFC ID sang runtime ID của model hiện mở. Nếu thiếu ID hoặc phiên bản model khác, chỉ hiện số liệu đã lưu, không tự tô/chọn sai cấu kiện.

## Giới hạn và kiểm tra thực tế

Firestore giới hạn 1 MiB cho mỗi document, nên dữ liệu lớn phải chia phần. Spark có hạn mức lưu trữ/đọc/ghi; dashboard chỉ lấy 40 View và 10 file gần nhất mỗi lần mở/làm mới. Cloud Storage hiện yêu cầu gói Blaze, nên chưa bật hoặc liên kết thanh toán. Nếu lượng file tăng nhiều, chuyển phần file và ledger lớn sang Storage; giữ Firestore cho metadata và dashboard. Xem tài liệu [Firestore quotas](https://firebase.google.com/docs/firestore/quotas), [Storage billing](https://firebase.google.com/docs/storage/faqs-storage-changes-announced-sept-2024), [Firebase Security Rules](https://firebase.google.com/docs/rules/get-started).

Sau khi deploy, cần anh thử trong Trimble bằng chính tài khoản Google trên: mở v2.8, bấm **Đăng nhập Google để đồng bộ**, nạp một file, tô một nhóm nhỏ, kiểm tra dòng **Đã đồng bộ** rồi Save View; mở lại View trên phiên/trình duyệt khác và đối chiếu màu, SL, m³, tấn. Trước khi có thử nghiệm này, trạng thái là code/deploy và cấu hình Firebase đã chuẩn bị, chưa được xác nhận hoạt động trên model thật.
