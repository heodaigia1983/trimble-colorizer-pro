# HANDOFF — Trimble Model Control Center

**Cập nhật:** 09/10/2026, sau khi triển khai v2.7. Đây là file bàn giao để mở chat mới: đọc hết, rồi đối chiếu trạng thái repo và Trimble hiện tại. Bản handoff cũ 30/06/2026 được giữ tại `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\_handoff-backups\HANDOFF_TRIMBLE_TOOL-2026-10-09-before-update.md`. Bản sao mới nhất của file này cũng nằm ở thư mục cha `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\HANDOFF_TRIMBLE_TOOL.md` để anh thấy ngay trong thư mục làm việc.

## Anh Thảo và quy ước làm việc

Anh Lê Văn Thảo (Anh Heo), Head of Shop Drawing DDC, làm Tekla, IFC, Trimble Connect và kết cấu thép. Xưng em, gọi anh; trả lời tiếng Việt ngắn, thẳng vào kết quả. Chỉ sửa phần anh yêu cầu, giữ phần còn lại. Không suy đoán số liệu hoặc báo đã test trong Trimble nếu chưa quan sát thực tế.

Trước mỗi thay đổi trên GitHub project của anh, giữ backup có thể khôi phục. Trong phiên 09/10/2026 anh đã cho phép Codex sửa code, commit và push; quy trình cũ trong handoff 30/06 (chỉ viết prompt Claude Code, anh tự push) không còn là ràng buộc hiện tại. Với yêu cầu mới, theo chỉ dẫn mới nhất của anh. Nếu soạn prompt Claude Code thì phải bắt đầu “Bước 1: Chẩn đoán”, sửa sau, find/replace chính xác, checklist rõ. Khi sửa tool, tăng `v=N` đồng bộ trong `index.html` và `manifest.json`, cập nhật phiên bản/ngày giờ hiển thị.

Anh yêu cầu **sau mỗi 15 thao tác làm việc có ý nghĩa**, lưu quá trình, quy trình có thể lặp lại và kinh nghiệm. Đếm các bước thực chất như khảo sát, chẩn đoán, backup, sửa, commit, push, xác minh; không tính tin nhắn trạng thái hay lần đọc lặp. Nếu việc kết thúc trước mốc 15, ghi kết quả và việc còn mở. Cập nhật mục “Nhật ký 15 thao tác” cuối file. Bộ đếm bắt đầu lại từ bản handoff này.

## Repo, file và bản đang chạy

- Repo có `.git`: `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\trimble-colorizer-pro\`.
- GitHub: `https://github.com/heodaigia1983/trimble-colorizer-pro`, nhánh `main`.
- Tool 1: `https://heodaigia1983.github.io/trimble-colorizer-pro/manifest.json`; trang v2.7: `https://heodaigia1983.github.io/trimble-colorizer-pro/index.html?v=32`.
- Tool 2 riêng ở `progress-tracker/`; không sửa khi việc chỉ liên quan Tool 1.
- IFC anh dùng: `C:\Users\Admin\Downloads\KC Gia Binh.ifc`. File này được chọn trong browser để tính thể tích; không có bằng chứng đã upload lên GitHub.
- Code v2.7 ở `1d9aed7` (09/10/2026); v2.6 ở `7eb7755`, v2.5 ở `088de34`. Kiểm tra lại HEAD, branch, remote, status và manifest trước khi tiếp tục.

## Diễn biến phiên 09/10/2026

Anh mở Tool A trên Trimble; chọn hàng chục rồi hàng nghìn cấu kiện nhưng bảng **KL (Tấn)** bằng 0/“Chưa có dữ liệu”. Viewer không cấp đủ thuộc tính KL cho model đó. Anh yêu cầu lấy thể tích từ IFC, chọn tay nhiều nhóm cấu kiện theo màu, tick màu để xem tổng m³ và tấn.

| Commit | Phần đã triển khai |
|---|---|
| `5aafc82` | Thể tích IFC và nhóm màu có thể chọn. |
| `c2fbd7f` | Làm rõ quy trình Slot 3: Apply → chọn màu → OK. |
| `872cc45` | Chọn màu trong bảng thì cấu kiện tương ứng được chọn trên model. |
| `8fa730e` | Hiện phiên bản và thời điểm cập nhật. |
| `5265e58` | Sửa chuyển runtime ID khi lấy chỉ số. |
| `2ff8bd2` | Đọc IFC gốc vì Viewer chỉ cấp model đã chuyển đổi. |
| `bf7a6c3` | Khôi phục số liệu nhóm màu theo View từ trình duyệt. |
| `088de34` | v2.5: lưu tóm tắt m³/tấn nhóm màu trong mô tả Trimble View. |
| `7eb7755` | v2.6: thử lại đọc màu khi View tải xong; nếu chưa có ID vẫn hiện m³/tấn lưu trong View. |
| `1d9aed7` | v2.7: kết nối Trimble API ngay khi mở tool; nút Xem khối lượng cũng chủ động đọc View. |

Anh đã thấy bảng **SL, KL (Tấn), V (m³), ρ (kg/m³)**. Mật độ thép mặc định là `7850 kg/m³` khi thiếu KL; tấn = kg/1000. Worker cộng thể tích có dấu từ tam giác IFC; về nguyên tắc phần rỗng của ống/hộp được trừ nếu lưới IFC biểu diễn đúng, không dùng thể tích hộp bao. Anh phản hồi “ok r em”, “hoạt động tốt” sau giai đoạn sửa thể tích. Chưa có bảng đối chiếu độc lập từng cấu kiện với Tekla, nên không khẳng định độ chính xác tuyệt đối.

Vấn đề kế tiếp: Save View hôm nay, hôm sau mở lại vẫn phải xem nhóm vàng/xanh bao nhiêu tấn. Trước v2.5 dữ liệu chỉ nằm trong `localStorage` cùng browser. Anh hỏi Google/GitHub; phương án đã chọn là lưu tóm tắt nhỏ trong **Trimble View description**. Nếu Trimble không giữ được thì mới tính tiếp Google/cloud khác. Chưa có Google endpoint hoặc quyền xác thực cho việc này.

16:36 cùng ngày, anh gửi ảnh v2.5: model còn dải vàng và xanh nhưng Slot 3 không có bảng màu, ô lựa chọn báo “Chưa chọn gì”. Chẩn đoán từ code: v2.5 gọi `getColoredObjects()` đúng một lần khi nhận View ID; nếu Viewer chưa tải xong màu thì ledger rỗng, vòng poll sau đó chỉ nhìn ID nên không thử lại. Đây là nguyên nhân có căn cứ từ source, nhưng chưa được xác nhận bằng log API của phiên Trimble anh. v2.6 thử lại tối đa 12 lần cách 5 giây, kiểm tra count và chữ ký tập cấu kiện, thử `getObjects` theo màu khi cần; đồng thời hiện bảng **số liệu lúc Save View** từ `MCCQ1` khi chưa lấy được ID, cho tick màu cộng m³/tấn nhưng chưa cho chọn đối tượng trên model. Nếu View không có `MCCQ1`, bảng tổng hợp không thể tự dựng từ ảnh màu; cần kiểm tra View đã được Save bằng tool v2.5 sau khi nạp IFC hay chưa.

17:24 anh gửi ảnh v2.6: IFC gốc đã chọn, model màu cam/xanh, Slot 3 trống. Anh xác nhận View được lưu bằng Save View của tool sau khi bảng từng có m³/tấn; đợi hơn một phút vẫn không có cảnh báo LOG. Đọc tiếp source thấy nguyên nhân gốc chưa xử lý ở v2.6: `getAPI()` chỉ chạy sau một số thao tác như Apply/Save, còn khi mở tool chỉ gọi `renderViewList()`. Chọn IFC và nút Xem khối lượng khi ledger rỗng cũng không gọi API. Vì vậy `initActiveView()` và vòng poll khôi phục View chưa khởi động. v2.7 gọi `getAPI()` khi mở tool, tránh kết nối trùng bằng `_apiPromise`, cho nút Xem khối lượng chủ động đọc View, và thử đọc màu đang hiển thị nếu `getCurrentView()` chưa trả ID. Chưa dùng fallback `Presentation`; chỉ nghiên cứu API này nếu v2.7 vẫn không trả object IDs.

## Kiến trúc và giới hạn hiện tại

`_colorLedger` trong `app.js` là Map khóa `modelId:runtimeId`, nguồn dữ liệu màu của Slot 3. Bảng gộp theo màu; checkbox cộng m³/tấn; click ô màu gọi `selectLedgerColorOnViewer()` để chọn cấu kiện trên model. Mỗi item có màu, ghi chú, mật độ và có thể có `volumeM3`.

`ifc-volume-worker.js` dùng `web-ifc@0.0.78`, cần IFC gốc có header `ISO-10303-21;` và đơn vị SI rõ ràng. `app.js` dùng `convertToObjectIds` đổi runtime IDs sang GUID IFC, worker tìm Express ID, tính thể tích mesh có dấu rồi đổi m³. Nếu không khớp GUID, không mở được IFC hay không rõ đơn vị thì báo thiếu, không lấy 0 làm thể tích thật. Cache thể tích theo từng object và ghi vào ledger cục bộ.

Save View v2.5: `rebuildColorLedgerSummary()` tính các nhóm. Khi **mọi nhóm** đủ m³/tấn và mô tả dài tối đa 240 ký tự, `makeViewDescription()` thêm `MCCQ1:` vào View description. Mỗi nhóm lưu màu, số lượng, thể tích, KL, mật độ và chữ ký tập `modelId:runtimeId` bằng base36. Độ phân giải m³ là 10⁻⁶, KL kg là 0.001. Luồng `createView` → `updateView` → `selectView` → `getView` đọc lại; chỉ log xác nhận cloud khi các số đọc lại khớp. Nếu không, cảnh báo và vẫn giữ bản cục bộ. Code ở đầu `app.js` và khoảng dòng 1180–1245.

Mở View: `getCurrentView`/`getView` lấy mô tả. Nếu máy không có localStorage, thử `getColoredObjects()` dựng lại ledger từ màu Viewer. Chỉ dùng m³/tấn lưu trên View khi màu, số lượng, mật độ và chữ ký cấu kiện khớp. Nếu model đổi hoặc runtime IDs thay đổi giữa các phiên, số có thể không hiện vì chữ ký lệch; đây là cơ chế tránh dùng sai số cũ. View tạo trước v2.5 chưa có `MCCQ1:`; cần nạp IFC và Save lại View mới. Mô tả 240 ký tự cũng giới hạn số nhóm màu có thể lưu cloud.

## Bằng chứng và việc chưa xác minh

Trước commit v2.5, `node --check app.js` và `git diff --check` đều exit 0; không thêm/chạy test suite. Backup trước v2.5: `backup/mcc-before-cloud-view-qty-v30-20261009` → `bf7a6c3`. Backup trước sửa handoff: `backup/mcc-before-handoff-20261009` → `088de34`; bản handoff đầu tiên: `backup/mcc-before-handoff-finalize-20261009` → `a0b482f`. Các tag đã push và đối chiếu trên origin. Workflow GitHub Pages `37903677651` deploy thành công; HTTP công khai trả manifest `v=30`, index `v2.5` và JS có `MCCQ1:`.

**Chưa được xác minh trong phiên Trimble thật:** Sau `updateView`, Trimble có giữ `description` không; v2.7 có khôi phục ledger và m³/tấn trên View anh hay không. Phiên Chrome Codex không có tab Trimble đăng nhập. Không coi deploy Pages hay API docs là bằng chứng tính năng đã hoạt động trên model của anh.

Anh có thể kiểm tra: mở Tool 1 thấy v2.7; mở View cũ có màu, chọn IFC gốc rồi đợi bảng màu tự hiện hoặc bấm Xem khối lượng. Xem LOG ở cuối panel: phải có dòng “Đã kết nối Trimble API.” Nếu View chưa có `MCCQ1`, code vẫn thử dựng nhóm từ màu đang hiện và dùng IFC gốc để tính m³/tấn. Nếu không có bảng, lấy LOG và ID/trạng thái View; chẩn đoán dữ liệu API thật trước khi sửa tiếp. Nếu phải chuyển Google, cần hỏi anh nơi lưu và quyền phù hợp; không nhúng token vào trang công khai.

## Việc khác còn mở

SPM `.xsr` và GUID assembly Tekla **chưa được chứng minh** khớp Viewer. GUID Tekla khác `IfcGloballyUniqueId`; trước khi làm matching phải chẩn đoán property Viewer thật. Handoff cũ 30/06 liệt kê lọc Assembly Position là chưa làm, nhưng Tool 1 hiện đã có UI và code cho mục này; thông tin cũ không còn chính xác. Không trộn nhiệm vụ SPM vào sửa m³/tấn nếu anh chưa yêu cầu.

## Quy trình tiếp tục và nhật ký 15 thao tác

Khi có yêu cầu mới: đọc file này → kiểm tra repo/phiên bản và bằng chứng mới từ anh → chẩn đoán → backup → sửa đúng phạm vi → tăng version/cache-bust → commit/push khi được giao → kiểm tra deploy → tách bạch kiểm tra công khai với thao tác Trimble thực. Không tuyên bố “đã chạy tốt trên Trimble” nếu chưa quan sát hoặc anh chưa xác nhận.

**Mốc 09/10/2026:** v2.5 đã lên Pages, backup đầy đủ; còn chờ thử Save/Open View trong Trimble. Kinh nghiệm: Viewer có thể thiếu KL và chỉ cấp IFC chuyển đổi, cần IFC gốc để tính m³. localStorage không đủ cho mở View trên máy khác. Lưu Trimble View cần đọc lại xác nhận; khi model đổi cần đối chiếu tập cấu kiện để tránh hiển thị số cũ. Bắt đầu đếm 15 thao tác mới sau bản handoff này; các mốc sau ghi nối tiếp tại đây.

**Mốc 09/10/2026 16:45, sau 15 thao tác:** Đã xem ảnh lỗi, đọc code và tài liệu Viewer API chính thức, tìm điểm chỉ đọc màu một lần, tạo tag backup `backup/mcc-before-view-recovery-v31-20261009` → `b4e47bd`, sửa `app.js`/`index.html`/`manifest.json`, kiểm tra cú pháp và diff, commit/push `7eb7755`. Backup ngay trước sửa handoff: `backup/mcc-v26-before-handoff-update-20261009` → `7eb7755`. Bài học: View ID có thể xuất hiện trước khi màu cấu kiện tải xong; cần retry có giới hạn và luôn hiện rõ khi chỉ đang xem số liệu lịch sử, chưa chọn được cấu kiện hiện tại. **Chưa có bằng chứng v2.6 chạy đúng trên phiên Trimble của anh**; chờ anh mở View và kiểm tra log/bảng. Bộ đếm 15 thao tác bắt đầu lại sau mốc này.

**Mốc 09/10/2026 17:34:** Sau ảnh v2.6 và câu trả lời của anh, phát hiện nhánh khôi phục View không hề chạy vì app không kết nối API khi mở; đây là bài học quan trọng hơn giả thuyết Viewer tải chậm. Backup `backup/mcc-before-presentation-recovery-v32-20261009` → `001c2ad`; sửa khởi động và nút Xem khối lượng; commit/push `1d9aed7`. GitHub Pages workflow `37918381657` thành công; HTTP công khai xác nhận manifest `v=32`, index v2.7, JS có startup connect. Backup trước sửa handoff: `backup/mcc-v27-before-handoff-update-20261009` → `1d9aed7`. **Chưa có test thực tế v2.7 trên Trimble của anh.** Kinh nghiệm: trước khi chẩn đoán retry hay dữ liệu API, xác minh đường gọi khởi tạo API có thực sự chạy trong luồng người dùng. Bắt đầu đếm 15 thao tác mới sau mốc này.

**Mốc 09/10/2026 18:35, sau hơn 15 thao tác — Firebase đang triển khai:** Anh chọn dashboard ngay trong MCC và yêu cầu tạo Firebase riêng. Đã tạo project `trimble-model-control-center` dưới tài khoản `heodaigia1983@gmail.com`, gói Spark miễn phí; bật Firestore ở production mode, bật Google Sign-in, cho phép miền `heodaigia1983.github.io`. Firestore rules đã công bố: chỉ email Google đã xác minh `heodaigia1983@gmail.com` được đọc/ghi nhánh `mccProjects/{projectId}/...`; nhánh khác vẫn bị chặn. Backup GitHub trước khi sửa: tag `backup/mcc-v27-before-firebase-20261009` → `48c36cc`, đã push và kiểm tra trên origin. Đang sửa v2.8/cache `v=33`: `mcc-cloud.js` dùng Firestore lưu tóm tắt View cho dashboard, bản màu/file nhập chia thành các document nhỏ có nén; Google Auth; tự lưu draft sau thao tác; dùng external IFC ID để đổi lại runtime ID khi mở lại. IFC gốc không được tải lên Firebase. Cloud Storage cần Blaze nên chưa bật thanh toán; dữ liệu lớn hiện chia phần trong Firestore. **Chưa commit/push v2.8, chưa xác minh trên Trimble thật.** Bước tiếp: rà tính đúng, cập nhật bản bàn giao, commit/push, kiểm tra Pages và để anh thử Google Sign-in/Save/Open View trong Trimble. Kinh nghiệm: project Firebase đã có của Dashboard DDC không nên dùng lẫn khi anh chọn project riêng; Firestore giới hạn 1 MiB/doc, nên không ghi toàn bộ ledger hay file vào một doc; runtime ID có thể đổi nên phải kiểm tra external ID trước khi áp lại màu.
