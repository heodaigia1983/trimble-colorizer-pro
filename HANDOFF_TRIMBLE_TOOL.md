# HANDOFF — Trimble Model Control Center

**Cập nhật:** 09/10/2026, sau khi triển khai v2.5. Đây là file bàn giao để mở chat mới: đọc hết, rồi đối chiếu trạng thái repo và Trimble hiện tại. Bản handoff cũ 30/06/2026 được giữ tại `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\_handoff-backups\HANDOFF_TRIMBLE_TOOL-2026-10-09-before-update.md`.

## Anh Thảo và quy ước làm việc

Anh Lê Văn Thảo (Anh Heo), Head of Shop Drawing DDC, làm Tekla, IFC, Trimble Connect và kết cấu thép. Xưng em, gọi anh; trả lời tiếng Việt ngắn, thẳng vào kết quả. Chỉ sửa phần anh yêu cầu, giữ phần còn lại. Không suy đoán số liệu hoặc báo đã test trong Trimble nếu chưa quan sát thực tế.

Trước mỗi thay đổi trên GitHub project của anh, giữ backup có thể khôi phục. Trong phiên 09/10/2026 anh đã cho phép Codex sửa code, commit và push; quy trình cũ trong handoff 30/06 (chỉ viết prompt Claude Code, anh tự push) không còn là ràng buộc hiện tại. Với yêu cầu mới, theo chỉ dẫn mới nhất của anh. Nếu soạn prompt Claude Code thì phải bắt đầu “Bước 1: Chẩn đoán”, sửa sau, find/replace chính xác, checklist rõ. Khi sửa tool, tăng `v=N` đồng bộ trong `index.html` và `manifest.json`, cập nhật phiên bản/ngày giờ hiển thị.

Anh yêu cầu **sau mỗi 15 thao tác làm việc có ý nghĩa**, lưu quá trình, quy trình có thể lặp lại và kinh nghiệm. Đếm các bước thực chất như khảo sát, chẩn đoán, backup, sửa, commit, push, xác minh; không tính tin nhắn trạng thái hay lần đọc lặp. Nếu việc kết thúc trước mốc 15, ghi kết quả và việc còn mở. Cập nhật mục “Nhật ký 15 thao tác” cuối file. Bộ đếm bắt đầu lại từ bản handoff này.

## Repo, file và bản đang chạy

- Repo có `.git`: `C:\Users\Admin\Projects\6.Tool\Trimble tool\Trimble tool\trimble-colorizer-pro\`.
- GitHub: `https://github.com/heodaigia1983/trimble-colorizer-pro`, nhánh `main`.
- Tool 1: `https://heodaigia1983.github.io/trimble-colorizer-pro/manifest.json`; trang v2.5: `https://heodaigia1983.github.io/trimble-colorizer-pro/index.html?v=30`.
- Tool 2 riêng ở `progress-tracker/`; không sửa khi việc chỉ liên quan Tool 1.
- IFC anh dùng: `C:\Users\Admin\Downloads\KC Gia Binh.ifc`. File này được chọn trong browser để tính thể tích; không có bằng chứng đã upload lên GitHub.
- Code v2.5 ở `088de34488f78034574e1ff1a178fed7814d49c5` (09/10/2026 15:13 UTC+7); file handoff được đưa vào repo tại `a0b482f` sau đó. Kiểm tra lại HEAD, branch, remote, status và manifest trước khi tiếp tục.

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

Anh đã thấy bảng **SL, KL (Tấn), V (m³), ρ (kg/m³)**. Mật độ thép mặc định là `7850 kg/m³` khi thiếu KL; tấn = kg/1000. Worker cộng thể tích có dấu từ tam giác IFC; về nguyên tắc phần rỗng của ống/hộp được trừ nếu lưới IFC biểu diễn đúng, không dùng thể tích hộp bao. Anh phản hồi “ok r em”, “hoạt động tốt” sau giai đoạn sửa thể tích. Chưa có bảng đối chiếu độc lập từng cấu kiện với Tekla, nên không khẳng định độ chính xác tuyệt đối.

Vấn đề kế tiếp: Save View hôm nay, hôm sau mở lại vẫn phải xem nhóm vàng/xanh bao nhiêu tấn. Trước v2.5 dữ liệu chỉ nằm trong `localStorage` cùng browser. Anh hỏi Google/GitHub; phương án đã chọn là lưu tóm tắt nhỏ trong **Trimble View description**. Nếu Trimble không giữ được thì mới tính tiếp Google/cloud khác. Chưa có Google endpoint hoặc quyền xác thực cho việc này.

## Kiến trúc và giới hạn hiện tại

`_colorLedger` trong `app.js` là Map khóa `modelId:runtimeId`, nguồn dữ liệu màu của Slot 3. Bảng gộp theo màu; checkbox cộng m³/tấn; click ô màu gọi `selectLedgerColorOnViewer()` để chọn cấu kiện trên model. Mỗi item có màu, ghi chú, mật độ và có thể có `volumeM3`.

`ifc-volume-worker.js` dùng `web-ifc@0.0.78`, cần IFC gốc có header `ISO-10303-21;` và đơn vị SI rõ ràng. `app.js` dùng `convertToObjectIds` đổi runtime IDs sang GUID IFC, worker tìm Express ID, tính thể tích mesh có dấu rồi đổi m³. Nếu không khớp GUID, không mở được IFC hay không rõ đơn vị thì báo thiếu, không lấy 0 làm thể tích thật. Cache thể tích theo từng object và ghi vào ledger cục bộ.

Save View v2.5: `rebuildColorLedgerSummary()` tính các nhóm. Khi **mọi nhóm** đủ m³/tấn và mô tả dài tối đa 240 ký tự, `makeViewDescription()` thêm `MCCQ1:` vào View description. Mỗi nhóm lưu màu, số lượng, thể tích, KL, mật độ và chữ ký tập `modelId:runtimeId` bằng base36. Độ phân giải m³ là 10⁻⁶, KL kg là 0.001. Luồng `createView` → `updateView` → `selectView` → `getView` đọc lại; chỉ log xác nhận cloud khi các số đọc lại khớp. Nếu không, cảnh báo và vẫn giữ bản cục bộ. Code ở đầu `app.js` và khoảng dòng 1180–1245.

Mở View: `getCurrentView`/`getView` lấy mô tả. Nếu máy không có localStorage, thử `getColoredObjects()` dựng lại ledger từ màu Viewer. Chỉ dùng m³/tấn lưu trên View khi màu, số lượng, mật độ và chữ ký cấu kiện khớp. Nếu model đổi hoặc runtime IDs thay đổi giữa các phiên, số có thể không hiện vì chữ ký lệch; đây là cơ chế tránh dùng sai số cũ. View tạo trước v2.5 chưa có `MCCQ1:`; cần nạp IFC và Save lại View mới. Mô tả 240 ký tự cũng giới hạn số nhóm màu có thể lưu cloud.

## Bằng chứng và việc chưa xác minh

Trước commit v2.5, `node --check app.js` và `git diff --check` đều exit 0; không thêm/chạy test suite. Backup trước v2.5: `backup/mcc-before-cloud-view-qty-v30-20261009` → `bf7a6c3`. Backup trước sửa handoff: `backup/mcc-before-handoff-20261009` → `088de34`; bản handoff đầu tiên: `backup/mcc-before-handoff-finalize-20261009` → `a0b482f`. Các tag đã push và đối chiếu trên origin. Workflow GitHub Pages `37903677651` deploy thành công; HTTP công khai trả manifest `v=30`, index `v2.5` và JS có `MCCQ1:`.

**Chưa được xác minh trong phiên Trimble thật:** Sau `updateView`, Trimble có giữ `description` không; mở View ngày khác có khôi phục ledger và m³/tấn không. Phiên Chrome Codex thấy khi đó không có tab Trimble đăng nhập. Không coi deploy Pages hay API docs là bằng chứng tính năng đã hoạt động trên model của anh.

Anh có thể kiểm tra: mở Tool 1 thấy v2.5; mở cùng model, chọn IFC gốc, tô vài nhóm, thấy m³/tấn; Save **View mới**; xem log “Đã xác nhận số liệu nhóm màu được lưu trên Trimble View”; đóng/mở lại View, click/tick màu và so tổng. Nếu log cảnh báo hoặc số mất, lấy log và trạng thái View, chẩn đoán description/quyền View/runtime IDs. Nếu phải chuyển Google, cần hỏi anh nơi lưu và quyền phù hợp; không nhúng token vào trang công khai.

## Việc khác còn mở

SPM `.xsr` và GUID assembly Tekla **chưa được chứng minh** khớp Viewer. GUID Tekla khác `IfcGloballyUniqueId`; trước khi làm matching phải chẩn đoán property Viewer thật. Handoff cũ 30/06 liệt kê lọc Assembly Position là chưa làm, nhưng Tool 1 hiện đã có UI và code cho mục này; thông tin cũ không còn chính xác. Không trộn nhiệm vụ SPM vào sửa m³/tấn nếu anh chưa yêu cầu.

## Quy trình tiếp tục và nhật ký 15 thao tác

Khi có yêu cầu mới: đọc file này → kiểm tra repo/phiên bản và bằng chứng mới từ anh → chẩn đoán → backup → sửa đúng phạm vi → tăng version/cache-bust → commit/push khi được giao → kiểm tra deploy → tách bạch kiểm tra công khai với thao tác Trimble thực. Không tuyên bố “đã chạy tốt trên Trimble” nếu chưa quan sát hoặc anh chưa xác nhận.

**Mốc 09/10/2026:** v2.5 đã lên Pages, backup đầy đủ; còn chờ thử Save/Open View trong Trimble. Kinh nghiệm: Viewer có thể thiếu KL và chỉ cấp IFC chuyển đổi, cần IFC gốc để tính m³. localStorage không đủ cho mở View trên máy khác. Lưu Trimble View cần đọc lại xác nhận; khi model đổi cần đối chiếu tập cấu kiện để tránh hiển thị số cũ. Bắt đầu đếm 15 thao tác mới sau bản handoff này; các mốc sau ghi nối tiếp tại đây.
