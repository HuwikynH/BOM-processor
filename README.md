# QUY TRÌNH ĐÀO TẠO SỬ DỤNG BOM PROCESSOR

| Phiên bản | Ngày ban hành | Tạo mới / Sửa đổi | Kiểm tra | Chấp nhận | Diễn giải |
| :---: | :---: | :---: | :---: | :---: | :--- |
| v1.5 | 9/7/2026 | | | | Phát hành tài liệu hướng dẫn sử dụng chi tiết toàn tập từ   |

---

## MỤC LỤC 

1.0 [GIỚI THIỆU TỔNG QUAN](#10-giới-thiệu-tổng-quan)

2.0 [KHỞI ĐỘNG VÀ NẠP DỮ LIỆU TỪ  ](#20-khởi-động-và-nạp-dữ-liệu-từ- )

3.0 [THAO TÁC TRÊN BẢNG DỮ LIỆU (MAIN GRID)](#30-thao-tác-trên-bảng-dữ-liệu-main-grid)

4.0 [TÙY BIẾN CỘT VÀ XUẤT/SAO CHÉP DỮ LIỆU](#40-tùy-biến-cột-và-xuấtsao-chép-dữ-liệu)

5.0 [TRÌNH QUẢN LÝ QUY TẮC HAO HỤT (RULES EDITOR)](#50-trình-quản-lý-quy-tắc-hao-hụt-rules-editor)

6.0 [TRÌNH QUẢN LÝ TỪ ĐIỂN (DICTIONARY EDITOR)](#60-trình-quản-lý-từ-điển-dictionary-editor)

7.0 [LOGIC TOÁN HỌC VÀ LÀM TRÒN (MATH & ROUNDING)](#70-logic-toán-học-và-làm-tròn-math--rounding)

8.0 [XỬ LÝ SỰ CỐ THƯỜNG GẶP (TROUBLESHOOTING)](#80-xử-lý-sự-cố-thường-gặp-troubleshooting)

---

## 1.0 GIỚI THIỆU TỔNG QUAN

**BOM Processor** là phần mềm desktop chuyên dụng được thiết kế nhằm tự động hóa quy trình phân tích Bảng kê linh kiện (Bill of Materials - BOM).
Thay vì phải dùng mắt đọc từng dòng và dùng hàm Excel thủ công, hệ thống sử dụng thuật toán dò tìm từ khóa để:
- Tự động chuẩn hóa dữ liệu thô từ nhiều nguồn ERP khác nhau.
- Nhận diện chính xác loại linh kiện (IC, Resistor, Wire...) và kích thước đóng gói (0402, QFN...).
- Tự động gán tỷ lệ hao hụt (Attrition) an toàn cho sản xuất.
- Trực quan hóa dữ liệu bằng hệ thống mã màu.

---

## 2.0 KHỞI ĐỘNG VÀ NẠP DỮ LIỆU TỪ  

### 2.1 Khởi chạy ứng dụng
1. Nhấp đúp vào file `.exe` đã đóng gói hoặc file `run_app.bat`.
2. Ứng dụng sẽ tự động mở ở chế độ **Toàn màn hình (Maximized)** để bạn có không gian quan sát bảng dữ liệu tốt nhất. Thanh trạng thái (màu vàng nhạt) ở dưới cùng sẽ hiển thị "Ready".

### 2.2 Nạp file Excel (Open File)
1. Nhấp vào nút **"📂 Open File"** ở góc trên cùng bên trái.
2. Chọn file BOM định dạng Excel (`.xlsx`, `.XLSX`). 
   - *Lưu ý:* Ứng dụng được tích hợp bộ lọc chặn lỗi định dạng (`styles.xml`), giúp nó đọc mượt mà ngay cả những file Excel bị lỗi style xuất ra từ các hệ thống ERP đời cũ.
3. **Logic dò tìm tiêu đề (Header Detection)**: Ứng dụng không yêu cầu file BOM của bạn phải bắt đầu từ dòng số 1. Nó sẽ quét thông minh từ dòng 1 đến dòng 30. Khi thấy một dòng có chứa các từ khóa như `Description`, `Qty`, `Part Number`, nó sẽ chốt đó là dòng tiêu đề và bỏ qua các dòng logo/thông tin công ty phía trên.
4. Thanh trạng thái sẽ báo số lượng dòng dữ liệu đã tải thành công.

### 2.3 Phân tích dữ liệu (Process)
1. Sau khi nạp file, nút **"▶ Process"** sẽ sáng lên. Nhấp vào nút này.
2. Ứng dụng sẽ chạy thuật toán lõi:
   - Tách chuỗi văn bản từ cột `Description` và `Part Type` gốc.
   - Quét qua **Từ điển (Dictionary)** để dịch các từ viết tắt.
   - Nhận diện Package (ví dụ thấy `0805SMD` sẽ hiểu là `0805`).
   - Tra bảng **Quy tắc (Rules)** để cấp phát % hao hụt.
3. Toàn bộ bảng dữ liệu sẽ được hiển thị ngay lập tức với hệ thống màu sắc rực rỡ.

---

## 3.0 THAO TÁC TRÊN BẢNG DỮ LIỆU (MAIN GRID)

Bảng chính gồm 7 cột chuẩn hóa: `#`, `Part Type`, `Description (original)`, `MFR P/N`, `Internal P/N`, `Qty`, và `Attrition %`.

### 3.1 Hệ thống mã màu (Gradient Attrition)
Ứng dụng tự động nội suy màu nền cho tỷ lệ % hao hụt để người dùng nhận diện rủi ro bằng mắt thường:
- **0% (Xám)**: Không hao hụt (Linh kiện lắp ráp cơ khí, Sheet metal).
- **0.5% - 1% (Xanh lục)**: Hao hụt thấp (IC, Res/Cap lớn).
- **2% - 5% (Xanh dương - Cam)**: Hao hụt trung bình (Diode, Dây điện, Res/Cap nhỏ).
- **10% trở lên (Đỏ)**: Hao hụt cao (Linh kiện siêu nhỏ 0201, 0402).

### 3.2 Sửa loại linh kiện siêu tốc (Inline Edit)
Nếu ứng dụng phân loại sai do Description quá lạ, bạn KHÔNG cần sửa file Excel:
1. **Nhấp đúp chuột (Double-click)** vào ô `Part Type` của linh kiện bị sai.
2. Một hộp thoại tìm kiếm (Searchable Combobox) thông minh sẽ xuất hiện. Nếu nó che khuất dữ liệu bên dưới, nó sẽ tự động nhận diện màn hình và **xổ ngược lên trên**.
3. Bạn có thể gõ từ khóa (ví dụ: `Capacitor`). Dropdown sẽ gợi ý các loại Capacitor cùng kích thước (vd: `Capacitor [0402]`, `Capacitor [0603]`).
4. Dùng phím mũi tên Lên/Xuống để chọn và nhấn **Enter**. 
5. Ngay lập tức, `Part Type` được cập nhật, và `% Hao hụt` tự động được tính toán lại theo quy tắc của loại bạn vừa chọn. Thanh trạng thái báo màu xanh lá cây xác nhận thành công.

### 3.3 Sắp xếp dữ liệu (Sorting)
- Click trực tiếp vào tiêu đề của bất kỳ cột nào để sắp xếp (  hoặc Tăng dần).
- Click thêm lần nữa vào cột đó để đảo ngược thứ tự (Z-A hoặc Giảm dần).
- Sắp xếp hoạt động với mọi cột, kể cả cột số lượng (`Qty`) hay phần trăm (`Attrition %`).

### 3.4 Bộ lọc đa năng (Filter)
- **Ô Filter Text**: Nằm ở thanh công cụ phía trên. Khi bạn gõ chữ (ví dụ: `connector`), bảng sẽ lọc **ngay lập tức theo thời gian thực (real-time)** ra các dòng có chứa từ "connector" trong mô tả.
- **Dropdown % Hao hụt**: Kế bên ô Filter Text. Bạn có thể chọn lọc riêng các linh kiện có `%` cụ thể (ví dụ: `5.0%`), hoặc chọn `Unknown` để lọc ra toàn bộ các linh kiện mà ứng dụng chưa hiểu để bạn xử lý thủ công. Chọn `All` để hiện lại toàn bộ.

---

## 4.0 TÙY BIẾN CỘT VÀ XUẤT/SAO CHÉP DỮ LIỆU

### 4.1 Tùy biến cột Mã linh kiện (MFR & Internal P/N)
Đôi khi hệ thống ERP xuất ra file có tên cột P/N rất dị, hoặc trộn lẫn mã nội bộ và mã nhà sản xuất.
1. Nhìn lên tiêu đề cột `MFR P/N` hoặc `Internal P/N`, bạn sẽ thấy một biểu tượng **dấu mũi tên thả xuống (▼)**.
2. Click vào biểu tượng ▼ này. Một danh sách chứa **toàn bộ tên cột có trong file Excel gốc** của bạn sẽ hiện ra.
3. Chỉ cần chọn đúng cột bạn muốn gán. Dữ liệu trên bảng sẽ thay đổi ngay lập tức để lấy giá trị từ cột gốc bạn vừa chọn.
4. Chọn **Auto** để trả quyền nhận diện tự động lại cho ứng dụng.

### 4.2 Sao chép (Copy)
- **Copy nguyên dòng/nhiều dòng**: Giữ `Ctrl` hoặc `Shift` để chọn nhiều dòng. Nhấn `Ctrl+C`. Dữ liệu sẽ vào Clipboard ngăn cách bằng phím Tab (hoàn hảo để paste vào Excel).
- **Copy một ô duy nhất**: Nhấp chuột phải (Right-click) vào chính xác ô bạn muốn copy. Một menu nhỏ ghi `Copy: [Nội dung ô]` sẽ hiện ra. Nhấp vào để copy.

### 4.3 Xuất file (Export)
1. Nhấp nút **"📥 Export"**.
2. File Excel xuất ra tuân thủ nghiêm ngặt nguyên tắc **Thấy gì xuất nấy (WYSIWYG)**. Nó sẽ chỉ xuất đúng **7 cột** đang hiển thị trên giao diện.
3. Mọi thao tác Lọc (Filter), Sắp xếp (Sort), Đổi tên cột P/N, hay Chỉnh sửa thủ công Part Type của bạn đều được giữ nguyên 100% trong file Excel xuất ra. (Lưu ý: Tiêu đề cột P/N trong Excel xuất ra sẽ tự động đổi thành tên cột gốc mà bạn đã chọn qua nút ▼).

---

## 5.0 TRÌNH QUẢN LÝ QUY TẮC HAO HỤT (RULES EDITOR)

Trình quản lý này chứa "Luật" cấp phát % hao hụt.
- Nhấp nút **"⚙ Edit Rules"** để mở. Cửa sổ này nổi trên ứng dụng chính (không chặn thao tác của ứng dụng chính).
- Có 2 Tab:
  - **SMT / PCBA**: Dành cho linh kiện bo mạch (Resistor, IC, Diode...). Logic: Xét Component Type -> Xét Package -> Ra %.
  - **Cable / Box**: Dành cho dây dẫn, cơ khí, ốc vít. Logic: Trực tiếp lấy % theo Component Type.

### Hướng dẫn thao tác:
1. **Sửa rule**: Nhấp đúp vào bất kỳ ô nào (`Component Type`, `Package`, hoặc `Attrition %`). Một Dropdown sẽ xổ ra để bạn chọn chuẩn xác, tránh gõ sai chính tả.
2. **Thêm rule**: Nhấp nút **"+ Add Row"** ở dưới cùng để thêm một dòng trắng, sau đó click đúp để gán giá trị.
3. **Lưu phiên bản (Versioning)**: Khi bạn có thay đổi, nút **"💾 Save as New Version"** sẽ sáng lên. Hệ thống **không bao giờ ghi đè file gốc**. Nó sẽ tạo ra một file có gắn ngày tháng (ví dụ: `_v2026-09-07_0930.json`).
4. Khung Dropdown trên cùng cho phép bạn "quay ngược thời gian" (Rollback) về bất kỳ phiên bản luật nào trước đây, hoặc về file "Original". Nút 🗑 bên cạnh dùng để xóa phiên bản hiện tại (không thể xóa Original).

---

## 6.0 TRÌNH QUẢN LÝ TỪ ĐIỂN (DICTIONARY EDITOR)

Trình quản lý này là "Bộ não dịch thuật" của ứng dụng.
- Nhấp nút **"📚 Edit Dictionary"** để mở. Cửa sổ cũng nổi song song với app chính.
- Gồm **6 Tab** quản lý 6 loại dịch thuật khác nhau:
  1. **SMT / PCBA**: Gắn từ khóa trong mô tả (vd: `RES`, `CAP`) thành linh kiện chuẩn (`RESISTOR`, `CAPACITOR`).
  2. **Cable / Box**: Gắn từ khóa cơ khí (vd: `CBL`, `SCREW`) thành `WIRE`, `SCREW_NUT_WASHER`.
  3. **Misc / Added**: Các từ khóa phụ cần gộp chung.
  4. **Package**: Dịch các cụm từ dị thành chuẩn đóng gói. (vd: `0805SMD`, `0805_1%` -> dịch thành chuẩn `0805`). Điều này giúp Rule SMT bắt trúng kích thước.
  5. **Units**: Xác định xem một cụm từ đơn vị (`EA`, `PCS`, `M`, `FEET`) thuộc nhóm Đếm số (`PCS`) hay nhóm Đo chiều dài (`LENGTH`).
  6. **Headers**: Dạy cho thuật toán quét file Excel biết những từ nào (vd: `MFG P/N`, `VENDOR PN`) được coi là cột Mã nhà sản xuất.

### Hướng dẫn thao tác:
Tương tự Rules Editor, nhấp đúp vào ô `Maps To` để mở Dropdown chọn loại chuẩn. Hỗ trợ hệ thống **Versioning** hoàn chỉnh (Save as New Version, Rollback, Delete).

---

## 7.0 LOGIC TOÁN HỌC VÀ LÀM TRÒN (MATH & ROUNDING)

BOM Processor có thuật toán tự động làm tròn `Final Qty` (Số lượng cuối cùng sau hao hụt) khi xuất ra Excel, dựa trên bản chất đơn vị:
1. **Nhóm Đếm số (PCS, EA, Cái, Bộ)**: Không thể mua nửa cái IC. Do đó thuật toán dùng `Ceiling` (Làm tròn LÊN số nguyên gần nhất). 
   - *Ví dụ:* 100 cái x hao hụt 1.5% = 101.5 -> App sẽ tự động làm tròn lên thành **102**.
2. **Nhóm Chiều dài (M, FT, INCH)**: Có thể cắt lẻ dây điện. Thuật toán giữ lại tối đa 3 chữ số thập phân.
   - *Ví dụ:* 25 mét x hao hụt 5% = 26.25 -> App xuất ra đúng **26.25**.

---

## 8.0 XỬ LÝ SỰ CỐ THƯỜNG GẶP (TROUBLESHOOTING)

1. **Lỗi "Không thấy dữ liệu sau khi Process"**:
   - Đảm bảo file Excel gốc của bạn có chứa cột mang tên "Description" hoặc các biến thể của nó nằm trong phạm vi 30 dòng đầu tiên. Nếu tiêu đề nằm ở dòng 35, hãy xóa bớt các dòng trống/logo ở trên.
2. **Cột MFR P/N và Internal P/N giống hệt nhau**:
   - Trường hợp file BOM chỉ cung cấp 1 cột "Part Number" chung chung, ứng dụng sẽ điền dữ liệu cột đó cho cả MFR và Internal. Bạn hãy click vào biểu tượng ▼ trên tiêu đề để trỏ tay về cột bạn muốn.
3. **Phần trăm hao hụt hiện 0% hoặc Part Type hiện "???"**:
   - Từ khóa mô tả linh kiện này chưa có trong Từ điển. Hãy mở **Edit Dictionary**, thêm từ khóa đó vào tab SMT hoặc Cable. Lần Process sau ứng dụng sẽ thông minh hơn.
   - Cách nhanh nhất: Double-click thẳng vào ô Part Type đó và chỉ định tay cho nhanh!
