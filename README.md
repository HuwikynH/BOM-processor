# ⚡ BOM Processor v1.5 – Hướng Dẫn Sử Dụng Chi Tiết (User Manual)

Chào mừng bạn đến với tài liệu hướng dẫn sử dụng **BOM Processor v1.5**. Ứng dụng này được thiết kế nhằm tự động hóa quy trình phân tích bảng kê linh kiện (Bill of Materials - BOM), nhận diện loại linh kiện và tính toán tự động số lượng hao hụt (attrition) để phục vụ cho việc nhập/xuất kho chính xác.

---

## 1. 🚀 Khởi động ứng dụng

Để chạy ứng dụng, bạn có thể thực hiện 1 trong 2 cách sau:
- **Cách 1**: Bấm đúp vào file `run_app.bat` (nếu dùng Windows).
- **Cách 2**: Mở terminal/cmd ở thư mục chứa ứng dụng và chạy lệnh:
  ```bash
  pip install customtkinter openpyxl pandas
  python main.py
  ```

---

## 2. 🎯 Quy trình xử lý BOM cơ bản

Một phiên làm việc thông thường sẽ trải qua 3 bước:

1. **📂 Open File**: Bấm nút này để chọn file BOM định dạng Excel (`.xlsx`, `.XLSX`). Ứng dụng sẽ đọc 30 dòng đầu tiên để tự động dò tìm vị trí các cột dữ liệu quan trọng (Description, Part Type, Qty...).
2. **▶ Process**: Sau khi file được tải xong, nút Process sẽ sáng lên. Bấm vào đây để ứng dụng bắt đầu phân tích từng dòng:
   - Nhận diện loại linh kiện (Resistor, Capacitor, Wire...).
   - Nhận diện kích thước/chuẩn đóng gói (Package như 0402, 0603, SOP, QFN...).
   - Áp dụng quy tắc tính % hao hụt (Attrition %).
3. **📥 Export**: Cuối cùng, bấm nút Export để xuất kết quả bảng đã được xử lý ra một file Excel mới. 

---

## 3. 🖥️ Khám phá Giao diện chính (Main UI) & Bảng dữ liệu

### Cấu trúc bảng kết quả
Bảng dữ liệu (Treeview) hiển thị tổng quan các dòng BOM đã qua xử lý với các cột sau:
- **#**: Số thứ tự dòng trên giao diện.
- **Part Type**: Loại linh kiện đã được ứng dụng chuẩn hóa (Ví dụ: `Resistor [0402]`, `Wire / Cable`).
- **Description (original)**: Mô tả gốc từ file BOM chưa qua chỉnh sửa.
- **MFR P/N**: Mã linh kiện của nhà sản xuất.
- **Internal P/N**: Mã linh kiện nội bộ.
- **Qty**: Số lượng gốc từ BOM (chưa tính hao hụt).
- **Attrition %**: Tỷ lệ hao hụt được ứng dụng cấp phát. **Mỗi % hao hụt có một màu sắc hiển thị khác nhau (màu gradient tự động nội suy).**

### Thanh trạng thái (Status Bar)
Nằm ở dưới cùng của cửa sổ ứng dụng (nền màu vàng nhạt). Đây là nơi hiển thị trạng thái hiện tại: File đang mở, số lượng linh kiện không nhận dạng được (unknown), thông báo lỗi, trạng thái copy dữ liệu, hoặc thông báo lưu thành công.

---

## 4. 🛠 Các tính năng tương tác và Tùy biến dữ liệu

Từ nhỏ đến lớn, bạn có rất nhiều công cụ để tinh chỉnh dữ liệu trực tiếp trên bảng hiển thị:

### 4.1. Tìm kiếm và Lọc (Filter & Search)
- **Ô Filter (Tìm kiếm văn bản)**: Nhập từ khóa (tên, thông số, mô tả...). Bảng sẽ lọc trực tiếp (real-time) ra các dòng có chứa từ khóa trong cột *Description* hoặc *Part Type*.
- **Dropdown Attrition**: Bấm vào đây để lọc các linh kiện theo một % hao hụt cụ thể, hoặc lọc ra những linh kiện chưa nhận diện được (`Unknown`).

### 4.2. Sắp xếp (Sorting)
Bạn có thể sắp xếp tăng/giảm dần (A-Z, Z-A) bằng cách **Click trực tiếp vào tiêu đề của bất kỳ cột nào**. Click lần nữa để đảo ngược chiều sắp xếp.

### 4.3. Chọn cột hiển thị Mã linh kiện (MFR P/N & Internal P/N)
Hai cột **MFR P/N** và **Internal P/N** có biểu tượng **▼** cạnh tiêu đề.
- **Click vào tiêu đề có biểu tượng ▼**, một menu thả xuống sẽ hiện ra liệt kê tất cả các tên cột có trong file BOM gốc của bạn.
- Bạn có thể chủ động chọn cột dữ liệu gốc bạn muốn ánh xạ vào, nội dung cột sẽ thay đổi tức thì. 
- Chọn **Auto (tự động detect)** để đưa ứng dụng về chế độ tự tìm cột tối ưu nhất.

### 4.4. Chỉnh sửa thủ công Loại linh kiện (Inline Edit Part Type)
Nếu ứng dụng nhận dạng sai một linh kiện, bạn có thể sửa trực tiếp trên bảng:
1. **Double-click** (nhấp đúp) vào ô **Part Type** của dòng linh kiện đó.
2. Một hộp thoại tìm kiếm (Searchable Combobox) sẽ xuất hiện. 
3. Bạn có thể gõ từ khóa (ví dụ: `Capacitor`) và dùng phím mũi tên để chọn loại/kích thước linh kiện đúng. 
4. Ngay khi bấm **Enter**, hệ thống sẽ tự động tính toán lại % hao hụt cho dòng đó dựa theo loại bạn vừa gán.

### 4.5. Sao chép dữ liệu (Copy Data)
- **Sao chép nhiều dòng**: Chọn một hoặc nhiều dòng (giữ `Shift` hoặc `Ctrl`), sau đó bấm `Ctrl + C`. Dữ liệu sẽ được copy dưới định dạng cách nhau bởi dấu Tab, có thể Paste trực tiếp vào file Excel cực kỳ thẳng hàng.
- **Sao chép 1 ô (Cell)**: **Click chuột phải** vào bất kỳ ô nào bạn muốn, chọn lệnh **Copy...** hiện ra để sao chép duy nhất nội dung của ô đó.

---

## 5. 📥 Xuất dữ liệu (Export)

Tính năng **Export** tuân theo nguyên tắc "Thấy gì xuất nấy" (WYSIWYG). File Excel sinh ra sẽ phản ánh **chính xác 7 cột bạn đang thấy trên màn hình**, kể cả khi bạn đã dùng Filter, đổi cột P/N, hay sửa Part Type.

Thứ tự xuất sẽ là: `#`, `Part Type`, `Description (original)`, `MFR P/N`, `Internal P/N`, `Qty`, `Attrition %`.

*(Chú ý: Nếu bạn có tùy chọn đổi cột MFR P/N hay Internal P/N trên giao diện, tên tiêu đề xuất ra file Excel cũng sẽ tự động mang tên cột bạn đã chọn để đảm bảo tính đồng bộ).*

---

## 6. ⚙️ Quản lý Quy tắc Hao hụt (Edit Rules)

Nút **Edit Rules** sẽ mở một cửa sổ mới giúp bạn thay đổi tỷ lệ % hao hụt cho từng nhóm linh kiện.

- **2 Tab phân chia logic**: 
  - **SMT / PCBA**: Chứa quy tắc hao hụt cho các linh kiện điện tử dán/cắm (thường được xác định thông qua cả Component Type và Package).
  - **Cable / Box**: Chứa quy tắc hao hụt cho nhóm dây cáp, lắp ráp cơ khí, vỏ hộp, vỏ gen...
- **Sửa dữ liệu**: Bạn có thể click đúp vào các cột `Component Type`, `Package / Size` hoặc `Attrition %` để sửa trực tiếp thông qua các Dropdown tiện lợi.
- **Thêm/Xóa dòng**: Hỗ trợ nút `+ Add Row` để định nghĩa loại linh kiện mới.
- **Lưu phiên bản (Versioning)**: Khi bấm `Save as New Version`, hệ thống **không ghi đè** file gốc mà sẽ tạo ra một file cấu hình mới đánh dấu bằng thời gian. Bạn có thể tự do rollback lại các phiên bản cũ qua menu thả xuống ở góc trên.

---

## 7. 📚 Quản lý Từ điển (Edit Dictionary)

Nút **Edit Dictionary** chứa bộ não dịch thuật của ứng dụng, giúp biến những đoạn chuỗi viết tắt (như `RES`, `CBL`, `0805`, `EA`) thành thuật ngữ chuẩn. Có 6 Tab quan trọng:

1. **SMT / PCBA (Keyword)**: Gắn từ viết tắt với linh kiện bảng mạch (VD: `RES` -> `RESISTOR`).
2. **Cable / Box (Keyword)**: Gắn từ viết tắt với linh kiện cáp/cơ khí (VD: `CBL` -> `WIRE`).
3. **Misc / Added**: Các từ vựng chung.
4. **Package**: Gắn chuỗi thông số với chuẩn kích thước đóng gói (VD: `0805SMD` -> `0805`).
5. **Units**: Xác định xem một loại đơn vị tính là dạng đếm số lượng (PCS) hay dạng đo chiều dài (LENGTH). Tùy vào loại đơn vị mà cách tính toán làm tròn (Rounding logic) hao hụt sẽ khác nhau.
6. **Headers**: Định nghĩa các từ khóa mà hệ thống dùng để đi "săn" tên cột trong file BOM Excel (ví dụ thấy cột có chữ `MFG P/N`, nó sẽ hiểu đó là `mfr_pn_col`).

Tất cả các thay đổi trong từ điển cũng hỗ trợ **Lưu phiên bản (Versioning)** y hệt như Edit Rules, giữ cho hệ thống của bạn an toàn trước những thay đổi sai lầm.

---

## 8. ⚠️ Xử lý sự cố (Troubleshooting)

- **Lỗi không đọc được file Excel**: Đảm bảo file chưa bị khóa bởi phần mềm khác. Ứng dụng đã được tích hợp bộ lọc chặn lỗi định dạng (`styles.xml`) từ các phần mềm ERP xuất ra, nên có thể đọc hầu hết các chuẩn Excel 2007 trở lên.
- **Không có dữ liệu hiện ra sau khi Process**: Đảm bảo rằng file Excel gốc của bạn có chứa cột mang tên "Description" hoặc các biến thể của nó ở trong vòng 30 dòng đầu tiên.
- **Internal P/N hoặc MFR P/N bị trống**: Bạn có thể dùng tính năng Click vào tiêu đề (biểu tượng ▼) để chỉ định ép buộc cột dữ liệu.
- **Linh kiện ra chữ "???" hoặc % hao hụt là 0**: Ứng dụng không tìm thấy từ khóa trong Từ Điển. Hãy dùng nút `Edit Dictionary` để bổ sung từ viết tắt đang có trong Description của linh kiện đó. Hoặc nhấp đúp vào ô Part Type để chỉ định thủ công!

---
*BOM Processor v1.5 - Phát triển nội bộ cho Fab9 Engineering.*