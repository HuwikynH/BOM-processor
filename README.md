# ⚡ BOM Processor - Công cụ Tính toán Linh kiện (Fab9 Edition)

BOM Processor là phần mềm chuyên dụng hỗ trợ bộ phận Báo giá và Mua hàng trong việc tự động hóa quy trình xử lý bảng vật tư (Bill of Materials). Hệ thống giúp chuẩn hóa dữ liệu, tự động phân loại linh kiện và tính toán chính xác tỷ lệ hao hụt (Attrition), nhằm tối ưu hóa thời gian và giảm thiểu sai sót thao tác thủ công.

## 🌟 Tóm tắt Tính năng

1. **Tự động trích xuất vùng dữ liệu:** Hệ thống tự động quét các hàng đầu tiên của file Excel để xác định bảng dữ liệu BOM chính, bỏ qua các thông tin ngoại vi (logo, metadata dự án) không liên quan.
2. **Phân loại linh kiện nâng cao (Tokenization):** Xử lý các chuỗi định dạng tên linh kiện phức tạp (ví dụ: `CAP-CERAM`, `RES_SMD_0805`, `CBL-POWER`) để phân tách và nhận diện chính xác chủng loại.
3. **Chỉnh sửa dữ liệu trực tiếp (Inline Edit):** Hỗ trợ cập nhật nhanh phân loại linh kiện bằng cách nhấp đúp chuột (Double-click) vào ô *Part Type*. Tỷ lệ Attrition tương ứng sẽ được hệ thống tính toán và cập nhật theo thời gian thực.
4. **Lọc và Sắp xếp đa chiều:** Hỗ trợ sắp xếp dữ liệu theo các trường thông tin, đồng thời cung cấp bộ lọc kết hợp thanh tìm kiếm để phân nhóm linh kiện theo tỷ lệ Attrition hoặc trạng thái nhận diện.
5. **Xuất dữ liệu theo giao diện (WYSIWYG Export):** Xuất chính xác tập dữ liệu đang được hiển thị trên giao diện ra định dạng Excel. Báo cáo đầu ra tự động tổng hợp số lượng gốc (BOM Qty) và số lượng thực tế cần mua (Final Qty = BOM Qty + Attrition).
6. **Xử lý chuyên biệt Cụm thành phẩm (Assemblies):** Nhận diện các linh kiện thuộc nhóm thành phẩm (Cable Assy, PCBA, Screw Mount, v.v.), tự động áp dụng mức Attrition 0% và bảo toàn định dạng tên gốc nhằm hỗ trợ công tác theo dõi.
7. **Giao diện UI Tiêu chuẩn Fab9:** Thiết kế tối giản, tập trung vào hiển thị dữ liệu theo chuẩn phần mềm doanh nghiệp, tuân thủ hệ thống nhận diện thương hiệu Fab9 (Light Mode).

## 🚀 Hướng dẫn Vận hành

1. Khởi chạy ứng dụng thông qua tệp `run_app.bat`.
2. Nhấn **[ 📂 Open File ]** và chọn tệp Excel BOM cần xử lý.
3. Nhấn **[ ▶ Process ]** để hệ thống thực thi thuật toán phân loại.
4. Rà soát kết quả trên lưới dữ liệu. Nhấp đúp (Double-click) vào cột *Part Type* để hiệu chỉnh phân loại linh kiện nếu cần thiết.
5. (Tùy chọn) Sử dụng bộ lọc hoặc thanh tìm kiếm để truy xuất các nhóm linh kiện cụ thể.
6. Nhấn **[ 📥 Export ]** để xuất dữ liệu đã xử lý thành tệp Excel mới, sẵn sàng cho quy trình Mua hàng.

## ⚙️ Cấu hình Tỷ lệ Hao hụt (Attrition Rules)

Quản trị viên có thể điều chỉnh tỷ lệ hao hụt tùy biến theo quy chuẩn của từng dự án hoặc yêu cầu từ đối tác:

- Nhấn vào tùy chọn **[ ⚙ Edit Rules ]** ở góc trên cùng bên phải giao diện.
- Nhấp đúp chuột vào giá trị % của loại linh kiện tương ứng để cập nhật mức tỷ lệ mới.
- Nhấn **Save** để hoàn tất. Hệ thống sẽ ngay lập tức áp dụng bộ quy tắc mới lên bảng dữ liệu hiện hành.

---
*Phát triển nội bộ - Hỗ trợ chuẩn hóa và tự động hóa quy trình xử lý dữ liệu BOM.*