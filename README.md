# ⚡ BOM Processor v1.5 – Công cụ Xử lý BOM (Fab9 Edition)

Phần mềm desktop tự động hóa quy trình xử lý Bill of Materials (BOM):
- Chuẩn hóa dữ liệu, phân loại linh kiện tự động
- Tính tỷ lệ hao hụt (Attrition) & số lượng mua thực tế
- Quản lý quy tắc & từ điển theo phiên bản (versioned) – không ghi đè file gốc
- Giao diện Light Mode chuẩn doanh nghiệp, hỗ trợ copy/dán, tìm kiếm, sort, filter
- Mở toàn màn hình (maximized) mặc định trên Windows

---

## 🚀 Cài đặt & Khởi chạy
```bash
# Cài thư viện (chạy 1 lần)
pip install customtkinter openpyxl pandas
# Hoặc dùng file batch
run_app.bat
```

---

## 🎯 Quy trình cơ bản
1. **📂 Open File** – Chọn file Excel BOM (`.xlsx`, `.XLSX`). Hệ thống tự tìm dòng header (tối đa 30 dòng đầu) bằng cách match tên cột.
2. **▶ Process** – Chạy thuật toán:  
   - Tách từ khóa từ cột *Description* / *Part Type*  
   - Nhận diện package (0402, 0805, SOD, SOT…), đơn vị (EA/PCS/IN/FT/M…)  
   - Tra bảng **Attrition Rules** → ra *Canonical Type*, *Attrition %*, *Final Qty*
3. **Rà soát trên lưới** – Cột hiển thị:  
   `#` | Part Type | Description | **MFR P/N** | **Internal P/N** | Qty | Attrition %
   - **Double-click cột Part Type** → sửa loại linh kiện → Attrition % tự cập nhật (dropdown gợi ý)
   - **Double-click cột Attrition %** (trong Editor) → sửa % trực tiếp (dropdown gợi ý)
4. **🔍 Filter / Search / Sort** –  
   - Ô *Filter*: tìm theo Description / Part Type  
   - Dropdown *Attrition*: lọc theo % cụ thể hoặc *Unknown* (tự cập nhật % thực tế trong dữ liệu)  
   - Click header cột → sort tăng/giảm (tất cả cột đều sort được)
4. **📋 Copy dữ liệu** –  
   - `Ctrl+C` → copy dòng/nhiều dòng (tab-separated, dán vào Excel tách cột)  
   - Right-click 1 ô → *Copy: …* → copy đúng ô đó
5. **📥 Export** – Xuất file Excel đúng thứ tự & filter đang xem (11 cột: Row, Original Part Type, Canonical Type, Description, **MFR Part Number**, **Internal Part Number**, BOM Qty, Unit, Attrition %, Attrition Rate, Final Qty)
6. **🔧 Chọn cột P/N thủ công** – Dropdown **Internal P/N Col** / **MFR P/N Col** trên thanh filter: chọn cột bất kỳ từ file BOM → cột P/N cập nhật ngay; chọn *Auto* → tự detect lại

---

## 🎨 Mã màu Attrition % (gradient tự động)
Mỗi % hao hụt có màu riêng, nội suy mượt từ **0% → 10%+**, không phụ thuộc giá trị cố định:
| % | Màu | Ý nghĩa |
|---|-----|---------|
| 0% | Xám | Không hao hụt (Assembly, Sheet Metal…) |
| 0.5% | Xanh ngọc | Connector/Housing, Terminal… |
| 1% | Xanh lá | Resistor/Cap/IC/LED/Relay/Switch/Transformer… |
| 2% | Xanh dương | Diode, Transistor, Inductor, Jumper… |
| 3% | Tím | Connector SMT, Oscillator… |
| 5% | Cam | Res/Cap 0201-0603, Wire, Jumper, Cable Tie, Heat Shrink… |
| 10%+ | Đỏ | Res/Cap 0201/0402, các loại nhỏ, rủi ro cao |

> Khi bạn sửa rule tạo ra % mới (ví dụ 1.5%, 4%, 7%) màu vẫn được tính gradient tự động – không cần cấu hình thêm.

---

## ⚙️ **Edit Rules** – Cửa sổ cấu hình Attrition
Mở bằng nút **⚙ Edit Rules** (góc trên phải). Mở toàn màn hình (maximized).

### Giao diện
- **Rules version** (dropdown) – chọn phiên bản đang dùng. *Lưu luôn tạo file mới `attrition_rules_v<timestamp>.json`, file gốc `attrition_rules.json` không bao giờ bị ghi đè.*
- **🗑 Delete** – xóa phiên bản đang chọn (không xóa được bản Original). Nếu xóa bản đang active → tự fallback về Original.
- 2 tab: **🔲 SMT / PCBA** | **🔌 Cable / Box** (màu chữ đen/trắng rõ ràng)
- **🔍 Filter** – lọc nhanh theo tên loại / package
- Bảng: *Component Type* | *Package / Size* | *Attrition %*
  - **Double-click 3 cột đều có dropdown gợi ý:**
    - **Component Type** → dropdown: tất cả loại từ rule active (RESISTOR, CAPACITOR, IC, CONNECTOR, WIRE, TERMINAL, HEAT_SHRINK, SCREW_NUT_WASHER, OTHER_SPECIAL…) – sửa xong rule tự move/rename
    - **Package** (SMT) → dropdown: 0201/0402/0603/0805/1206/SOD/SOT/SOP/QFP/QFN/BGA/(default)… Cable: *(all)*
    - **Attrition %** → dropdown: 0, 0.5, 1, 2, 3, 5, 10, 15, 20 (có thể gõ tùy ý)
  - Dòng đầu mỗi nhóm loại có nền nhấn nhẹ
- **＋ Add Row** – thêm dòng mới (mặc định `_default: 1%`)
- **💾 Save as New Version** – chỉ bật khi có thay đổi (dirty). Sau save: dropdown refresh, chọn bản mới, engine reload tự động.

### Logic tra Attrition
| Loại | Bảng dùng | Điều kiện |
|------|-----------|-----------|
| WIRE, TERMINAL, HEAT_SHRINK, CABLE_TIE, LABEL, POWER_MONITOR | **cable_box_rules** | Luôn |
| CONNECTOR | **cable_box_rules** | Description chứa `HOUSING` hoặc `CRIMP` |
| CONNECTOR | **smt_rules** | Các trường hợp còn lại (Header, RCPT trên PCBA…) |
| Có đơn vị đo độ dài (IN/FT/M/MM/CM/INCH/FEET) | **cable_box_rules** | Context cable |
| Các loại khác | **smt_rules** | Tra package (0201, 0402…) → fallback `_default` |

---

## 📚 **Edit Dictionary** – Cửa sổ từ điển viết tắt
Mở bằng nút **📚 Edit Dictionary** (giữa ⚙ Edit Rules và 📥 Export). Mở toàn màn hình (maximized).

### 6 Tab theo danh mục
| Tab | JSON Section / Group | Cột: Keyword → Maps To |
|-----|----------------------|------------------------|
| 🔲 SMT / PCBA | `keyword_to_type / SMT_COMPONENTS` | Từ viết tắt → Loại linh kiện (RESISTOR, CAPACITOR, IC…) |
| 🔌 Cable / Box | `keyword_to_type / CABLE_BOX_COMPONENTS` | Từ viết tắt → Loại (WIRE, CONNECTOR, TERMINAL…) |
| 📦 Misc / Added | `keyword_to_type / MISC_ADDED` | Các từ thêm lẻ (CE-*, SCR, CBL, CE-WIRE…) |
| 📐 Package | `package_keywords` | Từ khóa mô tả → Nhóm package (0805, SOD, SOT, QFN…) |
| 📏 Units | `unit_aliases` | Đơn vị file BOM → Canonical (**PCS** hoặc **LENGTH**) |
| 📋 Headers | `column_header_aliases` | Alias tựa cột → Role (description_col, part_type_col, quantity_col, unit_col, partnumber_col, mpn_col, internal_pn_col) |

### Chỉnh sửa
- Double-click ô **Keyword** → sửa tên (check trùng trong tab)
- Double-click ô **Maps To** → chọn từ **dropdown** (danh sách loại từ rule active / PCS-LENGTH / 7 role chuẩn) – state normal cho phép gõ tùy ý
- **＋ Add Row** – thêm dòng mới (mặc định: OTHER_SPECIAL / PCS / description_col)
- **🗑 Delete Row** – xóa dòng đang chọn (bật khi chọn dòng)
- **🔍 Filter** – lọc realtime theo keyword hoặc maps-to
- **Versioning** – tương tự Edit Rules: lưu bản mới, dropdown chọn phiên bản, 🗑 xóa bản (trừ Original)

> Sau Save / chuyển phiên bản: `attrition_engine` & `bom_reader` reload tự động – lần **Process** sau dùng từ điển mới.

---

## 🔢 Logic số lượng (Qty / Final Qty)
| Đơn vị | Final Qty | Ví dụ |
|--------|-----------|-------|
| **EA, PCS, PC, EACH, NOS, CÁI…** (đếm) | `ceil(BOM Qty × (1 + Attrition%))` | 100 × 1.05 → **105** ; 1.2 → **2** |
| **IN, FT, M, MM, CM, MTR, INCH, FEET…** (độ dài) | `round(BOM Qty × (1 + Attrition%), 3)` | 75 m × 1.05 → **78.75 m** |

> Cột **Qty** trên lưới = số lượng gốc từ file BOM (chưa cộng attrition). Cột **Final Qty** chỉ xuất ra file Excel.

---

## 🏷️ MFR P/N & Internal P/N – Tự nhận diện cột + Chọn thủ công
Hệ thống quét header bằng **chuẩn hóa** (bỏ dấu chấm, gạch, khoảng trắng, không phân biệt hoa/thường) → match alias:

| Role | Ví dụ alias nhận diện |
|------|-----------------------|
| **MFR P/N** | `MPN`, `MFR P/N`, `MFR PART NUMBER`, `MFG P/N`, `MANUFACTURER P/N`, `AMAT PART NUMBER`, `VENDOR P/N`, `SUPPLIER PN`… |
| **Internal P/N** | `ITEM NO`, `INTERNAL P/N`, `IPN`, `INTEL IPN`, `STOCKCODE`, `HOUSE P/N`, `OWN P/N`, `FAB9 P/N`, `PART NUMBER` (generic)… |
| **Generic Part Number** (fallback) | `PART NUMBER`, `P/N`, `PN`, `ITEM NUMBER`, `NUMBER`, `STOCKCODE`… |

> - Nếu file có **cả 2 cột riêng** → MFR P/N, Internal P/N hiển thị tách biệt
> - Nếu chỉ có **cột generic** (`Part Number` / `Item No`…) → dùng cho **cả 2 cột** (thường là mã nội bộ = MFR code)

### Dropdown chọn cột thủ công (thanh filter)
- Sau **Process**, 2 dropdown **Internal P/N Col** & **MFR P/N Col** tự nạp **tất cả tên cột thực tế** từ header file BOM
- Mặc định chọn cột auto-detect (`internal_pn_col`, `mpn_col`, fallback `partnumber_col`)
- Chọn cột bất kỳ → cột P/N trên bảng cập nhật ngay giá trị từ cột đó
- Chọn **Auto** → tự re-detect toàn bộ (re-process)

---

## 📋 Thanh trạng thái (Status Bar)
- Nền **vàng nhạt** (#fef9c3), chữ **đậm cỡ 15** – dễ chú ý
- Hiển thị: sẵn sàng / đang đọc / lỗi / số dòng copy / export thành công / rule/dictionary updated…
- Màu chữ đổi theo ngữ cảnh: xanh (thành công), cam (chưa lưu), đỏ (lỗi), đen (thông thường)

---

## 📁 Quản lý phiên bản (Versioning) – Tổng quan
| Loại | File gốc | File version | Con trỏ active |
|------|----------|--------------|----------------|
| Attrition Rules | `attrition_rules.json` | `attrition_rules_v<YYYY-MM-DD_HHMMSS>.json` | `_active_rules.json` |
| Keyword Dictionary | `component_dictionary.json` | `component_dictionary_v<YYYY-MM-DD_HHMMSS>.json` | `_active_dictionary.json` |

- **Không bao giờ ghi đè file gốc** – an toàn rollback
- Dropdown trong từng Editor liệt kê: `Original (default)` → `v2026-08-25 14:30:15`…
- Nút 🗑 chỉ bật ở bản version, xóa xong refresh dropdown, fallback về Original nếu cần
- Thư mục chứa file version + pointer đã được thêm vào `.gitignore`

---

## 📦 Cấu trúc file quan trọng
```
bom_processor/
├── main.py                   # Entry point
├── app.py                    # UI chính, table, process, export, copy, status, P/N column selector
├── bom_reader.py             # Đọc Excel, detect header (chuẩn hóa), trả dict rows + header values
├── attrition_engine.py       # Phân loại, tra attrition, gradient màu, reload hot
├── attrition_editor.py       # Editor Attrition Rules (versioned, 3 cột dropdown)
├── dictionary_editor.py      # Editor Keyword Dictionary (versioned, 6 tab, dropdown Maps To)
├── file_versions.py          # Helper chung versioning (load/save/list/delete)
├── component_dictionary.json # Từ điển gốc (keyword, package, unit, header alias)
├── attrition_rules.json      # Rule attrition gốc (smt + cable_box + assembly_context)
├── run_app.bat               # Khởi chạy nhanh Windows
└── README.md                 # Tài liệu này
```

---

## 🛠️ Mở rộng / Tuỳ biến
- **Thêm alias cột mới**: sửa `component_dictionary.json → column_header_aliases` → restart app
- **Thêm loại linh kiện mới**:  
  1. Thêm keyword vào `keyword_to_type` (tab SMT/Cable/Misc)  
  2. Thêm entry vào `attrition_rules.json` (smt_rules / cable_box_rules) với `_default` + package con nếu cần  
  3. Save version → reload
- **Sửa màu gradient**: chỉnh `_ATT_ANCHORS` trong `app.py` (list tuple % + RGB)
- **Thêm đơn vị độ dài**: bổ sung vào `unit_aliases` (→ LENGTH) và `assembly_context.length_units` trong rule
- **Thêm từ khóa viết tắt**: bổ sung vào `keyword_to_type` (SOM→IC, DIO/ZNR→DIODE, ICS→IC…)

---

## ⚠️ Lưu ý & Troubleshooting
- **File Excel có style lỗi** → đã có hotfix bỏ qua `styles.xml` (openpyxl `apply_stylesheet = None`)
- **Không nhận header** → kiểm tra dòng 1–30 có tên cột khớp alias không; mở `component_dictionary.json` bổ sung alias
- **Attrition % sai** → kiểm tra: Part Type / Description có từ khóa đúng không? Package detect đúng không? Đơn vị có bị hiểu nhầm (EA vs M) không?
- **Internal P/N trống** → file có thể chỉ có 1 cột Part Number chung → cả 2 cột dùng chung giá trị đó; dùng dropdown chọn cột đúng
- **Copy không dán được** → đảm bảo chọn dòng (hoặc ô) trước khi `Ctrl+C` / right-click
- **Đổi cột P/N không cập nhật** → đảm bảo đã Process, header đúng; nếu vẫn lỗi chọn *Auto* để re-process

---

## 📜 Phiên bản
| Phiên bản | Ngày | Thay đổi chính |
|-----------|------|----------------|
| **v1.5**  | 2026-08-27 | Dropdown chọn cột P/N thủ công (Internal/MFR); Fix bug đổi cột P/N; 3 cột Editor Rule dropdown (Component Type/Package/Attrition%); Maps To dropdown Dictionary Editor; Add Row cả 2 Editor; Mở toàn màn hình mặc định; Thêm từ khóa SOM/ICS→IC, DIO/ZNR→DIODE; Cả 3 cửa sổ maximized |
| v1.2      | 2026-08-26 | Gradient màu Attrition động; Editor Dictionary 6-tab versioned; Copy Ctrl+C/Right-click; MFR/Internal P/N detection nâng cấp; Status bar nổi bật; bỏ legend % cố định |
| v1.1      | 2026-08-24 | Editor Rule versioned + dropdown + delete; Internal P/N column; Qty column; rounding logic length vs PCS |
| v1.0      | 2026-08-?? | Release nội bộ: process, inline edit Part Type, export WYSIWYG, assembly 0% attrition |

---

*Phát triển nội bộ – Hỗ trợ chuẩn hóa & tự động hóa quy trình xử lý dữ liệu BOM – Fab9 Engineering*