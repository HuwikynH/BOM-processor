# ⚡ BOM Processor 

Chào bạn! Đây là công cụ Desktop App ra đời để giải quyết một nỗi đau rất cụ thể của kỹ sư/báo giá: **Dọn dẹp và phân tích file BOM (Bill of Materials) khổng lồ**.

Thay vì bạn phải căng mắt nhìn hàng nghìn dòng Excel, lẩm nhẩm xem "RES D,0402" là cái gì, rồi điền tay tỷ lệ hao hụt (attrition) từng món một... BOM Processor sẽ làm chuyện đó trong chưa tới 1 giây.

## Tóm tắt những gì app làm được:
1. **Nuốt trọn mọi file BOM (dạng .xlsx)** — Dù file của bạn xài tên cột là `Part Type`, `Category` hay `Item Desc`, app cũng tự động dò tìm ra đúng cột cần quét.
2. **Từ điển thông minh** — Tự động nhận diện những từ viết tắt hầm hố (như `RES` -> Resistor, `CAP` -> Capacitor) và bắt được cả kích thước linh kiện (như `0402`, `QFN`).
3. **Áp tỷ lệ hao hụt tự động** — Đẩy ngay tỷ lệ % khấu hao chính xác (theo luật SMT hoặc Cable/Box).
4. **Chỉnh sửa Rules tại chỗ** — Bảng khấu hao có thể được bạn update nóng ngay trên app (Double-click là sửa, không cần code).

---

## 🛠 Cách dùng siêu tốc

> 💡 **Thông minh & Tự động:** App sẽ tự động dò tìm trong 30 hàng đầu tiên của file Excel để tìm xem hàng nào chứa tiêu đề bảng (chứa chữ `Description`, `Item Desc`...). 
> Vậy nên, **bạn không cần phải mất công xóa logo hay text rác ở đầu file nữa**, app sẽ tự lướt qua chúng và lấy đúng dữ liệu cần thiết!

1. **Double-click file `run_app.bat`** để khởi chạy (hoặc gõ `python main.py` ở terminal).
2. Bấm nút **📂 Open File** (màu xanh dương) và chọn file Excel của bạn.
3. Bấm nút **▶ Process** (màu xanh lá) và bùm — hàng ngàn dòng BOM đã được bóc tách.

---

## 🎨 Giao diện & Đọc kết quả

Bảng kết quả được thiết kế to, rõ ràng và có màu sắc trực quan (lấy cảm hứng từ giao diện Vercel / Linear):

* Các dòng được tô màu dựa theo độ hao hụt để dễ nhận biết:
  * 🔴 **Đỏ (10%)** — Mấy linh kiện dán tí hon dễ rơi rớt (như 0201, 0402).
  * 🟠 **Cam (5%)** — Linh kiện nhỡ (0805, 1206) hoặc dây điện (Wire).
  * 🔵 **Xanh dương (2%)** — Diode, Transistor, IC nhỏ.
  * 🟢 **Xanh lá (1%)** — IC lớn, vi điều khiển, connector.
  * ⚫ **Xám (0%)** — BGA đắt tiền, đồ cơ khí, tấm kim loại.
  * 🩷 **Hồng (?)** — Ca khó! App không hiểu description này. Bạn có thể mở `component_dictionary.json` để dạy thêm cho app hiểu từ vựng này.

Bên trên bảng có thanh **Filter**, gõ vài chữ vào là tìm ngay được linh kiện, hoặc lọc nhanh xem "chỉ hiển thị những món rớt 10%".

---

## ⚙️ Sửa bảng hao hụt (Attrition Rules)

Không thích luật mặc định? Bạn bấm nút **⚙ Edit Rules** ở góc trên cùng bên phải.
Một cửa sổ sẽ bung ra gồm 2 bảng: **SMT/PCBA** và **Cable/Box**.

* Cứ **Double-click** vào một ô `%`, gõ số bạn thích (ví dụ `8` cho 8%). Bấm Enter. 
* Nút **💾 Save Changes** sẽ lưu vĩnh viễn lựa chọn của bạn vào file `attrition_rules.json` và app sẽ tự động xài luật này cho các lần chạy sau. Nếu nhấn nhầm, bấm `↺ Reset` là xong.

---

## 💻 Yêu cầu hệ thống

Công cụ này nhẹ hều và chạy hoàn toàn offline trên máy tính (Windows). 
Bạn chỉ cần cài Python 3.10+, mở Terminal chạy đúng 1 lệnh này là đủ đồ chơi:
```bash
pip install customtkinter openpyxl
```

---
*Build with care, cho những người làm phần cứng.*
