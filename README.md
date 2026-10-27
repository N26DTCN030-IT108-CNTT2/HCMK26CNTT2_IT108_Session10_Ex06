# HCMK26CNTT2_IT108_Session10_Ex06
bài tập số 6
# 🚀 Bộ Hồ Sơ Triển Khai Nền Tảng Đấu Giá Bất Động Sản Smart-Home (Session 10 Ex06)

Bài giải hoàn thiện bộ tài liệu đặc tả SRS, quản lý dữ liệu bất động sản và Slide Pitch Deck gọi vốn cho dự án khởi nghiệp Smart-Home.

---

# Phần 1. Phân Tích So Sánh Bảng Excel "Xấu" vs Bảng Excel "Đúng"

## 1. So sánh chi tiết theo 5 Quy tắc Vàng Database Mindset

| Quy tắc / Tiêu chí | Bảng Excel "Xấu" (Vi phạm quy tắc) | Bảng Excel "Đúng" (Chuẩn Database) | Hậu quả kỹ thuật của Bảng "Xấu" |
| :--- | :--- | :--- | :--- |
| **1. Hàng tiêu đề (Header Row)** | Dùng nhiều hàng gộp lại (Merge Cells) để tạo tiêu đề trang trí lớn. | Chỉ có duy nhất **1 hàng Header ở Dòng 1**, các ô tên cột độc lập. | `AutoFilter` nhận diện sai tên trường, gây lỗi khi lọc dữ liệu. |
| **2. Tính nguyên tử (Atomic Data)** | Gộp chung Địa chỉ và Tên chủ nhà vào cùng 1 ô (ô chứa nhiều thông tin). | Tách riêng `Dia_Chi`, `Ten_Chu_Nha`, `So_Dien_Thoai` thành từng cột độc lập. | Không thể sắp xếp (Sort) hay lọc (Filter) theo từng tiêu chí nhỏ. |
| **3. Dòng/Cột trống (Empty Rows)** | Tùy tiện chèn các dòng trống nằm giữa dữ liệu để phân cách cho dễ nhìn. | Dữ liệu liên tục từ dòng 2 đến dòng cuối cùng, không chứa dòng trống ngắt đoạn. | `AutoFilter` và vùng quét công thức bị đứt gãy, bỏ sót dữ liệu phía dưới. |
| **4. Định dạng dữ liệu (Data Types)** | Nhập lẫn chữ và số trong cùng một cột (ví dụ: `"15 triệu"` thay vì `15000000`). | Mỗi cột giữ đúng 1 kiểu dữ liệu thống nhất (Number, Text, Date). | Không thực hiện được các hàm toán học (`SUM`, `AVERAGE`, `VLOOKUP`). |
| **5. Dòng tổng cộng (Summary Rows)** | Chèn dòng "TỔNG CỘNG" ngắt ngang giữa các bản ghi bất động sản. | Dòng tổng cộng nằm **hoàn toàn bên ngoài/dưới** vùng dữ liệu thô. | Dòng tổng bị coi là một Record bất động sản, làm sai lệch kết quả lọc. |

---

## 2. Phân tích khái niệm Record và Field trong quản lý BĐS

- **Field (Trường / Cột dữ liệu):** Là thuộc tính duy nhất dùng để mô tả một khía cạnh thông tin của tất cả các bất động sản.
  - *Ví dụ:* Cột `Gia_Thue_Thang_VND` (Cột C) chứa toàn bộ giá trị tiền thuê hàng tháng của các căn hộ.
- **Record (Bản ghi / Dòng dữ liệu):** Là tập hợp đầy đủ tất cả các thuộc tính mô tả trọn vẹn cho **đúng 1 đối tượng bất động sản cụ thể**.
  - *Ví dụ Record Dòng 2:* `BDS001 | 123 Nguyễn Văn Cừ, Q5 | 15,000,000 | MG01 | 5% | 750,000`

---

# Phần 2. Triển Khai Học Liệu Dự Án Smart-Home

## 1. File Word Đặc Tả SRS (`SRS_Smart_Home.docx`)
* **Phân cấp Heading Styles chuẩn:**
  - `Heading 1`: Các chương lớn (1. GIỚI THIỆU DỰ ÁN SMART-HOME, 2. YÊU CẦU CHỨC NĂNG, 3. QUẢN LÝ HOA HỒNG).
  - `Heading 2`: Các mục con (1.1 Bối cảnh, 1.2 Phạm vi, 2.1 Đăng ký BĐS, 2.2 Đấu giá trực tuyến).
  - `Heading 3`: Chi tiết luồng kỹ thuật (2.2.1 Luồng trả giá và khóa phiên đấu giá).
* **Mục lục tự động (Auto TOC):** Tích hợp trường `Table of Contents` tại Trang 1, tự động thu thập tiêu đề và cập nhật số trang thời gian thực.

---

## 2. File Excel Quản Lý Bất Động Sản (`Quan_Ly_Nha_Dat_SmartHome.xlsx`)
* **Sheet `DanhSachNhaDat`:** Quản lý mã BĐS, địa chỉ, giá thuê, mã loại môi giới và tính tiền hoa hồng.
* **Sheet `BangTraHoaHong`:** Bảng danh mục tỷ lệ chiết khấu hoa hồng theo cấp độ môi giới.
* **Công thức VLOOKUP tra cứu hoa hồng:**
  ```excel
  =VLOOKUP(D2, BangTraHoaHong!A$2:C$5, 2, FALSE)
