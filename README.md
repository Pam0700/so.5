# App xử lý dữ liệu xuất hàng – V7 nâng cấp

## Chức năng đã có

- Đọc file DG gồm: `Code`, `Tên khách`, `Điểm giao`, `Ship`.
- Cho phép chọn **nhiều file sản phẩm** gồm: `Code`, `Tên sản phẩm`.
- Tự tìm hàng tiêu đề, không bắt buộc tiêu đề ở dòng đầu.
- Xử lý theo luồng `Ship → Code → Sản phẩm`.
- Code trùng nhưng tên sản phẩm khác nhau vẫn lấy đầy đủ.
- Chỉ loại bỏ dòng trùng hoàn toàn `Code + Tên sản phẩm` khi bật tùy chọn.
- Gom theo `Tên khách + Điểm giao`; cùng tên nhưng khác điểm giao không gộp.
- Tên khách và điểm giao xuất hiện một lần tại cột J; sản phẩm nằm thành danh sách tại cột K.
- Mặc định xóa chữ VietGAP.
- Sắp xếp: `5kg → 2kg → 1kg → Khay → TF → 500g → nhóm khác`.
- Ghép nhóm khách để mỗi form gần giới hạn 25 sản phẩm nhất.
- Form đầu dùng 24 dòng theo mẫu gốc; form sau dùng 25 dòng.
- Dùng trực tiếp file `BM-QC-26-template.xlsx`; sửa XML bên trong để giữ border, ô gộp, chiều cao dòng, thiết lập in và textbox chữ ký.
- Nếu cần thêm form, app nhân khối form trong file mẫu và thêm ngắt trang.
- Có chế độ gộp nhiều Ship hoặc tách từng Ship thành ZIP.
- Có màn hình kiểm tra lỗi và bảng chuẩn hóa tên khách.

## Chạy ứng dụng

### Cách 1: GitHub Pages

Tải toàn bộ thư mục lên repository GitHub, sau đó bật **Settings → Pages**. Mở đường dẫn Pages được GitHub cung cấp.

### Cách 2: Chạy trên máy tính

Do trình duyệt thường chặn tải file mẫu khi mở trực tiếp bằng `file://`, nên chạy một web server đơn giản trong thư mục:

```bash
python -m http.server 8000
```

Sau đó mở `http://localhost:8000`.

Nếu vẫn mở trực tiếp `index.html`, hãy chọn thủ công file `BM-QC-26-template.xlsx` trong mục **Tùy chọn xuất → Mẫu Excel tùy chọn**.

## Dữ liệu mẫu

- `sample-dg.csv`
- `sample-products-a.csv`
- `sample-products-b.csv`
- `demo-output.xlsx`: bản Excel minh họa đã điền vào đúng mẫu gốc.

## Lưu ý về PDF

Bản static chạy trên GitHub Pages xuất Excel đúng biểu mẫu. Xuất PDF y hệt Excel cần một máy chủ có Microsoft Excel hoặc LibreOffice để render, nên chưa được đưa vào bản này để tránh tạo PDF sai bố cục.
