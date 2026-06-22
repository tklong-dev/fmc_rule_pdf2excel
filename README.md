# Firepower Policy PDF → Excel

Chuyển đổi file PDF báo cáo Access Control Policy từ Cisco Firepower Management Center (FMC) thành file Excel. Mỗi file PDF sinh ra một file `.xlsx` gồm 3 sheet: **Rules**, **Object Groups**, và **Network**.

## Yêu cầu

**Python 3.8+** — thư viện phụ thuộc được tự động cài khi chạy lần đầu:

```
pdfminer.six    openpyxl
```

## Cách dùng

1. Đặt tất cả file PDF cần chuyển đổi vào thư mục `input/`.
2. Chạy script:

```bash
python fmc_export_pdf2excel.py
```

3. File kết quả được tạo trong thư mục `output/`, mỗi PDF cho ra một file Excel riêng.

```
input/
  FW-INT-POLICY.pdf
  VUNI-INT-POLICY.pdf

output/
  FW-INT-POLICY.xlsx
  VUNI-INT-POLICY.xlsx
```

> Nếu thư mục `input/` chưa tồn tại, script sẽ tự tạo và dừng lại để bạn đặt file PDF vào.

## Đầu vào / Đầu ra

**Đầu vào:** Một hoặc nhiều file `.pdf` trong thư mục `input/` — báo cáo Access Control Policy export từ FMC.

**Đầu ra:** Mỗi PDF sinh một file `.xlsx` gồm **3 sheet**:

| Sheet | Nội dung |
|-------|----------|
| `<tên PDF>` | Toàn bộ Access Control Rules |
| `Object Groups` | Các nhóm object (group name + danh sách members) |
| `Network` | Các network object (tên + giá trị IP/CIDR/range) |

## Cấu trúc sheet Rules

Mỗi rule chiếm một dòng với 33 cột:

| Nhóm | Cột |
|------|-----|
| Thông tin rule | Policy, Rule Number, Rule Name, Enabled, Section |
| Traffic | Action, Source Zones, Destination Zones, Source Tunnels |
| Địa chỉ mạng | Source Networks, Original Client Networks, Destination Networks |
| Ứng dụng & User | Safe Search, Youtube EDU, VLAN Tags, Users, Application Filters |
| Cổng dịch vụ | Source Ports, Destination Ports |
| Identity | Source ISE Metadata, Destination ISE Metadata, Security Group Tag |
| Khác | Time Range, URLs, Intrusion Policy, Variable Set, File Policy |
| Logging | Log at Beginning of Connection, Log at End of Connection, Log File Events, Send Events to Defense Center, Send using specific syslog alert, Comments |

**Màu sắc theo Action:**

| Màu | Action |
|-----|--------|
| Xanh lá nhạt | Allow |
| Đỏ nhạt | Block |
| Vàng nhạt | Monitor |
| Xanh dương nhạt | Trust |

## Cấu trúc sheet Object Groups

| Cột | Mô tả |
|-----|-------|
| Policy | Tên policy (tên file PDF) |
| Group Name | Tên object group |
| Members | Danh sách member object, mỗi tên trên một dòng |

## Cấu trúc sheet Network

| Cột | Mô tả |
|-----|-------|
| Policy | Tên policy (tên file PDF) |
| Network Name | Tên network object |
| Value | Giá trị: IP host, CIDR subnet, hoặc IP range |

## Tính năng định dạng (tất cả sheet)

- Dòng tiêu đề cố định (freeze pane) — cuộn xuống vẫn thấy header.
- Auto-filter trên tất cả cột.
- Các ô chứa nhiều giá trị được xuống dòng trong cùng một ô.

## Ví dụ kết quả

```
input/FW-INT-POLICY.pdf   → output/FW-INT-POLICY.xlsx   (14 rules, 1 groups, 38 networks)
input/VUNI-INT-POLICY.pdf → output/VUNI-INT-POLICY.xlsx (279 rules, 19 groups, 547 networks)
input/VUNI-SF-POLICY.pdf  → output/VUNI-SF-POLICY.xlsx  (215 rules, 17 groups, 492 networks)
```

## Cơ chế hoạt động

PDF từ FMC có bố cục **2 cột**: cột trái là tên trường (label), cột phải là giá trị. Script dùng `pdfminer.six` để đọc tọa độ x, y của từng text box và ghép label–value theo vị trí dọc (y-coordinate), thay vì dùng text extraction tuyến tính (dễ bị lẫn thứ tự khi PDF có multi-column layout).

Phần **Referenced Objects** ở cuối PDF được nhận diện theo chiều cao của text box: header chính (~13.8 pt), subsection header (~11.5 pt), và data item (~8 pt) — từ đó tách riêng Object Groups và Network objects.

## Giới hạn

- Chỉ hỗ trợ định dạng PDF báo cáo của Firepower/FMC. PDF từ nguồn khác sẽ không parse được.
- Nếu PDF được export với layout khác phiên bản FMC hiện tại, cần kiểm tra lại ngưỡng `LABEL_X_MAX` (mặc định `200`) trong `fmc_export_pdf2excel.py`.
