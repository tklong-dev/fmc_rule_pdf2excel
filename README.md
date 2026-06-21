# Firepower Policy PDF → Excel

Chuyển đổi file PDF báo cáo Access Control Policy từ hệ thống Cisco Firepower thành file Excel, mỗi rule thành một dòng, mỗi trường thành một cột.

## Yêu cầu

**Python 3.8+** và hai thư viện:

```bash
pip install pdfminer.six openpyxl
```

## Cách dùng

1. Đặt tất cả file PDF cần chuyển đổi vào thư mục `input/`.
2. Chạy script:

```bash
python fmc_rule_pdf2excel.py
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

**Đầu vào:** Một hoặc nhiều file `.pdf` đặt trong thư mục `input/` — báo cáo Access Control Policy export từ Firepower Management Center (FMC).

**Đầu ra:** Mỗi file PDF sinh ra một file `.xlsx` tương ứng trong thư mục `output/`, mỗi file có một sheet duy nhất chứa toàn bộ rules.

## Cấu trúc Excel

Mỗi sheet có 33 cột:

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

**Tính năng định dạng:**
- Dòng tiêu đề cố định (freeze pane) — cuộn xuống vẫn thấy header.
- Auto-filter trên tất cả cột.
- Các ô chứa nhiều giá trị (ví dụ: Source Networks liệt kê nhiều host/network) được xuống dòng trong cùng một ô.

## Ví dụ kết quả

```
input/FW-INT-POLICY.pdf   → output/FW-INT-POLICY.xlsx   (14 rules)
input/VUNI-INT-POLICY.pdf → output/VUNI-INT-POLICY.xlsx (279 rules)
input/VUNI-SF-POLICY.pdf  → output/VUNI-SF-POLICY.xlsx  (215 rules)
```

## Cơ chế hoạt động

PDF từ Firepower có bố cục **2 cột**: cột trái là tên trường (label), cột phải là giá trị. Script dùng `pdfminer.six` để đọc tọa độ x, y của từng text box và ghép label–value theo vị trí dọc (y-coordinate), thay vì dùng text extraction tuyến tính (dễ bị lẫn thứ tự khi PDF có multi-column layout).

## Giới hạn

- Chỉ hỗ trợ định dạng PDF báo cáo của Firepower/FMC. PDF từ nguồn khác sẽ không parse được.
- Nếu PDF được export với layout khác phiên bản FMC hiện tại, cần kiểm tra lại ngưỡng `LABEL_X_MAX` (mặc định `200`) trong file `pdf_to_excel.py`.
