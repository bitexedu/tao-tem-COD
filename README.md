# tem-cod

Tạo tem COD (Viettel Post) từ bảng kê hóa đơn và các file đơn trình; đồng thời tạo sheet Pivot trong bảng kê.

## Chạy
`powershell -NoProfile -File scripts\make_tem.ps1 [-bangKe "<tên file>"] [-outName "<tên file tem>"] [-root "<thư mục gốc>"]`

Cần Windows + Microsoft Excel (dùng COM).

## Cấu trúc thư mục dữ liệu (thư mục gốc `-root`)
- `Tem COD\BANG KE*.xlsx`: bảng kê (lấy file mới nhất nếu không chỉ định). Script thêm sheet `Pivot` vào file này.
- `GV\DON HANG TRO GIA MT dd.MM.yyyy.xlsx`, `HS\DON HANG TGHS dd.MM.yyyy.xlsx`, `Đơn lẻ\TRÌNH ĐƠN LẺ dd.MM.yyyy.xlsx`: đơn trình theo ngày ghi trong `Work Description`.
- `_auto\base\Tem_base.xlsx`: file khung tem Viettel (không đưa lên repo vì chứa dữ liệu khách hàng).
- Đầu ra: `Tem COD\TEM COD ... - moi.xlsx`.

## Quy tắc
- Số hóa đơn lấy từ cột `NWV VAT Number.`; nhiều hóa đơn của một đơn nối bằng `+`.
- Tên bill: `số HĐ-TG-đầu mối` (giáo viên), `số HĐ-TGHS-đầu mối` (học sinh; `PLUS` nếu mua máy 880 PLUS), `số HĐ-tên` (đơn lẻ).
- COD = cột thành tiền của file đơn trình.
- Địa chỉ HCM: `PHS - Nội tỉnh tiết kiệm thỏa thuận`; còn lại `LCOD - TMĐT tiết kiệm thỏa thuận` + `GDK`.
