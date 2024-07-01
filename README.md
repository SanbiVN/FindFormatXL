# FindFormatXL
 Tìm kiếm và định dạng chuỗi trong Ô và đối tượng Excel



	# HÀM TÌM KIẾM ĐỊNH DẠNG CHUỖI TRONG Ô, CHÚ THÍCH VÀ HÌNH DẠNG EXCEL				
	với Hàm FindFormat				
			- hàm tìm kiếm định dạng chuỗi trong ô, chú thích và hình dạng excel		
	## Hướng dẫn sử dụng hàm FindFormat:				
		- Hàm: 	=FindFormat(Finds, FindObject, Arguments,...)		
		-	Cách viết hàm nhanh, gõ vào ô chuỗi =FindFormat và ấn tổ hợp phím Ctrl+Shift+A		
		## Tham số :			
| Vị trí | Tham số    | Kiểu                                                    | Diễn giải                                                                                                                                |
|--------|------------|---------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------|
| 1      | Finds      | Chuỗi/mảng chuỗi/vùng ô                                 | Nhập chuỗi   "tìm a", mảng {"tìm a";"tìm b"}, vùng ô A1:A100                                                                             |
| 2      | FindObject | Chuỗi/mảng   chuỗi/vùng ô      hoặc hàm đối số bên dưới | Nếu nhập là ô thì   tìm trong ô, nếu nhập là chuỗi và mảng thì tìm trong các đối tượng hình dạng   (tự động hiểu hình dạng là chú thích) |
| 3      | Arguments  | Các hàm đối số bổ trợ                                   | Những màu tô cho   chuỗi đã tìm thấy                                                                                                     |
					
		-	Gõ hàm FindFormat_HuongDan() để hiển thị hướng dẫn này trong trang tính của bạn		
					
		- Các hàm dưới đây nhầm đặt giá trị để thực hiện tìm kiếm, và chúng phải được gõ trong hàm FindFormat			
					
| Các hàm nhập vào đối số FindObject | Kiểu                                       |   |
|------------------------------------|--------------------------------------------|---|
| fff_FindRange(Cells)               | Tìm nhiều vùng ô                           |   |
| fff_FindShape(Shapes)              | Tìm nhiều hình dạng                        |   |
| fff_WorksheetRange                 | Tìm trong tất cả ô trang tính gõ hàm       |   |
| fff_WorkbookRange                  | Tìm trong ô trên toàn Sổ làm việc          |   |
| fff_WorksheetShapes                | Tìm trong hình dạng trên trang tính gõ hàm |   |
| fff_WorkbookShapes                 | Tìm trong ô trên toàn Sổ làm việc          |   |
| fff_WorksheetComments              | Tìm trong ghi chú trên trang tính gõ hàm   |   |
| fff_WorkbookComments               | Tìm trong ghi chú trên toàn Sổ làm việc    |   |
					
| Các hàm nhập sau đối số FindObject | Kiểu                                                                                            |   |
|------------------------------------|-------------------------------------------------------------------------------------------------|---|
| fff_FindIndex(Index)               | thứ tự định dạng, ví dụ đặt là 3, nếu tổng lần tìm   được 5, chỉ có vị trí thứ 3 được định dạng |   |
| fff_CompareMode()                  | Đặt tìm kiếm có phân biệt ký tự hoa và thường                                                   |   |
| fff_Regex(groupIndex = -1)         | Sử dụng biểu thức chính quy để tìm kiếm, groupIndex   tương tự FindIndex                        |   |
| fff_Colors(Colors())               | Đổi màu phông chữ,   Xem thêm hướng dẫn nhập màu sắc bên dưới                                   |   |
| fff_DefaultColor(Color)            | Đặt màu phông mặt định, nếu không tìm thấy đặt lại màu   toàn bộ chuỗi                          |   |
| fff_Name(FontName)                 | Đổi tên phông                                                                                   |   |
| fff_Bold()                         | Đổi phông đậm                                                                                   |   |
| fff_Size(FontSize)                 | Đổi kích thước phông                                                                            |   |
| fff_Italic()                       | Đổi phông in nghiên                                                                             |   |
| fff_StrikeThrough()                | Đổi gạch giữa                                                                                   |   |
| fff_Underline(value = 2)           | Đổi gạch dưới                                                                                   |   |
| fff_Superscript()                  | Đổi chỉ số trên                                                                                 |   |
| fff_Subscript()                    | Đổi chỉ số dưới                                                                                 |   |
| fff_Target(target)                 | Đặt ô sẽ sao chép đến và tìm định dạng                                                          |   |
					
					Nhập màu sắc
					Ví dụ với số: fff_Colors(255, 65536, 16777215)
					Ví dụ với mã: fff_Colors("#FFFFFF", "FFFFF", "FF")
					Ví dụ với tên: fff_Colors("yellow", "blue")
					
| yellow, ye, yl |
|----------------|
| red, re        |
| blue           |
| green, gr      |
| cyan, cy       |
| magenta, ma    |
| white, wh, wi  |
| black, bl, bk  |
| orange, or     |
| pink           |
| purple, pu     |
| silver, si     |
| violet, vi     |
| Brown, br      |
| Beige, be      |


		Lưu ý: Để sử dụng được Hàm FindFormat trong dự án mới, trong VBA hãy sao chép module modFindFormatFont	



		Liên hệ Messenger: https://m.me/he.sanbi		
