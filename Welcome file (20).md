### PHẦN 1: Phân tích 4 tham số của hàm VLOOKUP

Để tra cứu dữ liệu từ Sheet `Danh_Muc` sang Sheet `Don_Hang`, 4 tham số trong hàm VLOOKUP được xác định như sau:

1.  **`lookup_value` (Giá trị cần tìm):**
    
    -   **Là gì:** Mã linh kiện nằm tại Sheet `Don_Hang` (ví dụ: ô `A2`).
        
    -   **Xử lý bẫy 2 (Khoảng trắng thừa):** Dùng hàm `TRIM(A2)` để tự động loại bỏ các khoảng trắng thừa ở đầu hoặc cuối mã linh kiện, giúp việc so sánh chính xác 100%.
        
2.  **`table_array` (Bảng chứa dữ liệu tra cứu):**
    
    -   **Nằm ở đâu:** Nằm ở Sheet `Danh_Muc` (ví dụ: vùng từ ô `A2` đến `C100` -> `Danh_Muc!$A$2:$C$100` hoặc chọn toàn bộ cột `Danh_Muc!$A:$C`).
        
    -   **Xử lý bẫy 1 (Chặn lỗi #N/A do trượt vùng tra cứu):** Bắt buộc phải dùng **dấu cố định `$`** (nhấn phím `F4`) cho vùng tra cứu (ví dụ:  `Danh_Muc!$A$2:$C$100`). Nếu không cố định, khi kéo công thức xuống, bảng tra sẽ bị dịch chuyển dần xuống dưới dẫn đến không tìm thấy dữ liệu.
        
3.  **`col_index_num` (Thứ tự cột lấy dữ liệu):**
    
    -   Tính từ cột chứa Mã linh kiện (cột 1) trong bảng tra:
        
        -   Lấy **Tên linh kiện**: Cột thứ **2**.
            
        -   Lấy **Đơn giá**: Cột thứ **3**.
            
4.  **`range_lookup` (Phương thức tìm kiếm):**
    
    -   Điền **`0`** (hoặc `FALSE`) để tìm kiếm chính xác tuyệt đối theo Mã linh kiện.
        

### PHẦN 2: Triển khai công thức hoàn chỉnh

Giả sử dữ liệu được sắp xếp như sau:

-   **Sheet `Don_Hang`**: Ô `A2` chứa Mã linh kiện.
    
-   **Sheet `Danh_Muc`**: Cột `A` (Mã linh kiện), Cột `B` (Tên linh kiện), Cột `C` (Đơn giá).
    

#### 1. Công thức lấy Tên linh kiện (Điền vào ô B2 của Sheet `Don_Hang`)

Excel

```
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A:$C, 2, FALSE), "Không tìm thấy")

```

#### 2. Công thức lấy Đơn giá (Điền vào ô C2 của Sheet `Don_Hang`)

Excel

```
=IFERROR(VLOOKUP(TRIM(A2), Danh_Muc!$A:$C, 3, FALSE), 0)

```

### 💡 Giải thích chi tiết cách xử lý lỗi & bẫy dữ liệu

-   **Bẫy 1 được giải quyết bằng `Danh_Muc!$A:$C` (hoặc `$A$2:$C$100`):** Giữ cố định vị trí bảng dữ liệu tra cứu khi copy/kéo công thức xuống 100 dòng.
    
-   **Bẫy 2 được giải quyết bằng `TRIM(A2)`:** Chuẩn hóa mã linh kiện, loại bỏ khoảng trắng vô tình bị thừa ở cuối chuỗi.
    
-   **Xử lý lỗi thân thiện bằng `IFERROR(...)`:** Nếu mã linh kiện không tồn tại trong danh mục, thay vì hiển thị lỗi `#N/A` gây mất thẩm mỹ, công thức sẽ trả về `"Không tìm thấy"` (đối với tên) hoặc `0` (đối với đơn giá).
