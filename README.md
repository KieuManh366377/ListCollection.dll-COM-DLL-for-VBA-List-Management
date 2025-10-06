# 📦 ListCollection.dll – COM DLL for VBA List Management

## 🧠 Giới thiệu

`ListCollection.dll` là một thư viện COM được phát triển bằng **C++ Builder của Delphi**, cung cấp đối tượng `List` cho môi trường **VBA (Visual Basic for Applications)**. Thư viện này giúp người dùng VBA thao tác với danh sách dữ liệu một cách linh hoạt và hiệu quả, hỗ trợ đầy đủ các chức năng như thêm, xóa, chèn, duyệt, sắp xếp và chuyển đổi sang mảng.

---

## 🔧 Công nghệ

- **Ngôn ngữ**: C++ (Delphi C++ Builder)
- **Mô hình COM**: Apartment Threading Model
- **Interface**: Dual Interface (`IDispatch` và `IUnknown`)
- **Tên DLL**: `ListCollection.dll`
- **Tương thích**: Excel, Word, Access, và các ứng dụng hỗ trợ VBA

---

## 📚 Các hàm chính và chức năng

1️⃣ **Add(Value)**  
Thêm một phần tử mới vào cuối danh sách.  
Ví dụ: `lst.Add "Apple"`

2️⃣ **Remove(Index)**  
Xóa phần tử tại vị trí chỉ định (chỉ số bắt đầu từ 1).  
Ví dụ: `lst.Remove 2`

3️⃣ **Insert(Index, Value)**  
Chèn phần tử vào vị trí cụ thể trong danh sách.  
Ví dụ: `lst.Insert 2, "Banana"`

4️⃣ **Clear()**  
Xóa toàn bộ danh sách, đưa về trạng thái rỗng.  
Ví dụ: `lst.Clear`

5️⃣ **Item(Index)** hoặc `lst(Index)`  
Truy xuất phần tử tại vị trí chỉ định.  
Ví dụ: `lst.Item(1)` hoặc `lst(1)`

6️⃣ **Count()**  
Trả về số lượng phần tử hiện có trong danh sách.  
Ví dụ: `Debug.Print lst.Count`

7️⃣ **Contains(Value)**  
Kiểm tra xem phần tử có tồn tại trong danh sách hay không.  
Trả về `True` hoặc `False`.  
Ví dụ: `found = lst.Contains("Apple")`

8️⃣ **Replace(Index, NewValue)**  
Thay thế phần tử tại vị trí chỉ định bằng giá trị mới.  
Ví dụ: `lst.Replace 2, "Tiger"`

9️⃣ **IndexOf(Value)**  
Trả về vị trí của phần tử đầu tiên tìm thấy.  
Ví dụ: `pos = lst.IndexOf("Banana")`

🔟 **IndexOfIgnoreCase(Value)**  
Tìm vị trí phần tử không phân biệt chữ hoa/thường.  
Ví dụ: `lst.IndexOfIgnoreCase("banana")`

1️⃣1️⃣ **IndexOfEx(Value, IgnoreCase)**  
Tìm vị trí phần tử với tùy chọn phân biệt hoặc không phân biệt chữ hoa/thường.  
Ví dụ: `lst.IndexOfEx("banana", True)`

1️⃣2️⃣ **Sort(Ascending)**  
Sắp xếp danh sách theo thứ tự tăng (`True`) hoặc giảm (`False`).  
Ví dụ: `lst.Sort True`

1️⃣3️⃣ **ToArray()**  
Xuất danh sách thành mảng Variant.  
Ví dụ: `arr = lst.ToArray()`

1️⃣4️⃣ **ToVariantArray()**  
Xuất danh sách dưới dạng mảng Variant chuẩn.  
Ví dụ: `arr = lst.ToVariantArray()`

1️⃣5️⃣ **_NewEnum()**  
Cho phép duyệt danh sách bằng vòng lặp `For Each` trong VBA.  
Ví dụ:
```vb
For Each item In lst
    Debug.Print item
Next
```

---

## 🧪 Ví dụ sử dụng trong VBA

### 🔄 Thay thế phần tử bằng `Replace`

```vb
Sub DemoReplace()
    Dim v
    Dim lst As New ListCollection.List
    lst.Add "Dog"
    lst.Add "Cat"
    lst.Add "Bird"

    lst.Replace 2, "Tiger"
    For Each v In lst
        Debug.Print v
    Next
End Sub
```

**Kết quả**:
```
Dog
Tiger
Bird
```

---

### 📋 Thao tác đầy đủ với danh sách

```vb
Sub DemoListCollection()
    Dim lst As New ListCollection.List
    Dim arr As Variant
    Dim v As Variant
    Dim i As Long
    Dim found As Boolean

    lst.Add "Apple"
    lst.Add "Banana"
    lst.Add "Cherry"
    Debug.Print "Count ="; lst.Count

    Debug.Print "Item(1) ="; lst.Item(1)
    Debug.Print "Item(2) ="; lst(2)

    For Each v In lst
        Debug.Print "Value:"; v
    Next

    lst.Remove 2
    lst.Insert 2, "NewBanana"
    Debug.Print "Item(2) ="; lst(2)

    found = lst.Contains("Cherry")
    Debug.Print "Contains 'Cherry'?"; found

    arr = lst.ToVariantArray()
    For i = LBound(arr) To UBound(arr)
        Debug.Print "Arr(" & i & ")=" & arr(i)
    Next

    lst.Clear
    lst.Add "X"
    lst.Add 123
    lst.Add #10/3/2025#
    lst.Add 45.67
    For Each v In lst
        Debug.Print "Type:"; TypeName(v); " Value:"; v
    Next

    Dim lst2 As New ListCollection.List
    lst2.Add "A"
    lst2.Add "B"
    For Each v In lst
        Dim v2 As Variant
        For Each v2 In lst2
            Debug.Print v, v2
        Next
    Next

    lst.Clear
    Debug.Print "Count after Clear ="; lst.Count
End Sub
```

---

### 🔁 Duyệt danh sách bằng `For Each`

```vb
Sub TestListEnum()
    Dim lst As New ListCollection.List
    Dim item

    lst.Add "A"
    lst.Add "B"
    lst.Add "C"

    For Each item In lst
        Debug.Print item
    Next
End Sub
```

---

### 🔢 Sắp xếp và tìm vị trí phần tử

```vb
Sub DemoSortAndIndex()
    Dim lst As New ListCollection.List
    lst.Add "Orange"
    lst.Add "Apple"
    lst.Add "Banana"

    lst.Sort True
    For Each v In lst
        Debug.Print v
    Next

    Debug.Print "IndexOf 'Banana' ="; lst.IndexOf("Banana")
    Debug.Print "IndexOfIgnoreCase 'apple' ="; lst.IndexOfIgnoreCase("apple")
End Sub
```

---

## 🎯 Ứng dụng thực tế

- Quản lý danh sách dữ liệu trong Excel VBA
- Xử lý chuỗi, số, ngày tháng hoặc đối tượng COM
- Tạo danh sách động để xử lý logic phức tạp trong Access hoặc Word
- Duyệt danh sách bằng `For Each` như mảng thông thường
- Tích hợp vào macro xử lý dữ liệu, báo cáo, hoặc tự động hóa

---

## 👤 Tác giả

**Kiều Mạnh**  
📧 Email: [kieumanh366377@gmail.com](mailto:kieumanh366377@gmail.com)

---
