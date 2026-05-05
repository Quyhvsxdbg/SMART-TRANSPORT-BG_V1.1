# Rule Auto Tinh NguoiKy

Ap dung cho cac trang quyet dinh, thong bao, cong van trong app `SMART-TRANSPORT-BG_V1.1`.

## 1. Rule lay ten tinh tu dong

Nguon du lieu:
- Bang `ThongTin`
- Cot dung: `Tinh`, `Ngay_hieu_luc`

Logic chuan:
1. Tai du lieu `ThongTin`.
2. Xac dinh moc ngay de chon tinh:
   - Neu ban ghi dang in/xuat co `NgayKy` thi dung `NgayKy`.
   - Neu khong co `NgayKy` thi dung `new Date()`.
3. Loc cac dong co `Ngay_hieu_luc <= moc ngay`.
4. Sap xep giam dan theo `Ngay_hieu_luc`.
5. Lay dong dau tien.
6. Dung truong `Tinh` de do vao:
   - header
   - noi dung van ban
   - footer ngay thang
   - du lieu export Word

Pseudo-code:

```js
function getTinhByNgayKy(thongTinList, ngayKy) {
    const mocTinhHieuLuc = parseDateValue(ngayKy) || new Date();

    const dsHopLe = (thongTinList || []).filter(item => {
        const ngayHieuLuc = parseDateValue(item.Ngay_hieu_luc);
        return !ngayHieuLuc || ngayHieuLuc <= mocTinhHieuLuc;
    });

    const sorted = dsHopLe.sort((a, b) =>
        new Date(b.Ngay_hieu_luc) - new Date(a.Ngay_hieu_luc)
    );

    return sorted[0]?.Tinh || '';
}
```

## 2. Rule lay nguoi ky tu dong

Nguon du lieu:
- Bang quyet dinh / thong bao: cot `NguoiKy`
- Bang `NguoiPhuTrach`: cot `IDNguoi`, `HoTen`, `ChucVu`, `ThamQuyen`

Logic chuan:
1. Lay gia tri `NguoiKy` tu ban ghi dang xu ly.
2. Query bang `NguoiPhuTrach` theo `IDNguoi = NguoiKy`.
3. Lay ban ghi tim duoc.
4. Dung:
   - `HoTen` -> ten nguoi ky
   - `ChucVu` -> chuc vu ky
   - `ThamQuyen` -> tham quyen ky

Quy tac bat buoc:
- Khong dung `TenNguoiKy` nua.
- Ten nguoi ky chi lay tu `NguoiPhuTrach.HoTen`.
- Khong xet `NgayKy`, `NgayBatDau`, `NgayKetThuc` de chon nguoi ky.
- Nguoi ky phu thuoc duy nhat vao gia tri `NguoiKy` duoc chon trong AppSheet.

Pseudo-code:

```js
function getNguoiPhuTrachInfo(record, nguoiPhuTrachList) {
    if (!record || !Array.isArray(nguoiPhuTrachList) || nguoiPhuTrachList.length === 0) {
        return null;
    }

    const nguoiKyId = String(record.NguoiKy || '').trim();
    if (!nguoiKyId) {
        return null;
    }

    return nguoiPhuTrachList.find(item =>
        String(item.IDNguoi || '').trim() === nguoiKyId
    ) || null;
}

function getNguoiKyName(nguoiPhuTrachInfo) {
    return String(nguoiPhuTrachInfo?.HoTen || '').trim();
}
```

## 2A. Rule chon nguoi ky tren AppSheet / web

Muc tieu:
- Cho nguoi dung chon `NguoiKy`
- Nhung chi duoc chon dung nhom nguoi co lien quan

Rule loc danh sach chon:
- Bang `NguoiPhuTrach`
- Chi lay nguoi thoa:
  - `[VaiTro] = "Lanh dao So"`
  - `ISBLANK([NgayKetThuc]) OR [NgayKetThuc] >= TODAY()`

AppSheet:
- Neu dung `Valid If` cho cot `NguoiKy` thi phai tra ve list `IDNguoi`

Cong thuc mau:

```appsheet
SELECT(
  NguoiPhuTrach[IDNguoi],
  AND(
    [VaiTro] = "Lãnh đạo Sở",
    OR(
      ISBLANK([NgayKetThuc]),
      [NgayKetThuc] >= TODAY()
    )
  )
)
```

Neu can gan mac dinh 1 nguoi ky thi moi dung `ANY(...)`.

Tren web:
1. Tai danh sach `NguoiPhuTrach`.
2. Loc theo dung rule tren.
3. Bind vao dropdown `NguoiKy`.
4. Khi nguoi dung doi dropdown:
   - Lay `IDNguoi` dang chon
   - Tim lai ban ghi `NguoiPhuTrach`
   - Do `HoTen`, `ChucVu`, `ThamQuyen` ra giao dien
   - Truyen cung bo gia tri do vao Word export

Pseudo-code:

```js
function filterNguoiPhuTrachForSigner(list) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return (list || []).filter(item => {
        const vaiTro = String(item.VaiTro || '').trim().toLowerCase();
        const ngayKetThuc = parseDateValue(item.NgayKetThuc);

        if (ngayKetThuc) {
            ngayKetThuc.setHours(0, 0, 0, 0);
        }

        return vaiTro === 'lanh dao so'.toLowerCase() &&
            (!ngayKetThuc || ngayKetThuc >= today);
    });
}
```

## 3. Rule ap dung tren giao dien va file Word

Bat buoc dong bo 2 noi:
- HTML hien thi
- Du lieu export Word

Map chuan:
- `nguoi-ky` -> `NguoiPhuTrach.HoTen`
- `tham-quyen-ky` -> `NguoiPhuTrach.ThamQuyen`
- `chuc-vu-ky` -> `NguoiPhuTrach.ChucVu`
- `tinh` / `tinh_upper` / `tinh-footer` -> ket qua tu `ThongTin` theo moc `NgayKy`

## 4. Checklist khi lam trang moi

1. Ban ghi chinh phai co cot `NguoiKy`.
2. Khong doc `TenNguoiKy`.
3. Neu la trang can cho nguoi dung chon nguoi ky, phai loc dropdown `NguoiKy` theo `VaiTro = "Lanh dao So"` va `NgayKetThuc` con hieu luc.
4. Gia tri luu cua `NguoiKy` phai la `IDNguoi`, khong luu ten.
5. Tinh phai lay tu bang `ThongTin`, khong hard-code.
6. Tinh phai tinh theo `NgayKy`, neu thieu `NgayKy` thi dung ngay hien tai.
7. HTML va Word phai dung chung mot logic.
8. Neu co nhieu cho do cung mot gia tri, tao helper dung chung.

## 5. Trang da ap dung rule nay

- `quyetdinhthuhoi_phuhieu_core.html`
- `quyetdinhthuhoi_gbkd_core.html`
