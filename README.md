# Pig Breeding OCS Workflow

Quy trình chọn giống heo: **kết quả BLUP → chỉ số chọn lọc (SI) → tối ưu hoá đóng góp & ghép đôi (OCS, optiSel)**.

| Bước | File | Ngôn ngữ | Chức năng |
|------|------|----------|-----------|
| 1 | `si_cal.py` | Python | Đọc EBV/ACC (Excel) và ma trận phương sai di truyền từ `blupf90.log`, tính SI, xuất Excel xếp hạng + PDF xu hướng EBV theo năm sinh |
| 2 | `optiSel_pig.R` | R | Chọn đực/cái SI cao, tối ưu đóng góp dưới ngưỡng kinship (theo Ne), ghép đôi có kinship cha mẹ thấp nhất |
| – | `ocs_webtool.py` | Python | Giao diện web chạy local, gọi 2 script trên (không cài lại logic) |

## Cài đặt

```bash
# Python
pip install -r requirements.txt

# R (một lần)
Rscript -e 'install.packages(c("optiSel","dplyr","readxl","clue"))'
```

## Cách dùng

### Web tool (khuyên dùng)

```bash
python ocs_webtool.py
```
Trình duyệt mở `http://127.0.0.1:8765`. Giữ 3 file cùng thư mục, hoặc dùng `--scripts-dir`, `--si-script`, `--ocs-script`.

### Dòng lệnh

Bước 1 – chỉ số chọn lọc:
```bash
python si_cal.py --base-dir ../data_260929/blup_LL_T2613 \
    --dat-file datLL_T2613.dat --genetic-effect 7 \
    --ref-min-year 2025 --sd-mode genetic --weights D105=2,Live=1.5
```

Bước 2 – OCS:
```bash
Rscript optiSel_pig.R --input-dir ../data_260929/blup_LL_T2613 \
    --ped-file PedLL.ped --si-file EBV_ACC_SI.xlsx \
    --sires 40 --dams 400 --birth-year 2026
```

Xem toàn bộ tuỳ chọn: `python si_cal.py --help`, `Rscript optiSel_pig.R --help`.

## Dữ liệu đầu vào

- `EBV_ACC.xlsx`: cột `ID`, `EBV_<trait>`, `ACC_<trait>`
- `blupf90.log`: chứa "Genetic variance(s) for effect N"
- File `.dat`: không header, ID ở cột 1, 2 ký tự đầu cột 4 = năm sinh (YY)
- File pedigree: không header — `ID Sire Dam BirthDate(YYYYMMDD) Sex(F/M)`

> Dữ liệu thực (pedigree, EBV…) **không** được đưa lên repo (xem `.gitignore`).

## License

MIT
