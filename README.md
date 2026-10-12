# Animal Breeding OCS Workflow

Breeding selection workflow for pigs: **BLUP results → selection index (SI) → optimum contribution selection & mating (OCS, optiSel)**.

| Step | File | Language | Function |
|------|------|----------|----------|
| 1 | `si_cal.py` | Python | Reads EBV/ACC (Excel) and the genetic variance matrix from `blupf90.log`, computes SI, outputs a ranked Excel file + a PDF of EBV trends by birth year |
| 2 | `optiSel_pig.R` | R | Shortlists top-SI sires and dams, optimises contributions under a kinship ceiling (based on Ne), and pairs animals with the lowest parental kinship |
| – | `ocs_webtool.py` | Python | Local web interface that calls the two scripts above (no logic is re-implemented) |

## Installation

```bash
# Python
pip install -r requirements.txt

# R (one-time)
Rscript -e 'install.packages(c("optiSel","dplyr","readxl","clue"))'
```

## Usage

### Web tool (recommended)

```bash
python ocs_webtool.py
```
The browser opens `http://127.0.0.1:8765`. Keep all 3 files in the same folder, or use `--scripts-dir`, `--si-script`, `--ocs-script`.

### Command line

Step 1 – selection index:
```bash
python si_cal.py --base-dir ../data_260929/blup_LL_T2613 \
    --dat-file datLL_T2613.dat --genetic-effect 7 \
    --ref-min-year 2025 --sd-mode genetic --weights D105=2,Live=1.5
```

Step 2 – OCS:
```bash
Rscript optiSel_pig.R --input-dir ../data_260929/blup_LL_T2613 \
    --ped-file PedLL.ped --si-file EBV_ACC_SI.xlsx \
    --sires 40 --dams 400 --birth-year 2026
```

For all options: `python si_cal.py --help`, `Rscript optiSel_pig.R --help`.

## Input data

- `EBV_ACC.xlsx`: columns `ID`, `EBV_<trait>`, `ACC_<trait>`
- `blupf90.log`: contains "Genetic variance(s) for effect N"
- `.dat` file: no header, ID in column 1, first 2 characters of column 4 = birth year (YY)
- Pedigree file: no header — `ID Sire Dam BirthDate(YYYYMMDD) Sex(F/M)`

> Real data (pedigree, EBVs, etc.) is **not** committed to the repo (see `.gitignore`).

## License

MIT
