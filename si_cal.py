#!/usr/bin/env python3
"""
Selection index (SI) from BLUPF90 EBVs + mean-EBV-by-birth-year charts (PDF).
Single self-contained script.

Inputs (all in --base-dir unless an absolute path is given):
    EBV_ACC.xlsx      columns: ID, EBV_<trait>, ACC_<trait> for every trait (1 to 12)
    blupf90.log       contains "Genetic variance(s) for effect N"
    <dat file>        passed with --dat-file; whitespace-separated, no header;
                      animal ID in --dat-id-col (default 1); first 2 characters
                      of --dat-year-col (default 4) = birth year YY

Steps:
    1. Read the genetic covariance matrix of the ANIMAL effect from the log.
    2. Read EBV/ACC and run QC.
    3. Birth year = "20" + first 2 characters of column 4 in the .dat file.
    4. SI_Raw = sum(weight * direction * EBV / genetic_SD)
       SI_100 = 100 + 10 * (SI_Raw - mean_ref) / scale
         mean_ref : mean SI_Raw of the reference group
                    (all animals, or birth year >= --ref-min-year)
         scale    : --sd-mode group    -> SD of SI_Raw in the reference group
                    --sd-mode genetic  -> sqrt(a' G a), genetic SD of the index
    5. Save the ranked Excel file (with a Year column).
    6. Save a PDF with one chart per trait: mean EBV by birth year (mean +/- 1 SE).

IMPORTANT: --genetic-effect must be the effect number of the ANIMAL (additive)
effect (renadd07.ped -> 7). The log may hold several "Genetic variance(s)"
blocks (e.g. 3, 6, 7); a wrong number still parses without any error.

Traits (1 to 12) are set per run, without editing this file:
    --traits     D105,BF105,Loin105,Live,BWt,Wt21
        Names in the SAME ORDER as trait 1..N of the BLUPF90 model (= order of the
        rows/columns of the genetic covariance matrix in the log). Each name must
        match the Excel columns EBV_<name> and ACC_<name>. Allowed characters:
        letters, digits, "_", ".", "-". Default: the six traits listed below.
    --weights    D105=2,Live=1.5       (per 1 genetic SD, >= 0; default 1)
    --directions D105=-1,Live=+1       (-1 = lower is better, +1 = higher is better)
With --traits, a direction must be given for every trait that is not one of the
default traits.

Example:
    python si_cal.py --base-dir "../data_260929/blup_LL_T2613" \
        --dat-file datLL_T2613.dat \
        --genetic-effect 7 --ref-min-year 2025 --sd-mode genetic \
        --weights D105=2,Live=1.5

    # a project with other traits
    python si_cal.py --base-dir ../data_YY --dat-file datYY.dat --genetic-effect 5 \
        --traits ADG,FCR,LMP --directions ADG=+1,FCR=-1,LMP=+1
"""

import argparse
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import matplotlib

matplotlib.use("Agg")  # no display needed
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


# ============================================================
# 1. CONFIGURATION
# ============================================================

MAX_TRAITS = 12

# Defaults, used when --traits is not given. Keep the trait order identical to
# trait 1..N of the BLUPF90 model (and therefore of the covariance matrix).
DEFAULT_TRAITS = ["D105", "BF105", "Loin105", "Live", "BWt", "Wt21"]

# Direction of selection: -1 = lower is better, +1 = higher is better.
DEFAULT_DIRECTION = {
    "D105": -1,
    "BF105": -1,
    "Loin105": 1,
    "Live": 1,
    "BWt": 1,
    "Wt21": 1,
}

# Active configuration. These objects are modified IN PLACE by configure_traits(),
# so every function below always sees the current traits.
TRAITS = list(DEFAULT_TRAITS)
DIRECTION = dict(DEFAULT_DIRECTION)
# Relative weights per 1 genetic SD (equal by default).
WEIGHTS = {trait: 1.0 for trait in TRAITS}

ID_COL = "ID"
EBV_COLS = {t: f"EBV_{t}" for t in TRAITS}
ACC_COLS = {t: f"ACC_{t}" for t in TRAITS}

TRAIT_NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.\-]*$")

YEAR_MIN, YEAR_MAX = 2000, 2035  # plausibility range for birth years


# ============================================================
# 2. COMMAND-LINE ARGUMENTS
# ============================================================

def parse_args():
    p = argparse.ArgumentParser(
        description="Compute SI and plot mean EBV by birth year in one run.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--base-dir", type=Path, required=True,
                   help="Directory containing the input files.")
    p.add_argument("--ebv-file", default="EBV_ACC.xlsx",
                   help="Input Excel with ID, EBV_<trait>, ACC_<trait>.")
    p.add_argument("--log-file", default="blupf90.log",
                   help="BLUPF90 log file.")
    p.add_argument("--dat-file", required=True,
                   help="BLUPF90 data file (whitespace-separated, no header); "
                        "filename relative to --base-dir, or an absolute path.")
    p.add_argument("--genetic-effect", type=int, required=True,
                   help="Effect number of the ANIMAL (additive) effect in the "
                        "log, e.g. 7 for renadd07.ped.")
    p.add_argument("--out-xlsx", default="EBV_ACC_SI.xlsx",
                   help="Output Excel filename or path.")
    p.add_argument("--out-pdf", default="EBV_trend.pdf",
                   help="Output PDF filename or path.")

    # SI options
    p.add_argument("--min-acc", type=float, default=None,
                   help="Optional minimum ACC (0-1) for the SI ranking. Check "
                        "whether ACC is accuracy (r) or reliability (r^2).")
    p.add_argument("--acc-traits", nargs="+", default=None,
                   help="Traits to which --min-acc is applied (default: all).")
    p.add_argument("--ref-min-year", type=int, default=None,
                   help="Reference group for SI_100: birth year >= this "
                        "value (default: all animals).")
    p.add_argument("--sd-mode", choices=["group", "genetic"], default="group",
                   help="Scale for SI_100: 'group' = SD of SI_Raw in the "
                        "reference group; 'genetic' = sqrt(a'Ga).")

    # .dat / chart options
    p.add_argument("--dat-id-col", type=int, default=1,
                   help="1-based column of the animal ID in the .dat file.")
    p.add_argument("--dat-year-col", type=int, default=4,
                   help="1-based column whose first 2 characters are YY.")
    p.add_argument("--century", default="20",
                   help="String prepended to the 2-digit year.")
    p.add_argument("--min-n", type=int, default=10,
                   help="Years with fewer animals are not plotted.")

    # traits, weights and directions
    p.add_argument("--traits", default=None, metavar="T1,T2,...",
                   help="Trait names (1 to 12), in the order of trait 1..N of the "
                        "BLUPF90 model; must match the Excel columns EBV_<name> and "
                        "ACC_<name>. Default: D105,BF105,Loin105,Live,BWt,Wt21.")
    p.add_argument("--weights", default=None, metavar="T=W,...",
                   help="Relative weights per 1 genetic SD (>= 0), e.g. "
                        "D105=2,Live=1.5. Traits not listed keep the default.")
    p.add_argument("--directions", default=None, metavar="T=D,...",
                   help="Selection direction per trait: -1 = lower is better, "
                        "+1 = higher is better, e.g. D105=-1,Loin105=+1.")
    return p.parse_args()


def parse_trait_map(text, option, kind):
    """Parse 'D105=2,Live=1.5' -> {'D105': 2.0, 'Live': 1.5} (kind: 'weight' or 'direction')."""
    result = {}

    for item in re.split(r"[,;\s]+", text.strip()):
        if not item:
            continue

        if "=" not in item:
            raise ValueError(f"{option}: '{item}' must look like TRAIT=VALUE.")

        name, value = item.split("=", 1)
        name, value = name.strip(), value.strip()

        if name not in TRAITS:
            raise ValueError(
                f"{option}: unknown trait '{name}'. Valid traits: {', '.join(TRAITS)}."
            )

        if name in result:
            raise ValueError(f"{option}: trait '{name}' is given more than once.")

        if kind == "direction":
            mapping = {"-1": -1, "+1": 1, "1": 1, "-": -1, "+": 1}

            if value not in mapping:
                raise ValueError(
                    f"{option}: direction of {name} must be -1 or +1, got '{value}'."
                )

            result[name] = mapping[value]
        else:
            try:
                number = float(value)
            except ValueError:
                raise ValueError(f"{option}: weight of {name} must be a number, got '{value}'.")

            if not np.isfinite(number) or number < 0:
                raise ValueError(f"{option}: weight of {name} must be >= 0, got '{value}'.")

            result[name] = number

    return result


def resolve_path(base_dir, path_value):
    """Absolute paths are used as-is; relative ones are taken from base_dir."""
    path = Path(path_value).expanduser()
    return path if path.is_absolute() else (base_dir / path).resolve()


def parse_trait_names(text):
    """Parse 'D105,BF105,...' and validate count, characters and uniqueness."""
    names = [x.strip() for x in re.split(r"[,;]+", text.strip()) if x.strip()]

    if not 1 <= len(names) <= MAX_TRAITS:
        raise ValueError(
            f"--traits: {len(names)} trait names given; between 1 and {MAX_TRAITS} are allowed."
        )

    for name in names:
        if not TRAIT_NAME_RE.match(name):
            raise ValueError(
                f"--traits: invalid name '{name}'. Use letters, digits, '_', '.', '-' "
                "(no spaces, commas or '='), starting with a letter or digit."
            )

    if len(set(names)) != len(names):
        dup = sorted({n for n in names if names.count(n) > 1})
        raise ValueError(f"--traits: duplicate names: {', '.join(dup)}.")

    return names


def configure_traits(names):
    """Make `names` the active traits (in place). Weights start at 1; directions
    are taken from the defaults when the name is a default trait."""
    TRAITS[:] = names
    DIRECTION.clear()
    DIRECTION.update({t: DEFAULT_DIRECTION[t] for t in names if t in DEFAULT_DIRECTION})
    WEIGHTS.clear()
    WEIGHTS.update({t: 1.0 for t in names})
    EBV_COLS.clear()
    EBV_COLS.update({t: f"EBV_{t}" for t in names})
    ACC_COLS.clear()
    ACC_COLS.update({t: f"ACC_{t}" for t in names})


def require_file(path, description):
    if not path.is_file():
        raise FileNotFoundError(f"{description} not found: {path}")


def validate_configuration(min_acc):
    if not 1 <= len(TRAITS) <= MAX_TRAITS:
        raise ValueError(f"Between 1 and {MAX_TRAITS} traits are allowed.")

    missing_dir = [t for t in TRAITS if t not in DIRECTION]

    if missing_dir:
        raise ValueError(
            "Selection direction missing for: " + ", ".join(missing_dir)
            + f". Add --directions {missing_dir[0]}=-1 (lower is better) or "
            + f"{missing_dir[0]}=+1 (higher is better), and the same for the other listed traits."
        )

    if set(TRAITS) != set(DIRECTION):
        raise ValueError("DIRECTION must define exactly the configured traits.")

    if set(TRAITS) != set(WEIGHTS):
        raise ValueError("WEIGHTS must define exactly the configured traits.")

    for trait, direction in DIRECTION.items():
        if direction not in (-1, 1):
            raise ValueError(f"Direction for {trait} must be -1 or +1.")

    for trait, weight in WEIGHTS.items():
        if not np.isfinite(weight) or weight < 0:
            raise ValueError(f"Weight for {trait} must be finite and non-negative.")

    if all(weight == 0 for weight in WEIGHTS.values()):
        raise ValueError("At least one trait weight must be greater than zero.")

    if min_acc is not None and not 0 <= min_acc <= 1:
        raise ValueError("--min-acc must be between 0 and 1.")


# ============================================================
# 3. GENETIC COVARIANCE MATRIX FROM THE BLUPF90 LOG
# ============================================================

def _parse_numeric_line(line):
    """Parse a line made only of numbers (supports 0.4E-01 and Fortran D)."""
    tokens = line.strip().split()

    if not tokens:
        return None

    values = []

    for token in tokens:
        token = token.replace("D", "E").replace("d", "e")

        try:
            value = float(token)
        except ValueError:
            return None

        if not np.isfinite(value):
            return None

        values.append(value)

    return values


def extract_genetic_covariance(log_file, effect_number, n_traits):
    """
    First complete n_traits x n_traits numeric block after the last
    'Genetic variance(s) for effect N' heading. Blank lines are ignored.
    """
    log_text = log_file.read_text(encoding="utf-8", errors="replace")

    heading_pattern = re.compile(
        rf"Genetic\s+variance\(s\)\s+for\s+effect\s+{effect_number}\b",
        flags=re.IGNORECASE,
    )
    matches = list(heading_pattern.finditer(log_text))

    if not matches:
        available = sorted(
            set(re.findall(
                r"Genetic\s+variance\(s\)\s+for\s+effect\s+(\d+)",
                log_text, flags=re.IGNORECASE)),
            key=int,
        )
        raise ValueError(
            f"Could not find 'Genetic variance(s) for effect {effect_number}' "
            f"in {log_file}. Effects found in the log: {available}"
        )

    section = log_text[matches[-1].end():]

    numeric_rows = []
    size = None
    matrix = None

    for line in section.splitlines():
        if not line.strip():
            continue  # ignore blank lines

        row = _parse_numeric_line(line)

        if row is None:
            numeric_rows, size = [], None   # text line: restart the search
            continue

        if size is None or len(row) != size:
            numeric_rows, size = [], len(row)   # (re)start with this row length

        numeric_rows.append(row)

        if len(numeric_rows) == size:
            matrix = np.asarray(numeric_rows, dtype=float)
            break

    if matrix is None:
        raise ValueError(
            f"Could not parse a complete covariance matrix for effect "
            f"{effect_number} in {log_file}."
        )

    if matrix.shape[0] != n_traits:
        raise ValueError(
            f"The genetic covariance matrix of effect {effect_number} in {log_file.name} "
            f"is {matrix.shape[0]} x {matrix.shape[0]} (the model has {matrix.shape[0]} traits), "
            f"but {n_traits} trait name(s) are configured ({', '.join(TRAITS)}). "
            "Use --traits to give exactly one name per trait, in the order of the model."
        )

    if not np.allclose(matrix, matrix.T, rtol=1e-3, atol=1e-6):
        raise ValueError(
            "The extracted covariance matrix is not symmetric. "
            "Check the effect number and the log format."
        )

    matrix = (matrix + matrix.T) / 2.0  # remove rounding asymmetry

    variances = np.diag(matrix)

    if np.any(variances <= 0):
        bad = [TRAITS[i] for i, v in enumerate(variances) if v <= 0]
        raise ValueError(f"Non-positive genetic variance for: {bad}")

    eigenvalues = np.linalg.eigvalsh(matrix)
    tolerance = max(1e-8, np.max(np.abs(eigenvalues)) * 1e-6)

    if np.min(eigenvalues) < -tolerance:
        raise ValueError(
            "The extracted genetic covariance matrix is not positive "
            "semidefinite. Verify the parsed matrix."
        )

    return matrix


# ============================================================
# 4. LOAD AND CLEAN EBV / ACC DATA
# ============================================================

def _to_numeric_checked(df, col, label):
    original = df[col]
    numeric = pd.to_numeric(original, errors="coerce")
    invalid = original.notna() & numeric.isna()

    if invalid.any():
        examples = original[invalid].astype(str).head(5).tolist()
        raise ValueError(f"Non-numeric {label} values in '{col}': {examples}")

    return numeric


def load_and_clean_data(ebv_file, min_acc=None, acc_traits=None):
    df = pd.read_excel(ebv_file)

    required = ([ID_COL] + [EBV_COLS[t] for t in TRAITS]
                + [ACC_COLS[t] for t in TRAITS])
    missing = [c for c in required if c not in df.columns]

    if missing:
        raise ValueError("Missing required Excel columns: " + ", ".join(missing))

    original_rows = len(df)

    if df[ID_COL].isna().any():
        raise ValueError(f"Column '{ID_COL}' contains blank IDs.")

    df[ID_COL] = df[ID_COL].astype(str).str.strip()

    if (df[ID_COL] == "").any():
        raise ValueError(f"Column '{ID_COL}' contains blank IDs.")

    duplicated = df.loc[df[ID_COL].duplicated(keep=False), ID_COL]

    if not duplicated.empty:
        examples = duplicated.drop_duplicates().head(10).tolist()
        raise ValueError(
            "Duplicate IDs found. Each animal must appear only once. "
            f"Examples: {examples}"
        )

    for trait in TRAITS:
        col = EBV_COLS[trait]
        df[col] = _to_numeric_checked(df, col, "EBV")

        if np.isinf(df[col].dropna().to_numpy(dtype=float)).any():
            raise ValueError(f"EBV column '{col}' contains infinite values.")

    for trait in TRAITS:
        col = ACC_COLS[trait]
        numeric = _to_numeric_checked(df, col, "ACC")
        bad_range = numeric.notna() & ((numeric < 0) | (numeric > 1))

        if bad_range.any():
            raise ValueError(
                f"ACC values in '{col}' must be between 0 and 1. "
                f"Examples: {numeric[bad_range].head(5).tolist()}"
            )

        df[col] = numeric

    # Placeholder rows (unknown parent groups): all EBVs and all ACCs are 0.
    ebv_zero = df[[EBV_COLS[t] for t in TRAITS]].fillna(0).eq(0).all(axis=1)
    acc_zero = df[[ACC_COLS[t] for t in TRAITS]].fillna(0).eq(0).all(axis=1)
    placeholder = ebv_zero & acc_zero
    placeholder_rows = int(placeholder.sum())
    df = df.loc[~placeholder].copy()

    # A complete six-trait index requires all six EBVs.
    complete = df[[EBV_COLS[t] for t in TRAITS]].notna().all(axis=1)
    missing_ebv_rows = int((~complete).sum())
    df = df.loc[complete].copy()

    acc_filter_traits = list(acc_traits) if acc_traits else list(TRAITS)

    if min_acc is not None:
        acc_values = df[[ACC_COLS[t] for t in acc_filter_traits]]
        keep = acc_values.notna().all(axis=1) & acc_values.ge(min_acc).all(axis=1)
        filtered_acc_rows = int((~keep).sum())
        df = df.loc[keep].copy()
    else:
        filtered_acc_rows = 0

    if df.empty:
        raise ValueError("No animals remain after QC. Check the input and filters.")

    qc = {
        "Input file": str(ebv_file),
        "Original rows": original_rows,
        "Placeholder rows removed": placeholder_rows,
        "Rows removed due to missing EBV": missing_ebv_rows,
        "Rows removed by ACC filter": filtered_acc_rows,
        "Animals included in index": len(df),
        "ACC filter": "Disabled" if min_acc is None else min_acc,
        "ACC filter traits": ("n/a" if min_acc is None
                              else ", ".join(acc_filter_traits)),
    }

    return df.reset_index(drop=True), qc


# ============================================================
# 5. BIRTH YEAR FROM THE .DAT FILE
# ============================================================

def read_birth_years(path, id_col, year_col, century):
    """Series: ID -> birth year (int), from the .dat file."""
    if id_col < 1 or year_col < 1:
        raise ValueError("Column numbers are 1-based and must be >= 1.")

    cols = sorted({id_col - 1, year_col - 1})

    dat = pd.read_csv(
        path, sep=r"\s+", header=None, usecols=cols,
        dtype=str, on_bad_lines="skip", engine="c",
    )
    dat.columns = cols  # usecols returns columns in file order

    dat = pd.DataFrame({
        "id": dat[id_col - 1],
        "raw": dat[year_col - 1],
    }).dropna()

    dat["id"] = dat["id"].str.strip()

    yy = dat["raw"].str.strip().str[:2]
    valid = yy.str.fullmatch(r"\d{2}")
    n_bad = int((~valid).sum())

    if n_bad:
        examples = dat.loc[~valid, "raw"].head(5).tolist()
        print(f"WARNING: {n_bad:,} rows in the .dat file have a non-numeric "
              f"start in column {year_col} (e.g. {examples}); ignored.")

    dat = dat.loc[valid].copy()
    dat["Year"] = (century + yy[valid]).astype(int)

    # One animal may have several records; the year should be identical.
    per_id = dat.groupby("id")["Year"].agg(["first", "nunique"])
    n_conflict = int((per_id["nunique"] > 1).sum())

    if n_conflict:
        print(f"WARNING: {n_conflict:,} IDs have more than one birth year in "
              f"the .dat file; the first one is used.")

    return per_id["first"].rename("Year")


def attach_birth_year(df, years, label):
    out = df.merge(years, left_on=ID_COL, right_index=True, how="left")

    n_all, n_match = len(out), int(out["Year"].notna().sum())
    print(f"[{label}] Birth year found for {n_match:,} of {n_all:,} animals "
          f"({100 * n_match / max(n_all, 1):.1f}%).")

    if n_match == 0:
        raise ValueError(
            "No ID was found in the .dat file. Check --dat-id-col "
            "(use `head` on the .dat file)."
        )

    if n_match < 0.8 * n_all:
        print("NOTE: fewer than 80% matched. Animals without phenotype "
              "records (e.g. ancestors) are not in the .dat file; if this "
              "seems too low, check --dat-id-col.")

    # Implausible years are treated as unknown.
    bad = out["Year"].notna() & ((out["Year"] < YEAR_MIN) | (out["Year"] > YEAR_MAX))

    if bad.any():
        print(f"WARNING: {int(bad.sum()):,} animals have a year outside "
              f"{YEAR_MIN}-{YEAR_MAX}; treated as unknown.")
        out.loc[bad, "Year"] = np.nan

    return out


# ============================================================
# 6. SELECTION INDEX
# ============================================================

def calculate_index(df, covariance_matrix, ref_min_year=None, sd_mode="group"):
    result = df.copy()

    genetic_sd = np.sqrt(np.diag(covariance_matrix))

    # Trait-order sanity check: SD(EBV) / genetic SD is usually below ~1 and of
    # similar size across traits. Strong genetic trends can push it higher.
    ebv_sd = np.array([result[EBV_COLS[t]].std(ddof=0) for t in TRAITS])
    sd_ratio = ebv_sd / genetic_sd

    print("Trait-order check: SD(EBV) / genetic SD")
    for trait, ratio in zip(TRAITS, sd_ratio):
        print(f"  {trait:8s}: {ratio:.2f}")

    if np.any(sd_ratio > 1.2) or sd_ratio.max() > 4 * max(sd_ratio.min(), 1e-12):
        print("  WARNING: ratios look inconsistent. Verify that the EBV column "
              "order equals trait 1..6 of the model and that --genetic-effect "
              "is the animal effect (a strong genetic trend can also cause "
              "this).")
    print()

    for i, trait in enumerate(TRAITS):
        z_ebv = result[EBV_COLS[trait]] / genetic_sd[i]
        result[f"ZEBV_{trait}"] = z_ebv
        result[f"SI_Contribution_{trait}"] = WEIGHTS[trait] * DIRECTION[trait] * z_ebv

    result["SI_Raw"] = result[[f"SI_Contribution_{t}" for t in TRAITS]].sum(axis=1)

    # ---- Reference group for SI_100 ----
    if ref_min_year is not None:
        ref_mask = (result["Year"] >= ref_min_year).fillna(False)
        ref_label = f"Year >= {ref_min_year}"
    else:
        ref_mask = pd.Series(True, index=result.index)
        ref_label = "All animals in file"

    ref_si = result.loc[ref_mask, "SI_Raw"]
    n_ref = len(ref_si)

    if n_ref < 2:
        raise ValueError(
            f"Reference group has {n_ref} animal(s) ({ref_label}). "
            "Need at least 2 to standardize."
        )

    if n_ref < 30:
        print(f"WARNING: reference group has only {n_ref} animals; the mean/SD "
              "used for SI_100 may be unstable.\n")

    ref_mean = ref_si.mean()
    ref_sd = ref_si.std(ddof=0)

    # Genetic SD of the index: sigma_I = sqrt(a' G a), a = w * dir / sdG
    a = np.array([WEIGHTS[t] * DIRECTION[t] / genetic_sd[i]
                  for i, t in enumerate(TRAITS)])
    sigma_i = float(np.sqrt(a @ covariance_matrix @ a))

    if sd_mode == "genetic":
        scale, scale_label = sigma_i, "Genetic SD of index, sqrt(a'Ga)"
    else:
        scale, scale_label = ref_sd, "SD of SI_Raw in reference group"

    if not np.isfinite(scale) or scale <= 0:
        raise ValueError(
            "Selection index scale is zero or invalid. The animals may have "
            "identical EBVs or the input is unsuitable."
        )

    result["SI_100"] = 100 + 10 * (result["SI_Raw"] - ref_mean) / scale

    # Higher SI is better under the configured directions.
    result["Rank"] = result["SI_Raw"].rank(method="min", ascending=False).astype(int)

    n = len(result)
    result["Percentile"] = (100 * (n - result["Rank"]) / (n - 1)) if n > 1 else 100.0

    result = result.sort_values(
        by=["SI_Raw", ID_COL], ascending=[False, True], kind="mergesort"
    ).reset_index(drop=True)

    genetic_parameters = pd.DataFrame({
        "Trait": TRAITS,
        "EBV column": [EBV_COLS[t] for t in TRAITS],
        "ACC column": [ACC_COLS[t] for t in TRAITS],
        "Direction": [DIRECTION[t] for t in TRAITS],
        "Relative weight": [WEIGHTS[t] for t in TRAITS],
        "Genetic variance": np.diag(covariance_matrix),
        "Genetic SD": genetic_sd,
        "SD of EBV": ebv_sd,
        "SD(EBV) / Genetic SD": sd_ratio,
    })

    covariance_df = pd.DataFrame(covariance_matrix, index=TRAITS, columns=TRAITS)
    covariance_df.index.name = "Trait"

    if ref_min_year is not None:
        ref_si100 = result.loc[(result["Year"] >= ref_min_year).fillna(False), "SI_100"]
    else:
        ref_si100 = result["SI_100"]

    index_summary = {
        "Animals included": n,
        "Reference group": ref_label,
        "Animals in reference group": n_ref,
        "SI_100 scale": scale_label,
        "Trait weights (direction)": ", ".join(
            f"{t} {WEIGHTS[t]:g} ({'-' if DIRECTION[t] < 0 else '+'})" for t in TRAITS),
        "Mean raw SI (reference group)": ref_mean,
        "SD raw SI (reference group, population)": ref_sd,
        "Genetic SD of index (sigma_I)": sigma_i,
        "Scale value used for SI_100": scale,
        "Mean SI_100 (reference group)": ref_si100.mean(),
        "SD SI_100 (reference group, population)": ref_si100.std(ddof=0),
        "Mean SI_100 (all animals)": result["SI_100"].mean(),
        "Highest SI_100": result["SI_100"].max(),
        "Lowest SI_100": result["SI_100"].min(),
    }

    return result, genetic_parameters, covariance_df, index_summary


# ============================================================
# 7. EXPORT EXCEL
# ============================================================

def export_results(output_file, ranked_df, genetic_parameters, covariance_df,
                   qc_summary, index_summary):
    output_file.parent.mkdir(parents=True, exist_ok=True)

    qc_df = pd.DataFrame(list(qc_summary.items()), columns=["QC item", "Value"])
    summary_df = pd.DataFrame(list(index_summary.items()),
                              columns=["Index summary", "Value"])

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        ranked_df.to_excel(writer, sheet_name="Ranked_Animals", index=False)
        genetic_parameters.to_excel(writer, sheet_name="Genetic_Parameters", index=False)
        covariance_df.to_excel(writer, sheet_name="Genetic_Covariance", index=True)
        qc_df.to_excel(writer, sheet_name="QC_Summary", index=False)
        summary_df.to_excel(writer, sheet_name="Index_Summary", index=False)

        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for column_cells in worksheet.columns:
                values = [len(str(c.value)) for c in column_cells[:1000]
                          if c.value is not None]
                width = min(max(values, default=10) + 2, 35)
                worksheet.column_dimensions[column_cells[0].column_letter].width = width


# ============================================================
# 8. TREND CHARTS (PDF)
# ============================================================

def summarize_by_year(df, min_n):
    """Mean, SE and N per birth year for each trait."""
    out = {}
    for t in TRAITS:
        g = df.groupby("Year")[EBV_COLS[t]].agg(["mean", "std", "count"])
        g["se"] = g["std"].fillna(0) / np.sqrt(g["count"])
        out[t] = g[g["count"] >= min_n]
    return out


def grid_shape(n):
    """(rows, columns) of the chart grid for n traits (1 to 12)."""
    return {1: (1, 1), 2: (1, 2), 3: (1, 3), 4: (2, 2), 5: (2, 3), 6: (2, 3),
            7: (3, 3), 8: (3, 3), 9: (3, 3)}.get(n, (3, 4))


def plot_pdf(df, summary, out_pdf, meta_text):
    n_tr = len(TRAITS)
    nrows, ncols = grid_shape(n_tr)
    small = n_tr > 6
    fig, axes = plt.subplots(nrows, ncols, figsize=(11.69, 8.27), squeeze=False)  # A4 landscape
    fig.suptitle("Mean EBV by birth year (mean \u00b1 1 SE)", fontsize=14,
                 fontweight="bold", y=0.985)

    for ax in axes.ravel()[n_tr:]:
        ax.set_axis_off()   # unused cells of the grid

    for ax, trait in zip(axes.ravel(), TRAITS):
        g = summary[trait]
        hint = "lower = better" if DIRECTION[trait] < 0 else "higher = better"
        ax.set_title(f"{trait}  ({hint})", fontsize=9.5 if small else 11)

        if g.empty:
            ax.text(0.5, 0.5, "No year with enough animals",
                    ha="center", va="center", transform=ax.transAxes)
            ax.set_axis_off()
            continue

        years = g.index.to_numpy()
        ax.fill_between(years, g["mean"] - g["se"], g["mean"] + g["se"],
                        alpha=0.25, linewidth=0)
        ax.plot(years, g["mean"], marker="o", linewidth=1.8)

        ax.set_xticks(years)
        ax.set_xticklabels(
            [f"{y}\nn={int(n)}" for y, n in zip(years, g["count"])],
            fontsize=5.5 if small else 7)
        ax.tick_params(axis="y", labelsize=7 if small else 9)
        ax.set_ylabel("Mean EBV", fontsize=8 if small else 9)
        ax.grid(alpha=0.3)

        # Linear trend over all animals in the plotted years (units / year).
        sub = df[df["Year"].isin(years)].dropna(subset=[EBV_COLS[trait]])

        if sub["Year"].nunique() >= 2:
            slope = np.polyfit(sub["Year"], sub[EBV_COLS[trait]], 1)[0]
            # Label goes in the empty corner: decreasing line -> top right.
            x_pos, h_align = (0.97, "right") if slope < 0 else (0.03, "left")
            ax.text(x_pos, 0.96, f"Trend: {slope:+.3f} / year",
                    transform=ax.transAxes, va="top", ha=h_align, fontsize=7 if small else 8.5,
                    bbox=dict(boxstyle="round", facecolor="white", alpha=0.8))

    fig.text(0.5, 0.012, meta_text, ha="center", fontsize=7.5, color="0.35")
    fig.tight_layout(rect=(0, 0.03, 1, 0.96))

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    with PdfPages(out_pdf) as pdf:
        pdf.savefig(fig)
        pdf.infodict()["Title"] = "Mean EBV by birth year"
    plt.close(fig)


# ============================================================
# 9. MAIN
# ============================================================

def main():
    args = parse_args()

    if args.traits:
        configure_traits(parse_trait_names(args.traits))

    if args.acc_traits:
        unknown = [t for t in args.acc_traits if t not in TRAITS]
        if unknown:
            raise ValueError(
                f"--acc-traits: unknown trait(s): {', '.join(unknown)}. "
                f"Valid traits: {', '.join(TRAITS)}."
            )

    if args.weights:
        WEIGHTS.update(parse_trait_map(args.weights, "--weights", "weight"))
    if args.directions:
        DIRECTION.update(parse_trait_map(args.directions, "--directions", "direction"))

    validate_configuration(args.min_acc)

    base = args.base_dir.expanduser().resolve()
    ebv_file = resolve_path(base, args.ebv_file)
    log_file = resolve_path(base, args.log_file)
    dat_file = resolve_path(base, args.dat_file)
    out_xlsx = resolve_path(base, args.out_xlsx)
    out_pdf = resolve_path(base, args.out_pdf)

    require_file(ebv_file, "EBV Excel file")
    require_file(log_file, "BLUPF90 log file")
    require_file(dat_file, ".dat file")

    print("=" * 65)
    print("SELECTION INDEX + EBV TREND")
    print("=" * 65)
    print(f"EBV file       : {ebv_file}")
    print(f"BLUPF90 log    : {log_file}  (effect {args.genetic_effect})")
    print(f".dat file      : {dat_file}  (ID col {args.dat_id_col}, "
          f"year col {args.dat_year_col})")
    print(f"Reference group: "
          f"{'Year >= ' + str(args.ref_min_year) if args.ref_min_year is not None else 'All animals'}")
    print(f"SI_100 scale   : {args.sd_mode}")
    print(f"ACC filter     : {args.min_acc if args.min_acc is not None else 'Disabled'}")
    print(f"Outputs        : {out_xlsx.name}, {out_pdf.name}")
    print(f"Traits ({len(TRAITS)}), in the order of the BLUP model; weights per 1 genetic SD:")
    total_w = sum(WEIGHTS.values())
    for t in TRAITS:
        print(f"  {t:8s} weight {WEIGHTS[t]:<6g} ({100 * WEIGHTS[t] / total_w:4.1f}%)  "
              f"{'lower is better (-)' if DIRECTION[t] < 0 else 'higher is better (+)'}")
    print()

    # 1. Genetic covariance matrix
    G = extract_genetic_covariance(log_file, args.genetic_effect, len(TRAITS))
    print("Genetic SD per trait: " + ", ".join(
        f"{t}={s:.3f}" for t, s in zip(TRAITS, np.sqrt(np.diag(G)))) + "\n")

    # 2. EBV data: one copy for the index (optional ACC filter) and one
    #    without the ACC filter for the charts (the filter would bias trends).
    df_idx, qc = load_and_clean_data(ebv_file, args.min_acc, args.acc_traits)
    df_plot, _ = load_and_clean_data(ebv_file, None, None)

    # 3. Birth years
    years = read_birth_years(dat_file, args.dat_id_col, args.dat_year_col,
                             args.century)
    df_idx = attach_birth_year(df_idx, years, "index")
    df_plot = attach_birth_year(df_plot, years, "charts")
    print()

    # 4. Selection index
    ranked, gen_par, cov_df, idx_summary = calculate_index(
        df_idx, G, ref_min_year=args.ref_min_year, sd_mode=args.sd_mode)
    ranked["Year"] = ranked["Year"].astype("Int64")

    # 5. Excel
    export_results(out_xlsx, ranked, gen_par, cov_df, qc, idx_summary)

    print("QC summary:")
    for k, v in qc.items():
        print(f"  {k}: {v}")

    print("\nTop 10 animals:")
    print(ranked[[ID_COL, "Year", "SI_Raw", "SI_100", "Rank"]]
          .head(10).to_string(index=False))
    print(f"\nExcel saved: {out_xlsx}\n")

    # 6. PDF charts
    df_plot = df_plot.dropna(subset=["Year"]).copy()
    df_plot["Year"] = df_plot["Year"].astype(int)

    if df_plot.empty:
        raise ValueError("No animals with a known birth year for the charts.")

    counts = df_plot["Year"].value_counts().sort_index()
    print("Animals per birth year (charts):")
    for y, n in counts.items():
        flag = "" if n >= args.min_n else f"   (< min-n {args.min_n}, not plotted)"
        print(f"  {y}: {n:,}{flag}")

    summary = summarize_by_year(df_plot, args.min_n)

    print("\nMean EBV per year:")
    print(pd.DataFrame({t: summary[t]["mean"] for t in TRAITS})
          .round(3).to_string())

    meta = (f"Source: {ebv_file.name} + {dat_file.name} | "
            f"animals plotted: {len(df_plot):,} | years with n >= {args.min_n}")
    plot_pdf(df_plot, summary, out_pdf, meta)
    print(f"\nPDF saved: {out_pdf}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)