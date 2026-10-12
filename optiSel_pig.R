#!/usr/bin/env Rscript
# ======================================================================
# Pig breeding OCS (optiSel) - command-line version
#
# Example
#   Rscript ocs_pig.R \
#       --input-dir ../data_260929/blup_LL_T2613 \
#       --ped-file PedLL.ped --si-file EBV_ACC_SI.xlsx \
#       --sires 40 --dams 400 --birth-year 2026
#   -> results in ./data_260929_LL_OCS/
#
# Run  Rscript ocs_pig.R --help  for all options.
# One-time install:  install.packages(c("optiSel", "dplyr", "readxl", "clue"))
# ======================================================================
options(rgl.useNULL = TRUE)
Sys.setenv(RGL_USE_NULL = "TRUE")
suppressPackageStartupMessages({
  library(optiSel)
  library(dplyr)
  library(readxl)
})

USAGE <- "
Usage: Rscript ocs_pig.R --ped-file F --si-file F --sires N --dams N
                         (--birth-year YEAR|all | --cand-file F) [options]

Required
  --ped-file F        pedigree, no header: ID Sire Dam BirthDate(YYYYMMDD) Sex(F/M)
  --si-file F         Excel (first sheet unless --sheet) with columns ID and the SI column
  --sires N           number of sires (boars) to use
  --dams N            number of dams (sows) to mate; one mating per dam
  and one of
  --birth-year Y      candidates born in year Y, or 'all' for every animal with an SI
  --cand-file F       text/csv with a header line and an ID column: the exact
                      animals available (overrides --birth-year)

Optional
  --input-dir D       folder holding the input files             [.]
  --si-column C       SI column in the Excel file                [SI_100]
  --sheet S           Excel sheet name or number                 [first sheet]
  --min-matings N     minimum matings per selected sire          [1]
  --max-matings N     maximum matings per sire                   [30]
  --ne N              target effective size (kinship rise per
                      generation = 1/(2*Ne))                     [50]
  --pop-size N        population size used by optiSel (drift)    [1000]
  --alternatives N    alternative sires listed for each dam      [2]
  --out-root D        folder in which the output folder is made  [.]
  --out-name NAME     output folder name; default is
                      <parent of input-dir>_<line>_OCS, e.g. data_260929_LL_OCS
  --tag T             line/breed label in the default name       [from input-dir]
  --help              show this text

File names in --ped-file/--si-file/--cand-file are taken relative to --input-dir
(absolute paths are used as given).
"

# ------------------------------------------------------------ arguments
defaults <- list(
  input_dir = ".", ped_file = NULL, si_file = NULL, sheet = NULL,
  si_column = "SI_100", sires = NULL, dams = NULL,
  birth_year = NULL, cand_file = NULL,
  min_matings = 1, max_matings = 30, ne = 50, pop_size = 1000, alternatives = 2,
  out_root = ".", out_name = NULL, tag = NULL
)
numeric_opts <- c("sires", "dams", "min_matings", "max_matings", "ne",
                  "pop_size", "alternatives")

parse_cli <- function() {
  a <- commandArgs(trailingOnly = TRUE)
  if (length(a) == 0 || any(a %in% c("--help", "-h"))) {
    cat(USAGE)
    quit(status = 0)
  }
  opt <- defaults
  i <- 1
  while (i <= length(a)) {
    if (!startsWith(a[i], "--")) stop("Unexpected argument: ", a[i])
    key <- gsub("-", "_", sub("^--", "", a[i]))
    if (!key %in% names(defaults)) stop("Unknown option: ", a[i])
    if (i == length(a)) stop("Missing value for ", a[i])
    val <- a[i + 1]
    if (key %in% numeric_opts) {
      num <- suppressWarnings(as.numeric(val))
      if (is.na(num)) stop(a[i], " needs a number, got '", val, "'")
      opt[[key]] <- num
    } else {
      opt[[key]] <- val
    }
    i <- i + 2
  }

  for (k in c("ped_file", "si_file", "sires", "dams")) {
    if (is.null(opt[[k]])) stop("--", gsub("_", "-", k), " is required (see --help).")
  }
  if (is.null(opt$birth_year) && is.null(opt$cand_file)) {
    stop("Give the candidates with --birth-year (a year or 'all') or --cand-file.")
  }
  if (!is.null(opt$birth_year)) {
    opt$birth_year <- if (tolower(opt$birth_year) == "all") NA_integer_ else {
      by <- suppressWarnings(as.integer(opt$birth_year))
      if (is.na(by)) stop("--birth-year must be a year (e.g. 2026) or 'all'.")
      by
    }
  }
  for (k in c("sires", "dams", "max_matings", "alternatives")) {
    if (opt[[k]] < 1 || opt[[k]] != round(opt[[k]])) stop("--", gsub("_", "-", k), " must be a positive integer.")
  }
  if (opt$min_matings < 0 || opt$min_matings != round(opt$min_matings)) stop("--min-matings must be an integer >= 0.")
  if (opt$min_matings > opt$max_matings) stop("--min-matings is larger than --max-matings.")
  if (opt$ne <= 0 || opt$pop_size <= 0) stop("--ne and --pop-size must be positive.")
  opt
}

opt <- parse_cli()

# ------------------------------------------------------------ paths
input_dir <- normalizePath(opt$input_dir, mustWork = TRUE)
path_in <- function(f) if (grepl("^(/|~|[A-Za-z]:)", f)) path.expand(f) else file.path(input_dir, f)

ped_path <- path_in(opt$ped_file)
si_path  <- path_in(opt$si_file)
cand_path <- if (!is.null(opt$cand_file)) path_in(opt$cand_file) else NULL
for (p in c(ped_path, si_path, cand_path)) if (!file.exists(p)) stop("File not found: ", p)

if (!requireNamespace("clue", quietly = TRUE)) {
  stop("Package 'clue' is needed for the mating allocation: install.packages('clue')")
}

# Output folder name carries the input information: <parent>_<line>_OCS
default_out_name <- function(input_dir, tag) {
  leaf   <- basename(input_dir)
  parent <- basename(dirname(input_dir))
  if (is.null(tag)) {
    parts <- strsplit(leaf, "_")[[1]]
    tag <- if (length(parts) >= 2 && tolower(parts[1]) == "blup") parts[2] else leaf
  }
  paste0(parent, "_", tag, "_OCS")
}
out_name <- if (!is.null(opt$out_name)) opt$out_name else default_out_name(input_dir, opt$tag)
dir.create(opt$out_root, showWarnings = FALSE, recursive = TRUE)
out_dir <- file.path(normalizePath(opt$out_root, mustWork = TRUE), out_name)
dir.create(out_dir, showWarnings = FALSE, recursive = TRUE)
out <- function(f) file.path(out_dir, f)

# ------------------------------------------------------------ log
log_con <- file(out("0_run_log.txt"), open = "wt")
sink(log_con, split = TRUE)
close_log <- function() { sink(); close(log_con) }

cat(strrep("=", 66), "\nPIG OCS (optiSel)   ", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n",
    strrep("=", 66), "\n", sep = "")
cat("R", paste(R.version$major, R.version$minor, sep = "."),
    "| optiSel", as.character(packageVersion("optiSel")), "\n")
cat("Input folder :", input_dir, "\nOutput folder:", out_dir, "\n\nSettings:\n")
show <- c(ped_file = ped_path, si_file = si_path, sheet = ifelse(is.null(opt$sheet), "first", opt$sheet),
          si_column = opt$si_column, sires = opt$sires, dams = opt$dams,
          birth_year = if (!is.null(cand_path)) "(cand-file)" else ifelse(is.na(opt$birth_year), "all", opt$birth_year),
          cand_file = ifelse(is.null(cand_path), "-", cand_path),
          min_matings = opt$min_matings, max_matings = opt$max_matings, ne = opt$ne,
          pop_size = opt$pop_size, alternatives = opt$alternatives)
for (k in names(show)) cat(sprintf("  %-12s %s\n", k, show[[k]]))
cat("\n")

# ------------------------------------------------------------ analysis
run <- function() {
  N_SIRES <- opt$sires
  N_DAMS  <- opt$dams
  N_MAT   <- N_DAMS                 # one mating per dam
  MIN_M   <- opt$min_matings
  MAX_M   <- opt$max_matings

  if (N_MAT < N_SIRES * MIN_M) stop("--dams is smaller than --sires x --min-matings.")
  if (N_MAT > N_SIRES * MAX_M) stop("Capacity too small: --sires x --max-matings = ",
                                    N_SIRES * MAX_M, " < --dams = ", N_DAMS,
                                    ". Raise --sires or --max-matings.")

  # ---- 1. pedigree
  ped0 <- read.table(
    ped_path, header = FALSE, sep = "",
    col.names = c("ID", "Sire", "Dam", "BirthDate", "Sex"),
    colClasses = "character", na.strings = c("", "NA"),
    stringsAsFactors = FALSE, comment.char = "", quote = ""
  )
  if (anyDuplicated(ped0$ID)) stop("Duplicate IDs in the pedigree.")

  ped0 <- ped0 %>%
    mutate(
      Sire = ifelse(is.na(Sire) | Sire %in% c("", "0"), NA, Sire),
      Dam  = ifelse(is.na(Dam)  | Dam  %in% c("", "0"), NA, Dam),
      Born = suppressWarnings(as.integer(substr(BirthDate, 1, 4))),
      Sex  = case_when(toupper(Sex) == "M" ~ "male", toupper(Sex) == "F" ~ "female",
                       TRUE ~ NA_character_)
    )
  if (anyNA(ped0$Born)) stop("Could not extract the birth year from column 4 of the pedigree.")
  if (anyNA(ped0$Sex))  warning("Some pedigree animals have unknown or invalid sex.")

  Pedig <- prePed(ped0 %>% select(Indiv = ID, Sire, Dam, Sex, Born))
  cat("Pedigree:", nrow(Pedig), "animals\n")

  # ---- 2. selection index
  si <- read_excel(si_path, sheet = if (is.null(opt$sheet)) 1 else {
    s <- suppressWarnings(as.integer(opt$sheet)); if (is.na(s)) opt$sheet else s
  }, .name_repair = "minimal")
  if (!all(c("ID", opt$si_column) %in% names(si))) {
    stop("SI file must contain columns ID and ", opt$si_column, " (use --si-column).")
  }
  si <- si %>%
    transmute(Indiv = as.character(ID), BV = suppressWarnings(as.numeric(.data[[opt$si_column]]))) %>%
    filter(is.finite(BV))
  if (anyDuplicated(si$Indiv)) stop("Duplicate IDs in the SI file.")

  # ---- 3. candidates and shortlist
  Phen_all <- Pedig %>% inner_join(si, by = "Indiv") %>% filter(Sex %in% c("male", "female"))

  if (!is.null(cand_path)) {
    first <- readLines(cand_path, n = 1, warn = FALSE)
    sep <- if (grepl(",", first)) "," else if (grepl("\t", first)) "\t" else ""
    cf <- read.table(cand_path, header = TRUE, sep = sep, colClasses = "character",
                     comment.char = "", quote = "\"", fill = TRUE, stringsAsFactors = FALSE)
    idc <- intersect(c("ID", "Indiv", "id"), names(cf))
    ids <- unique(trimws(cf[[if (length(idc)) idc[1] else 1]]))
    cat("Candidate list:", length(ids), "IDs;",
        sum(!ids %in% Phen_all$Indiv), "not found in pedigree/SI (ignored)\n")
    Phen_all <- Phen_all %>% filter(Indiv %in% ids)
  } else if (!is.na(opt$birth_year)) {
    Phen_all <- Phen_all %>% filter(Born == opt$birth_year)
  }
  if (nrow(Phen_all) == 0) stop("No animals with an SI match the chosen candidates.")

  males_all   <- Phen_all %>% filter(Sex == "male")   %>% arrange(desc(BV), Indiv)
  females_all <- Phen_all %>% filter(Sex == "female") %>% arrange(desc(BV), Indiv)
  cat("Candidates:", nrow(males_all), "males,", nrow(females_all), "females\n")

  if (nrow(males_all)   < N_SIRES) stop("Only ", nrow(males_all), " male candidates; --sires is ", N_SIRES, ".")
  if (nrow(females_all) < N_DAMS)  stop("Only ", nrow(females_all), " female candidates; --dams is ", N_DAMS, ".")

  # shortlist by SI; Born = 1 -> one non-overlapping generation
  Phen <- bind_rows(slice_head(males_all, n = N_SIRES), slice_head(females_all, n = N_DAMS)) %>%
    mutate(Born = 1, isCandidate = TRUE) %>% as.data.frame()
  male_ids <- Phen$Indiv[Phen$Sex == "male"]

  write.csv(
    Phen_all %>% filter(Indiv %in% Phen$Indiv) %>%
      transmute(ID = Indiv, Sex, BirthYear = Born, SI = BV) %>% arrange(Sex, desc(SI)),
    out("1_shortlist_candidates.csv"), row.names = FALSE)

  # ---- 4. kinship
  pKin <- pedIBD(Pedig, keep.only = Phen$Indiv)
  Kin_mating <- as.matrix(pKin)
  F_cand <- setNames(2 * diag(Kin_mating) - 1, rownames(Kin_mating))   # F = 2*f_ii - 1

  # ---- 5. optimum contributions
  cand <- candes(phen = Phen, N = opt$pop_size, pKin = pKin, quiet = TRUE, reduce.data = FALSE)
  current_pKin <- as.numeric(unlist(cand$mean["pKin"])[1])
  cat(sprintf("Mean kinship of the shortlist (now): %.5f\n", current_pKin))

  kinship_limit <- 1 - (1 - current_pKin) * (1 - 1 / (2 * opt$ne))

  # matings of a sire ~ 2 * N * oc  ->  oc between MIN/(2N) and MAX/(2N)
  con_base <- list(
    uniform = "female",
    lb = setNames(rep(MIN_M / (2 * N_MAT), length(male_ids)), male_ids),
    ub = setNames(rep(MAX_M / (2 * N_MAT), length(male_ids)), male_ids)
  )
  kmin <- opticont("min.pKin", cand, con_base, quiet = TRUE)$mean$pKin
  cat(sprintf("Smallest achievable kinship: %.5f | Ne-based ceiling: %.5f\n", kmin, kinship_limit))

  ub_used <- kinship_limit
  if (kinship_limit < kmin) {
    ub_used <- kmin + 1e-4
    warning("Ne-based kinship ceiling is below the smallest achievable kinship; using ",
            round(ub_used, 5), " instead. Raise --ne or --sires.")
  }

  ocs <- opticont("max.BV", cand, c(con_base, list(ub.pKin = ub_used)), quiet = TRUE)
  cat("OCS solver:", paste(ocs$info, collapse = " "), "\n")
  cat(sprintf("Expected mean SI: %.3f | expected mean kinship: %.5f (ceiling %.5f)\n",
              ocs$mean$BV, ocs$mean$pKin, ub_used))
  if (is.null(ocs$parent) || nrow(ocs$parent) == 0) stop("OCS returned no parent contributions.")

  # ---- 6. matings per animal
  parent <- as.data.frame(ocs$parent)
  parent$n <- noffspring(parent, N = N_MAT, random = FALSE)$nOff
  parent <- parent %>%
    left_join(Phen %>% select(Indiv, SI = BV), by = "Indiv") %>%
    arrange(desc(Sex), desc(oc))
  parent$F_inbreeding <- unname(F_cand[parent$Indiv])

  sire_table <- parent %>% filter(Sex == "male", n > 0)
  dam_table  <- parent %>% filter(Sex == "female", n > 0)

  if (sum(sire_table$n) != N_MAT || sum(dam_table$n) != N_MAT) stop("Matings do not add up to --dams.")
  if (any(dam_table$n != 1L)) stop("Each dam must have exactly one mating.")
  if (min(sire_table$n) < MIN_M || max(sire_table$n) > MAX_M) {
    stop("Sire matings fall outside [--min-matings, --max-matings] after rounding.")
  }
  write.csv(parent, out("2_OCS_parent_contributions.csv"), row.names = FALSE)

  # ---- 7. dam -> sire allocation (Hungarian algorithm, exact)
  sire_ids <- sire_table$Indiv
  dam_ids  <- dam_table$Indiv
  K_dm <- Kin_mating[dam_ids, sire_ids, drop = FALSE]
  if (any(!is.finite(K_dm))) stop("Kinship matrix contains missing or non-finite values.")

  sire_slots <- rep(seq_along(sire_ids), sire_table$n)
  asg <- clue::solve_LSAP(K_dm[, sire_slots, drop = FALSE])
  assigned_sire <- sire_ids[sire_slots][as.integer(asg)]

  mating_plan <- data.frame(
    Sire = assigned_sire, Dam = dam_ids, n = 1L,
    Kinship = K_dm[cbind(seq_along(dam_ids), match(assigned_sire, sire_ids))],
    SI_Sire = Phen$BV[match(assigned_sire, Phen$Indiv)],
    SI_Dam  = Phen$BV[match(dam_ids, Phen$Indiv)],
    stringsAsFactors = FALSE
  )
  mating_plan$SI_Mid <- (mating_plan$SI_Sire + mating_plan$SI_Dam) / 2
  mating_plan <- mating_plan %>% arrange(Sire, Dam)

  actual <- table(factor(mating_plan$Sire, levels = sire_ids))
  sire_qc <- data.frame(
    Indiv = sire_ids, OCS_contribution = sire_table$oc, SI = sire_table$SI,
    F_inbreeding = sire_table$F_inbreeding,
    PlannedMatings = sire_table$n, ActualMatings = as.integer(actual)
  )
  stopifnot(
    nrow(mating_plan) == length(dam_ids), !anyDuplicated(mating_plan$Dam),
    all(sire_qc$ActualMatings == sire_qc$PlannedMatings),
    max(sire_qc$ActualMatings) <= MAX_M
  )

  kin_random <- mean(K_dm %*% (sire_table$n / N_MAT))   # same quotas, random assignment
  kin_plan   <- mean(mating_plan$Kinship)

  write.csv(mating_plan, out("3_mating_plan.csv"), row.names = FALSE)
  write.csv(sire_qc, out("4_sire_mating_counts.csv"), row.names = FALSE)

  cat("\nDams assigned:", nrow(mating_plan), "| sires used:", nrow(sire_table), "\n")
  cat("Matings per sire: min", min(sire_qc$ActualMatings), "max", max(sire_qc$ActualMatings), "\n")
  cat(sprintf("Mean kinship of assigned pairs: %.5f  (random assignment: %.5f)\n", kin_plan, kin_random))

  # ---- 8. planned sire + alternatives for each dam
  alt <- lapply(seq_along(dam_ids), function(i) {
    others <- setdiff(sire_ids, assigned_sire[i])
    c(assigned_sire[i], head(others[order(K_dm[dam_ids[i], others], others)], opt$alternatives))
  })
  prio <- data.frame(DAM = dam_ids, SI_DAM = Phen$BV[match(dam_ids, Phen$Indiv)], stringsAsFactors = FALSE)
  for (k in seq_len(1 + opt$alternatives)) {
    s <- vapply(alt, function(x) if (length(x) >= k) x[k] else NA_character_, character(1))
    prio[[paste0("SIR_", k)]]     <- s
    prio[[paste0("KINSHIP_", k)]] <- ifelse(is.na(s), NA_real_,
      K_dm[cbind(match(dam_ids, rownames(K_dm)), match(s, colnames(K_dm)))])
    prio[[paste0("F_SIR_", k)]]   <- ifelse(is.na(s), NA_real_, unname(F_cand[s]))
  }
  write.csv(prio, out("5_dam_top_sires.csv"), row.names = FALSE, na = "")

  # ---- 9. QC summary
  items <- list(
    Input_folder = input_dir, Birth_year_filter = if (!is.null(cand_path)) "cand-file" else ifelse(is.na(opt$birth_year), "all", opt$birth_year),
    SI_column = opt$si_column,
    Candidates_male = nrow(males_all), Candidates_female = nrow(females_all),
    Target_sires = N_SIRES, Target_dams = N_DAMS, Target_matings = N_MAT,
    Sires_with_matings = nrow(sire_table), Dams_with_matings = nrow(dam_table),
    Sire_matings_min = min(sire_table$n), Sire_matings_max = max(sire_table$n),
    Mean_SI_shortlist_males = mean(Phen$BV[Phen$Sex == "male"]),
    Mean_SI_shortlist_females = mean(Phen$BV[Phen$Sex == "female"]),
    Kinship_current = current_pKin, Kinship_smallest_achievable = kmin,
    Kinship_upper_bound_used = ub_used, Target_Ne = opt$ne,
    Expected_mean_SI = ocs$mean$BV, Expected_mean_kinship = ocs$mean$pKin,
    Mean_kinship_assigned_pairs = kin_plan, Mean_kinship_random_assignment = kin_random,
    Mean_SI_midparent = mean(mating_plan$SI_Mid), Solver_info = paste(ocs$info, collapse = " ")
  )
  write.csv(data.frame(Metric = names(items), Value = vapply(items, as.character, character(1))),
            out("6_QC_summary.csv"), row.names = FALSE)

  cat("\nFiles written to:", out_dir, "\n")
  cat(paste0("  ", sort(list.files(out_dir)), collapse = "\n"), "\n")
}

withCallingHandlers(
  tryCatch(run(), error = function(e) {
    cat("\nERROR:", conditionMessage(e), "\n")
    close_log()
    quit(status = 1)
  }),
  warning = function(w) {
    cat("WARNING:", conditionMessage(w), "\n")
    invokeRestart("muffleWarning")
  }
)
close_log()
