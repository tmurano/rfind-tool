# =============================================================================
# RFind Analysis — Running Fisher (4-direction) enrichment scoring
# R implementation equivalent to the RFind Web Tool (JavaScript)
#
# Input:
#   1. FPKM/TPM expression matrix (gene x sample, CSV/Excel)
#   2. Sample info with ID and Dx columns (Dx="Control" required)
#   3. Reference index (Symbol + Fold Change, CSV)
#
# Output:
#   Per-sample RFind scores for each reference index
#
# Algorithm: Illumina BaseSpace Correlation Engine Running Fisher
#   (Tech Note 970-2014-007)
#
# Usage:
#   source("rfind_analysis.R")
#   results <- rfind(fpkm_file, demo_file, index_files)
# =============================================================================

library(readxl)

# ─── Parameters ───
THRESHOLD_FC <- 1.2    # |FC| < threshold → NA
CLIP_LIMIT   <- 300    # Score clipping ±300

# =============================================================================
# Running Fisher core
# =============================================================================

#' Compute Running Fisher score between two gene sets
#'
#' @param b1 data.frame with columns: Symbol, FC (reference index)
#' @param b2 data.frame with columns: Symbol, FC (sample fold changes)
#' @param P  integer, universe size (common genes between FPKM and index)
#' @param clip_limit numeric, score clipping limit (default 300)
#' @return list with: total (score), upup, updown, downup, downdown (each with score, pval, overlap)
running_fisher <- function(b1, b2, P, clip_limit = CLIP_LIMIT) {

  # Deduplicate and clean
  format_df <- function(df) {
    df <- df[!duplicated(tolower(trimws(df$Symbol))), ]
    df$Symbol <- tolower(trimws(df$Symbol))
    df <- df[df$FC != 0, ]
    df
  }

  b1 <- format_df(b1)
  b2 <- format_df(b2)

  # Rank genes by FC for a given direction
  rank_fc <- function(df, dir) {
    if (dir == "up") {
      df <- df[df$FC > 0, ]
      df <- df[order(-df$FC), ]
    } else {
      df <- df[df$FC < 0, ]
      df <- df[order(df$FC), ]
    }
    df
  }

  # Bidirectional Fisher enrichment for one direction pair
  fisher_bidirectional <- function(dir1, dir2) {
    ranked1 <- rank_fc(b1, dir1)
    ranked2 <- rank_fc(b2, dir2)

    if (nrow(ranked1) == 0 || nrow(ranked2) == 0) {
      return(list(score = 0, pval = 1, overlap = 0, bestM = 0))
    }

    K1 <- nrow(ranked2)
    K2 <- nrow(ranked1)
    set2 <- ranked2$Symbol
    set1 <- ranked1$Symbol

    # Direction 1: scan ranked1 against set2
    # Optimization: only compute phyper when overlap increases (hit-only)
    cum1 <- 0; p1 <- 1; m1 <- 0
    for (i in seq_len(nrow(ranked1))) {
      if (ranked1$Symbol[i] %in% set2) {
        cum1 <- cum1 + 1
        p <- phyper(cum1 - 1, K1, P - K1, i, lower.tail = FALSE)
        if (p < p1) { p1 <- p; m1 <- i }
      }
    }

    # Direction 2: scan ranked2 against set1
    cum2 <- 0; p2 <- 1; m2 <- 0
    for (i in seq_len(nrow(ranked2))) {
      if (ranked2$Symbol[i] %in% set1) {
        cum2 <- cum2 + 1
        p <- phyper(cum2 - 1, K2, P - K2, i, lower.tail = FALSE)
        if (p < p2) { p2 <- p; m2 <- i }
      }
    }

    mean_logP <- (-log10(max(p1, 1e-300)) + -log10(max(p2, 1e-300))) / 2
    overlap_n <- length(intersect(set1, set2))

    list(
      score   = mean_logP,
      pval    = 10^(-mean_logP),
      overlap = overlap_n,
      bestM   = round((m1 + m2) / 2)
    )
  }

  # 4 directions
  directions <- list(
    upup     = list(d1 = "up",   d2 = "up",   sign =  1),
    updown   = list(d1 = "up",   d2 = "down", sign = -1),
    downup   = list(d1 = "down", d2 = "up",   sign = -1),
    downdown = list(d1 = "down", d2 = "down", sign =  1)
  )

  result <- list()
  total_score <- 0

  for (name in names(directions)) {
    d <- directions[[name]]
    r <- fisher_bidirectional(d$d1, d$d2)
    raw_score <- -log10(max(r$pval, 1e-300)) * d$sign

    # Clip
    if (is.nan(raw_score) || is.na(raw_score)) raw_score <- 0
    if (is.infinite(raw_score)) raw_score <- sign(raw_score) * clip_limit
    raw_score <- max(-clip_limit, min(raw_score, clip_limit))

    result[[name]] <- list(score = raw_score, pval = r$pval, overlap = r$overlap, bestM = r$bestM)
    total_score <- total_score + raw_score
  }

  total_score <- max(-clip_limit, min(total_score, clip_limit))
  result$total <- total_score

  result
}

# =============================================================================
# FC computation
# =============================================================================

#' Compute fold change for each sample relative to Control mean
#'
#' @param fpkm_mat matrix, genes (rows) x samples (columns)
#' @param control_ids character vector of Control sample IDs
#' @param threshold numeric, FC threshold (default 1.2)
#' @return matrix of signed fold changes (NA for excluded genes)
compute_fold_change <- function(fpkm_mat, control_ids, threshold = THRESHOLD_FC) {
  ctrl_cols <- intersect(control_ids, colnames(fpkm_mat))
  if (length(ctrl_cols) == 0) stop("No Control samples found in expression matrix")

  ctrl_mean <- rowMeans(fpkm_mat[, ctrl_cols, drop = FALSE], na.rm = TRUE)

  fc_mat <- sweep(fpkm_mat, 1, ctrl_mean, "/")

  # Threshold: |FC| < threshold → NA
  fc_mat[fc_mat > 1/threshold & fc_mat < threshold] <- NA

  # FC <= 1 → -1/FC (signed conversion)
  idx <- which(!is.na(fc_mat) & fc_mat <= 1)
  fc_mat[idx] <- -1 / fc_mat[idx]
  fc_mat[is.infinite(fc_mat)] <- NA

  fc_mat
}

# =============================================================================
# Main RFind function
# =============================================================================

#' Run RFind analysis
#'
#' @param fpkm_file path to FPKM/TPM matrix (CSV or Excel)
#' @param demo_file path to sample info file (CSV or Excel, must have ID and Dx columns)
#' @param index_files character vector of paths to reference index files
#' @param threshold_fc FC threshold (default 1.2)
#' @param clip_limit score clipping limit (default 300)
#' @return data.frame with SampleID, Dx, and RFind scores for each index
rfind <- function(fpkm_file, demo_file, index_files,
                  threshold_fc = THRESHOLD_FC, clip_limit = CLIP_LIMIT) {

  cat("=== RFind Analysis ===\n")

  # ── Read FPKM ──
  cat("Reading expression matrix...\n")
  if (grepl("\\.xlsx?$", fpkm_file, ignore.case = TRUE)) {
    fpkm_raw <- as.data.frame(read_excel(fpkm_file))
  } else {
    fpkm_raw <- read.csv(fpkm_file, check.names = FALSE)
  }
  gene_syms <- toupper(fpkm_raw[[1]])
  fpkm_mat  <- as.matrix(fpkm_raw[, -1])
  rownames(fpkm_mat) <- make.unique(gene_syms)
  cat(sprintf("  %d genes x %d samples\n", nrow(fpkm_mat), ncol(fpkm_mat)))

  # ── Read sample info ──
  cat("Reading sample info...\n")
  if (grepl("\\.xlsx?$", demo_file, ignore.case = TRUE)) {
    demo <- as.data.frame(read_excel(demo_file))
  } else {
    demo <- read.csv(demo_file, check.names = FALSE)
  }

  # Find ID and Dx columns
  id_col <- grep("^(ID|SampleID|Sample.?ID|Sample|Subject|BrNum)$",
                  names(demo), ignore.case = TRUE, value = TRUE)[1]
  dx_col <- grep("^(Dx|Diagnosis|Diag|Group|Condition)$",
                  names(demo), ignore.case = TRUE, value = TRUE)[1]
  if (is.na(id_col)) stop("Cannot find ID column in sample info")
  if (is.na(dx_col)) stop("Cannot find Dx column in sample info")

  sample_ids <- as.character(demo[[id_col]])
  dx_values  <- as.character(demo[[dx_col]])

  # Match with FPKM
  matched <- intersect(sample_ids, colnames(fpkm_mat))
  cat(sprintf("  %d samples matched with expression matrix\n", length(matched)))

  control_ids <- sample_ids[dx_values == "Control"]
  control_ids <- intersect(control_ids, colnames(fpkm_mat))
  cat(sprintf("  %d Control samples\n", length(control_ids)))
  if (length(control_ids) == 0) stop("No Control samples found")

  # ── Read indexes ──
  indexes <- list()
  for (idx_file in index_files) {
    cat(sprintf("Reading index: %s\n", basename(idx_file)))
    if (grepl("\\.xlsx?$", idx_file, ignore.case = TRUE)) {
      idx_raw <- as.data.frame(read_excel(idx_file))
    } else {
      idx_raw <- read.csv(idx_file, check.names = FALSE, stringsAsFactors = FALSE)
    }

    # Find Symbol and FC columns
    sym_col <- grep("^(Symbol|Gene|GeneName)$", names(idx_raw), ignore.case = TRUE, value = TRUE)[1]
    fc_col  <- grep("^(Fold.?Change|FC|LogFC|Log2FC)$", names(idx_raw), ignore.case = TRUE, value = TRUE)[1]

    if (is.na(sym_col)) stop(paste("Cannot find Symbol column in", basename(idx_file)))

    idx_df <- data.frame(
      Symbol = toupper(trimws(idx_raw[[sym_col]])),
      stringsAsFactors = FALSE
    )

    if (!is.na(fc_col)) {
      idx_df$FC <- suppressWarnings(as.numeric(idx_raw[[fc_col]]))
    } else {
      # Gene group: all FC = 1
      idx_df$FC <- 1
      cat("  (Gene group mode: no FC column, all FC=1)\n")
    }

    idx_df <- idx_df[!is.na(idx_df$FC) & idx_df$FC != 0, ]

    idx_name <- sub("\\.[^.]+$", "", basename(idx_file))
    cat(sprintf("  %d genes (UP:%d, DOWN:%d)\n",
                nrow(idx_df), sum(idx_df$FC > 0), sum(idx_df$FC < 0)))

    indexes[[idx_name]] <- idx_df
  }

  # ── Compute scores ──
  cat("\nComputing RFind scores...\n")

  all_scores <- data.frame(SampleID = matched, stringsAsFactors = FALSE)
  all_scores$Dx <- dx_values[match(matched, sample_ids)]

  for (idx_name in names(indexes)) {
    cat(sprintf("  Index: %s\n", idx_name))
    idx_df <- indexes[[idx_name]]

    # Common genes
    common_genes <- intersect(rownames(fpkm_mat), idx_df$Symbol)
    P <- length(common_genes)
    cat(sprintf("    Common genes (P): %d\n", P))

    if (P < 10) {
      cat("    WARNING: Too few common genes, skipping\n")
      all_scores[[idx_name]] <- 0
      next
    }

    # Filter index to common genes
    b1 <- idx_df[idx_df$Symbol %in% common_genes, ]

    # Compute FC matrix
    fpkm_sub <- fpkm_mat[common_genes, matched, drop = FALSE]
    fc_mat <- compute_fold_change(fpkm_sub, control_ids, threshold_fc)

    # Score each sample
    scores <- numeric(length(matched))
    for (j in seq_along(matched)) {
      sid <- matched[j]
      fc_vec <- fc_mat[, sid]
      b2 <- data.frame(
        Symbol = names(fc_vec[!is.na(fc_vec)]),
        FC     = fc_vec[!is.na(fc_vec)],
        stringsAsFactors = FALSE
      )

      if (nrow(b2) < 5) {
        scores[j] <- 0
        next
      }

      rf <- running_fisher(b1, b2, P, clip_limit)
      scores[j] <- rf$total
    }

    all_scores[[idx_name]] <- scores
    cat(sprintf("    Done: %d samples scored\n", length(scores)))
  }

  # Add demographic info
  demo_matched <- demo[match(matched, sample_ids), ]
  result <- cbind(all_scores, demo_matched[, !(names(demo_matched) %in% c(id_col, dx_col)), drop = FALSE])

  cat(sprintf("\n=== Complete: %d samples x %d indexes ===\n",
              nrow(result), length(indexes)))

  result
}

# =============================================================================
# Example usage (uncomment to run)
# =============================================================================
# results <- rfind(
#   fpkm_file   = "FPKM_demo.xlsx",
#   demo_file   = "Demographics_demo.xlsx",
#   index_files = c("Dev.csv", "imGC_DEG.csv")
# )
# write.csv(results, "rfind_results.csv", row.names = FALSE)
