# =============================================================================
# Build Mouse/Rat → Human ortholog mapping
# R equivalent of build_ortholog_map.py
#
# Downloads RGD ortholog data and creates ortholog_map.json
# Only includes entries where uppercased symbols differ between species
#
# Usage:
#   source("build_ortholog_map.R")
#   (writes ../ortholog_map.json)
# =============================================================================

library(jsonlite)

RGD_URL <- "https://download.rgd.mcw.edu/data_release/orthologs/RGD_ORTHOLOGS_Ensembl.txt"

cat("Downloading RGD ortholog data...\n")
tmp <- tempfile(fileext = ".tsv")
download.file(RGD_URL, tmp, quiet = TRUE)

# Read, skipping comment lines
lines <- readLines(tmp)
header_idx <- grep("^RAT_GENE_SYMBOL", lines)[1]
if (is.na(header_idx)) stop("Cannot find header row")

con <- textConnection(lines[header_idx:length(lines)])
ortho <- read.delim(con, stringsAsFactors = FALSE, check.names = FALSE)
close(con)
cat(sprintf("Rows: %d\n", nrow(ortho)))

# Build reverse mapping
reverse <- list()

add_mapping <- function(key, human) {
  key <- toupper(trimws(key))
  human <- toupper(trimws(human))
  if (nchar(key) == 0 || nchar(human) == 0 || key == human) return()
  if (is.null(reverse[[key]])) reverse[[key]] <<- character(0)
  reverse[[key]] <<- unique(c(reverse[[key]], human))
}

for (i in seq_len(nrow(ortho))) {
  human <- ortho$HUMAN_ORTHOLOG_SYMBOL[i]
  if (is.na(human) || nchar(trimws(human)) == 0) next

  # Mouse → Human
  mouse <- ortho$MOUSE_ORTHOLOG_SYMBOL[i]
  if (!is.na(mouse)) add_mapping(mouse, human)

  # Rat → Human
  rat <- ortho$RAT_GENE_SYMBOL[i]
  if (!is.na(rat)) add_mapping(rat, human)
}

# Build final mapping (exclude ambiguous)
mapping <- list()
collisions <- 0
for (key in names(reverse)) {
  officials <- reverse[[key]]
  if (length(officials) == 1) {
    mapping[[key]] <- officials
  } else {
    collisions <- collisions + 1
  }
}

# Write JSON
output_path <- file.path(dirname(sys.frame(1)$ofile %||% "."), "..", "ortholog_map.json")
output_path <- normalizePath(output_path, mustWork = FALSE)
writeLines(toJSON(mapping, auto_unbox = TRUE), output_path)

cat(sprintf("\nMapping entries: %d\n", length(mapping)))
cat(sprintf("Ambiguous excluded: %d\n", collisions))
cat(sprintf("Output: %s\n", output_path))
