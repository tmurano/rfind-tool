# =============================================================================
# Build HGNC gene symbol mapping (alias → official symbol)
# R equivalent of build_hgnc_map.py
#
# Downloads HGNC data and creates hgnc_map.json
# Maps: previous symbols, alias symbols, Ensembl IDs → approved symbol
#
# Usage:
#   source("build_hgnc_map.R")
#   (writes ../hgnc_map.json)
# =============================================================================

library(jsonlite)

HGNC_URL <- paste0(
  "https://www.genenames.org/cgi-bin/download/custom?",
  "col=gd_app_sym&col=gd_prev_sym&col=gd_aliases&col=gd_pub_ensembl_id",
  "&status=Approved&hgnc_dbtag=on&order_by=gd_app_sym_sort&format=text&submit=submit"
)

cat("Downloading HGNC data...\n")
tmp <- tempfile(fileext = ".tsv")
download.file(HGNC_URL, tmp, quiet = TRUE)
hgnc <- read.delim(tmp, stringsAsFactors = FALSE, check.names = FALSE)
cat(sprintf("Downloaded: %d genes\n", nrow(hgnc)))

# Build reverse mapping: alias → set of official symbols
reverse <- list()

add_mapping <- function(key, official) {
  key <- toupper(trimws(key))
  if (nchar(key) == 0 || key == official) return()
  if (is.null(reverse[[key]])) reverse[[key]] <<- character(0)
  reverse[[key]] <<- unique(c(reverse[[key]], official))
}

for (i in seq_len(nrow(hgnc))) {
  official <- toupper(trimws(hgnc$`Approved symbol`[i]))
  if (nchar(official) == 0) next

  # Previous symbols
  prev <- trimws(unlist(strsplit(hgnc$`Previous symbols`[i], ",")))
  for (p in prev) add_mapping(p, official)

  # Alias symbols
  aliases <- trimws(unlist(strsplit(hgnc$`Alias symbols`[i], ",")))
  for (a in aliases) add_mapping(a, official)

  # Ensembl ID
  ensembl <- toupper(trimws(hgnc$`Ensembl gene ID`[i]))
  if (nchar(ensembl) > 0) add_mapping(ensembl, official)
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
output_path <- file.path(dirname(sys.frame(1)$ofile %||% "."), "..", "hgnc_map.json")
output_path <- normalizePath(output_path, mustWork = FALSE)
writeLines(toJSON(mapping, auto_unbox = TRUE), output_path)

cat(sprintf("\nMapping entries: %d\n", length(mapping)))
cat(sprintf("Ambiguous excluded: %d\n", collisions))
cat(sprintf("Output: %s\n", output_path))
