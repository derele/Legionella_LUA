### benötigte Pakete laden
library(readxl)
library(dplyr)
library(stringr)
library(stringi)
library(stringdist)
library(igraph)


## ------------------------------------------------------------
## Input: Lesen der Excel files aus dem LIMS („LDM_Abfragesicht_9“)
## ------------------------------------------------------------

files <- sapply(c(
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_31/",
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_32/",
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_33/"
), list.files, pattern = ".*2025*.xls", full.names = TRUE) |>
  unlist(use.names = FALSE)

xl_list <- lapply(files, read_excel)
common_colnames <- Reduce(intersect, lapply(xl_list, colnames))
LIMS <- do.call(rbind, lapply(xl_list, function (x) x[, common_colnames]))



## ============================================================
## Adressbereinigung und Deduplikation (deterministisch + fuzzy)
##
## Schritt 1:
##   Deterministische Textnormalisierung
##   (z.B. Vereinheitlichung von Straße/Str./straße, Entfernung
##   von Sonderzeichen, Korrektur häufiger Tippfehler,
##   Normalisierung von Hausnummern-Bereichen)
##
## Schritt 2a:
##   Schlüsselbasiertes Zusammenführen innerhalb gleicher PLZ
##   und gleichen Ortsnamens.
##   Zusätzliche Informationen nach der Hausnummer werden entfernt
##   (z.B. Einrichtungs- oder Gebäudenamen).
##
## Schritt 2b:
##   Fuzzy-Zusammenführung von Straßennamen innerhalb gleicher
##   PLZ, gleichem Ort und identischer Hausnummer.
##   Die Hausnummer (inkl. Bereich) wird dabei exakt behandelt,
##   nur der Straßenname wird fehlertolerant verglichen.
##
## Ergebnis:
##   Kanonische, vereinheitlichte Adresse zur weiteren Auswertung
##   in der Variablen:
##
##     LIMS2$addr_final
##
## ============================================================



## ------------------------------------------------------------
## Deterministic text normalization for address components
## ------------------------------------------------------------

norm_txt <- function(x) {
  x <- ifelse(is.na(x), "", x)
  x <- str_squish(x)
  x <- str_to_lower(x)

  ## remove literal "na"
  x <- str_replace_all(x, "\\bna\\b", "")

  ## transliterate (umlauts etc.)
  x <- stringi::stri_trans_general(x, "de-ASCII")
  x <- str_replace_all(x, "ß", "ss")

  ## unify punctuation/whitespace
  x <- str_replace_all(x, "[,;]", " ")
  x <- str_replace_all(x, "\\s+", " ")
  x <- str_trim(x)

  ## normalize street tokens
  x <- str_replace_all(x, "\\bstraße\\b", "strasse")
  x <- str_replace_all(x, "\\bstr\\.?\\b", "strasse")

  ## fix "steinfelderstrasse" / "steinfelderstr." style glue
  x <- str_replace_all(x, "([[:alpha:]])strasse\\b", "\\1 strasse")
  x <- str_replace_all(x, "([[:alpha:]])str\\.?\\b", "\\1 strasse")

  ## handle frequent typos / glued tokens (deterministic)
  ## - "waisenhausstrr" -> "waisenhaus strasse"
  x <- str_replace_all(x, "strr\\b", "strasse")

  ## hyphens to space (e.g. "ernestus-platz" -> "ernestus platz")
  x <- str_replace_all(x, "-", " ")

  ## split glued "platz"
  x <- str_replace_all(x, "([[:alpha:]])platz\\b", "\\1 platz")

  ## split glued "allee" and normalize variants like "nordallee1"
  x <- str_replace_all(x, "([[:alpha:]])allee\\b", "\\1 allee")
  x <- str_replace_all(x, "\\ballee(?=\\d)", "allee ")

  ## remove trailing dots before whitespace/end (e.g. "strasse." -> "strasse")
  x <- str_replace_all(x, "\\.(?=\\s|$)", "")

  ## normalize house number ranges: "20+22" / "20 + 22" -> "20-22"
  x <- str_replace_all(x, "(\\d)\\s*\\+\\s*(\\d)", "\\1-\\2")
  x <- str_replace_all(x, "(\\d)\\s*-\\s*(\\d)", "\\1-\\2")

  ## fix "strasse44" -> "strasse 44"
  x <- str_replace_all(x, "\\bstrasse(?=\\d)", "strasse ")

  str_squish(x)
}

pick_city <- function(gemeinde, ort) {
  g <- norm_txt(gemeinde)
  o <- norm_txt(ort)
  ifelse(o == "" | o == g, g, o)
}

make_short_address <- function(street, plz, city) {
  s <- norm_txt(street)
  p <- str_extract(norm_txt(plz), "\\b\\d{5}\\b")
  p <- ifelse(is.na(p), "", p)
  c <- norm_txt(city)

  out <- str_squish(paste0(s, ", ", p, " ", c))
  out <- str_replace_all(out, "^,\\s*", "")
  out <- str_replace_all(out, ",\\s*$", "")
  out
}

## ------------------------------------------------------------
## Step 0: build LIMS2 (fixes "short_addr not found" error)
## ------------------------------------------------------------

LIMS2 <- LIMS %>%
  mutate(
    city = pick_city(Gemeinde, Ort),
    short_addr = make_short_address(`Straße...9`, PLZ, city)
  )

## ------------------------------------------------------------
## Step 2a: Key merge within PLZ+Ort (remove extra info)
## ------------------------------------------------------------

extract_street_hnr <- function(short_addr_norm) {
  pre <- str_trim(str_extract(short_addr_norm, "^[^,]+"))

  ## keep only up to the first house number (optional letter and optional range)
  street_hnr <- str_trim(str_extract(
    pre,
    "^[[:alpha:] .]+\\s*\\d+[[:alpha:]]?(?:\\s*-\\s*\\d+[[:alpha:]]?)?"
  ))

  ifelse(is.na(street_hnr) | street_hnr == "", pre, street_hnr)
}

LIMS2 <- LIMS2 %>%
  mutate(
    short_addr_norm = norm_txt(short_addr),
    plz_norm  = str_extract(short_addr_norm, "\\b\\d{5}\\b"),
    plz_norm  = ifelse(is.na(plz_norm), "", plz_norm),
    city_norm = norm_txt(city),
    street_hnr = extract_street_hnr(short_addr_norm),

    addr_key = str_squish(paste(plz_norm, city_norm, street_hnr))
  ) %>%
  group_by(plz_norm, city_norm, addr_key) %>%
  mutate(
    short_addr_keymerge = short_addr[which.min(nchar(short_addr))]
  ) %>%
  ungroup()

## ------------------------------------------------------------
## Step 2b: Fuzzy merge street name only, house number exact
## ------------------------------------------------------------

split_street_hnr <- function(street_hnr) {
  x <- norm_txt(street_hnr)

  ## house number token at end: 12, 12a, 12-14, 12a-14b
  hnr <- str_extract(x, "\\b\\d+[[:alpha:]]?(?:-\\d+[[:alpha:]]?)?\\b\\s*$")
  hnr <- str_trim(ifelse(is.na(hnr), "", hnr))

  street <- str_trim(str_replace(x, "\\b\\d+[[:alpha:]]?(?:-\\d+[[:alpha:]]?)?\\b\\s*$", ""))
  street <- str_squish(street)

  tibble(street_name = street, hnr_exact = hnr)
}

cluster_streetnames <- function(street_names, method = "jw", max_dist = 0.10) {
  u <- unique(street_names)
  u <- u[!is.na(u) & u != ""]
  if (length(u) <= 1) return(rep(1L, length(street_names)))

  D <- stringdist::stringdistmatrix(u, u, method = method)
  adj <- (D <= max_dist) & (row(D) != col(D))

  g <- igraph::graph_from_adjacency_matrix(adj, mode = "undirected", diag = FALSE)
  memb <- igraph::components(g)$membership
  names(memb) <- u

  as.integer(memb[street_names])
}

tmp <- split_street_hnr(LIMS2$street_hnr)

LIMS2 <- LIMS2 %>%
  bind_cols(tmp) %>%
  group_by(plz_norm, city_norm, hnr_exact) %>%   ## house number exact
  mutate(
    street_cluster = cluster_streetnames(street_name, method = "jw", max_dist = 0.10)
  ) %>%
  group_by(plz_norm, city_norm, hnr_exact, street_cluster) %>%
  mutate(
    street_name_canon = names(which.max(table(street_name))),
    addr_final = str_squish(paste0(street_name_canon, " ", hnr_exact, ", ", plz_norm, " ", city_norm))
  ) %>%
  ungroup()

## ------------------------------------------------------------
## Diagnostics (optional)
## ------------------------------------------------------------

## show clusters with multiple spellings
multi <- LIMS2 %>%
  count(plz_norm, city_norm, hnr_exact, street_cluster, street_name, sort = TRUE) %>%
  group_by(plz_norm, city_norm, hnr_exact, street_cluster) %>%
  filter(n() > 1)

## table of final addresses
table(LIMS2$addr_final)

## ------------------------------------------------------------
## Final address variable for downstream work:
##   LIMS2$addr_final
## ------------------------------------------------------------

LIMS$Zweck <- gsub(".*(Zweck \\w).*", "\\1", LIMS$`Probenahme SOP`)

table(LIMS$Zweck)

LIMS$Zweck[LIMS$Zweck%in%"Probenahme gemäß SOP Q EX.T 0002 05 (Schöpfprobe nach DIN EN ISO 19458:2006-12)"] <- "Schöpfprobe"

LIMS$Ref <- substr(LIMS$Labornummer, 1, 2)

LIMS$Accuracy <- ifelse(grepl(">", LIMS$Ergebnis...19),
                        "gr", ifelse(grepl("<", LIMS$Ergebnis...19),
                                     "kl",
                                     ifelse(!is.na(as.numeric(LIMS$Ergebnis...19)),
                                            "exa", NA)))

LIMS$KolNum <- as.numeric(gsub(">|<| *", "", LIMS$Ergebnis...19))

## Okay, NAs unvermeidbar
LIMS$Ergebnis...19[is.na(LIMS$KolNum)]


table(LIMS$Accuracy, cut(LIMS$KolNum, breaks = c(0, 2, 10, 100, 1000, Inf)),
      useNA="ifany")

table(LIMS$Ergebnis...22,
      cut(LIMS$KolNum, breaks = c(0, 2, 10, 100, 1000, Inf)),
      useNA="ifany")

LIMS$Latex <- ifelse(grepl("[S|s]pezies|spec", LIMS$Ergebnis...22), "Spp",
                     ifelse(grepl("[S|s]erogruppe *1", LIMS$Ergebnis...22), "Ser 1",
                            ifelse(grepl("[S|s]erogrup+e *2", LIMS$Ergebnis...22), "Ser 2-14", NA)))


table(LIMS$Ergebnis...22, LIMS$Latex, useNA="ifany")


table(LIMS$Latex, cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      useNA="ifany")

table(LIMS$Latex, cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      LIMS$Accuracy, useNA="ifany")

