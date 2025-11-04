library(readxl)
library(readr)
library(magrittr)
library(dplyr)
library(tidygeocoder)
library(sf)
library(ggplot2)
library(luaRlp)


reGeoCode <- FALSE

files <- sapply(c(
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_31/",
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_32/",
  "O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/LIMS_2_RIDOM/Referat_33/"
), list.files, pattern = "*.xls", full.names = TRUE) |>
  unlist(use.names = FALSE)

xl_list <- lapply(files, read_excel)

common_colnames <- Reduce(intersect, lapply(xl_list, colnames))

LIMS <- do.call(rbind, lapply(xl_list, function (x) x[, common_colnames]))

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

table(LIMS$Accuracy, cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      useNA="ifany")

table(LIMS$Ergebnis...22,
      cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      useNA="ifany")

LIMS$Latex <- ifelse(grepl("[S|s]pezies|spec", LIMS$Ergebnis...22), "Spp",
                     ifelse(grepl("[S|s]erogruppe *1", LIMS$Ergebnis...22), "Ser 1",
                            ifelse(grepl("[S|s]erogrup+e *2", LIMS$Ergebnis...22), "Ser 2-14", NA)))


table(LIMS$Ergebnis...22, LIMS$Latex, useNA="ifany")


table(LIMS$Latex, cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      useNA="ifany")

table(LIMS$Latex, cut(LIMS$KolNum, breaks = c(0, 1, 10, 100, 1000, Inf)),
      LIMS$Accuracy, useNA="ifany")

LIMS$Gemeinde <- gsub("Bodenheim-Roxheim", "Bobenheim-Roxheim", LIMS$Gemeinde)

LIMS$Straße...9 <- gsub("Folzring", "Foltzring", LIMS$Straße...9)

LIMS$Straße...9 <- gsub("Stiftsstr", "Stiftstr", LIMS$Straße...9)

LIMS$Straße...9 <- gsub("Marianstr", "Marienstr", LIMS$Straße...9)

LIMS$Gemeinde <- gsub("Klingemünster", "Klingenmünster", LIMS$Gemeinde)

LIMS$Straße...9 <- gsub("Andreas-Holzamer-Weg 32", "Andreas-Holzamer-Ring 32",
                        LIMS$Straße...9)

LIMS$Straße...9 <- gsub("Pfarrer-Johannes-Wilhelmweilstr. 4",
                        "Pfarrer-Joh.-W.-Weil-Straße 4",
                        LIMS$Straße...9)


addresses <- paste0(LIMS$Straße...9, ", ", LIMS$Gemeinde, ", Rheinland-Pfalz, Deutschland") |>
  tibble(singlelineaddress = _)


if(reGeoCode){
  geocoded <- geo(address = addresses$singlelineaddress, method = "osm",
                  lat = latitude, long = longitude
  )
  nacodes <- geocoded[is.na(geocoded$latitude), ]

  ## patch the ones osm doesn't find with arcgis
  geocodedArc <- geo(address = nacodes$address, method = "arcgis",
                     lat = latitude, long = longitude)
  geocoded[is.na(geocoded$latitude), c("latitude", "longitude")] <-
    geocodedArc[, c("latitude", "longitude")]

  ## Make sure they are all in RLP
  geocoded_sf <- st_as_sf(
    geocoded,
    coords = c("longitude", "latitude"),
    crs = 4326  # WGS84
  )
  RLP_sf <- st_transform(RLP_geo$Land, st_crs(geocoded_sf))
  is_inside <- st_within(geocoded_sf, RLP_sf, sparse = FALSE)
  geocoded$latitude[!is_inside] <- NA
  geocoded$longitude[!is_inside] <- NA
  saveRDS(geocoded, "geocoded.RDS")
} else{
  geocoded <- readRDS("geocoded.RDS")
}

unique(geocoded[is.na(geocoded$latitude), ])
### Most of these addresses seem to be outside of RLP

RIDOM_pneumophila <- read_excel("O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/RIDOM_produktiv_projekt/Produktiv_Legionella_pneumophila.xlsx")

table(RIDOM_pneumophila$`Sample ID`%in%LIMS$Labornummer)

RIDOM_pneumophila$Sample_ID_LIMS <- gsub("(\\d{4}-20\\d\\d-\\d{6}).*", "\\1",
                                         RIDOM_pneumophila$`Sample ID`)

table(RIDOM_pneumophila$`Sample ID`%in%LIMS$Labornummer)


RIDOM_anisa <- read_excel("O:/Abteilung Humanmedizin (AHM)/Referat 32/32_6/Legionella Project/RIDOM_produktiv_projekt/Produktiv_Legionella_anisa.xlsx")

RIDOM_anisa$Sample_ID_LIMS <- gsub("(\\d{4}-20\\d\\d-\\d{6}).*", "\\1",
                                         RIDOM_anisa$`Sample ID`)


LIMS$sequenced <- ifelse(LIMS$Labornummer%in%RIDOM_pneumophila$Sample_ID_LIMS,
                         "L. pneumophila",
                         ifelse(LIMS$Labornummer%in%RIDOM_anisa$Sample_ID_LIMS,
                                "L. anisa", "no sequence"))

LIMS$`Probenahme  (Datum)` <- as.Date(LIMS$`Probenahme  (Datum)`)


# Extract time (HH:MM:SS) from "Probenahme (Zeit)"
LIMS$`Probenahme  (Zeit)` <- strftime(
  as.POSIXct(LIMS$`Probenahme  (Zeit)`),
  format = "%H:%M:%S"
)

OUT <- cbind(LIMS, geocoded,
             "Lat/Long of Isolation" = paste(geocoded$latitude, geocoded$longitude, sep=", "))

OUT_sf <- OUT |>
  filter(!is.na(latitude), !is.na(longitude)) %>%
  st_as_sf(coords = c("longitude", "latitude"), crs = 4326)

ggplot() +
  geom_sf(data = RLP_geo$RBZ, fill = "white", color = "grey40") +
  geom_sf(data = OUT_sf[is.na(OUT_sf$Latex),], alpha=0.3,
          aes(shape = Ref, size = log10(KolNum+1))) +
  geom_sf(data = OUT_sf[!is.na(OUT_sf$Latex),], alpha=0.3,
          aes(shape = Ref, color = Latex, size = log10(KolNum+1))) +
  geom_sf(data = OUT_sf[!OUT_sf$sequenced%in%"no sequence",],
           size = 1, aes(color = sequenced)) +
  theme_minimal() +
  ggtitle("Wasser-Legionellen in Kreise (RLP)")

ggsave("figures/LegionellenLD_Kreise.png", width=8, height=8, bg="white")



write_csv(OUT[OUT$sequenced%in%"L. pneumophila", ], "Imp_RIDOM_pneumophila.csv")
