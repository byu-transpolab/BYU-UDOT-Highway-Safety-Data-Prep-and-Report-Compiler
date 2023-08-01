library(tidyverse)
library(sf)
library(openxlsx)
library(ggmap)


# read in files
seg <- read_sf("data/shapefile/CAMS_stats_05Jul23_11_49.shp") %>%
  st_transform(crs = 4326)
int <- read.xlsx("data/output/ISAM_stats_05Jul23_11_49.xlsx") %>%
  st_as_sf(
    coords = c("LONGITUDE", "LATITUDE"), 
    crs = 4326,
    remove = F)


# choose site
site <- seg %>% filter(State_Rank == 1)


# get bounding box
bbox <- st_bbox(site)
bbox <- c(left = bbox[[1]], bottom = bbox[[2]], right = bbox[[3]], top = bbox[[4]])


# get map
map <- get_stamenmap(bbox, maptype = "terrain", zoom = 13)


# display map
ggmap(map) +
geom_sf(data = site %>% st_zm(), 
        aes(fill = YEAR), 
        linewidth = 4, 
        alpha = 0.5, 
        inherit.aes = FALSE) +
theme(axis.title = element_blank(),
      axis.text = element_blank(),
      axis.ticks = element_blank(),
      legend.position = "none")

