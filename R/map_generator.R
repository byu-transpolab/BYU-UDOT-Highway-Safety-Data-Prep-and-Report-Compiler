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
sites <- seg %>% filter(Region_Ran >= 1 & Region_Ran <= 10)


# indicate segment or intersection
site_type <- "Seg"


# loop through sites
for(i in 1:nrow(sites)){
  site <- sites[i,]

  
  # get bounding box
  bbox <- st_bbox(site)
  bbox <- c(left = bbox[[1]]-0.01, bottom = bbox[[2]]-0.01, right = bbox[[3]]+0.01, top = bbox[[4]]+0.01)
  
  
  # get map
  map <- get_stamenmap(bbox, maptype = "terrain", zoom = calc_zoom(bbox))
  
  
  # display map
  ggmap(map) +
    geom_sf(data = site %>% st_zm(), 
            aes(fill = YEAR), 
            linewidth = 4, 
            alpha = 0.5, 
            color = "#DC4C4C",
            inherit.aes = FALSE) +
    theme(axis.title = element_blank(),
          axis.text = element_blank(),
          axis.ticks = element_blank(),
          legend.position = "none")
  
  
  # get save location
  name <- paste0(site_type,"-StRank",site$State_Rank[1],"-Region",site$REGION[1])
  path <- paste0("data/output/maps/",name,".png")
  
  
  # save map
  ggsave(path, width = 6.5, height = 4, units = "in")
}




