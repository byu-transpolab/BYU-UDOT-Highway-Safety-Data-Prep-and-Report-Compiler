library(tidyverse)
library(sf)

aadt <- read_sf("data/shapefile/AADT_Unrounded.shp")
# plot(aadt)
aadt_col <- c("START_ACCU", 
              "END_ACCUM", 
              "ROUTE_NAME", 
              "AADT2020", 
              "SUTRK2020", 
              "CUTRK2020", 
              "AADT2019", 
              "SUTRK2019", 
              "CUTRK2019",
              "AADT2018", 
              "SUTRK2018", 
              "CUTRK2018",
              "AADT2017", 
              "SUTRK2017", 
              "CUTRK2017",
              "AADT2016", 
              "SUTRK2016", 
              "CUTRK2016",  
              "AADT2015", 
              "SUTRK2015", 
              "CUTRK2015",
              "AADT2014",
              "SUTRK2014", 
              "CUTRK2014",
              "Shape_Leng",
              "geometry")

fc <- read_sf("data/shapefile/Functional_Class_ALRS.shp")
# plot(fc)
fc_col <- c("ROUTE_ID",                          
            "FROM_MEASU",                             
            "TO_MEASURE",
            "FUNCTIONAL_CLASS",                             
            "RouteDir",                             
            "RouteType",
            "geometry")

speed <- st_read("data/shapefile/UDOT_Speed_Limits_2019.shp")
# plot(speed)
speed_col <- c("ROUTE_ID",                         
               "FROM_MEASU",                         
               "TO_MEASURE",                         
               "SPEED_LIMI",
               "geometry")

lane <- read_sf("data/shapefile/Lanes.shp")
# plot(lane)

urban_small <- read_sf("data/shapefile/Urban_Boundaries_Small.shp")
# plot(urban_small)
small_col <- c("ROUTE_ID",                         
               "FROM_MEASURE",                         
               "TO_MEASURE",                         
               "SPEED_LIMIT",
               "geometry")

urban_large <- read_sf("data/shapefile/Urban_Boundaries_Large.shp")
# plot(urban_large)
large_col <- c("OBJECTID",                         
               "TYPE_",
               "geometry")

intersection <- read_sf("data/shapefile/Intersections.shp")
# plot(intersection)

pm <- read_sf("data/shapefile/Pavement_Messages.shp")
# plot(pm)

driveway <- read_sf("data/shapefile/Driveway.shp")
# plot(driveway)

median <- read_sf("data/shapefile/Medians.shp")
# plot(median)

shoulder <- read_sf("data/shapefile/Shoulders.shp")
# plot(shoulder)

st_join(aadt, fc)

segment_breaks <- function(sdtm) {
  breaks_df <- sdtm %>% 
    group_by(ROUTE) %>% 
    summarize(startpoints = unique(c(START_ACCUM, END_ACCUM))) %>% 
    arrange(startpoints)
  return(breaks_df)
}
