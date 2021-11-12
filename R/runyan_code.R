# I'm writing my code here instead of directly on the .Rmd because I don't want
# to conflict with Tommy's code.

# INSTALL AND LOAD PACKAGES ################################

# Install pacman ("package manager") if needed
if (!require("pacman")) install.packages("pacman")

# pacman must already be installed; then load contributed
# packages (including pacman) with pacman
pacman::p_load(magrittr, pacman, tidyverse, readxl, sf)


# FILTER CAMS CRASH DATA ###################################

# change route 089A to 0011... There seems to be more to this though
crash$ROUTE_ID[crash$ROUTE_ID == '089A'] <- 0011

# change route ID to 194 from 0085 for milepoint less than 2.977 
crash$ROUTE_ID[crash$MILEPOINT < 2.977 & crash$ROUTE_ID == 0085] <- 0194

# eliminate rows with ramps
crash %>%
  as_tibble() %>%
  filter(RAMP_ID == 0 & is.na(RAMP_ID))
  



# Combine Additional Variables to Final Output #############

# SPATIAL ANALYSIS

rm("joinprep","location","vehicle","rollups")

df <- crash %>%
  select(crash_id,lat,long,crash_severity_id,weather_condition_id,manner_collision_id,pedalcycle_involved) %>%
  filter(lat > 30 & long < -95 & !is.na(lat) & !is.na(long))

crash_sf <- st_as_sf(df, 
               coords = c("long", "lat"), 
               crs = 4326,
               remove = F)

crash_sf <- st_transform(crash_sf, crs = 26912) #convert to UTM zone 12

intersections <- read_sf("data/Intersections/Intersections.shp") %>%
  st_transform(crs = 26912)

UTA_stops <- read_sf("data/UTA_Stops/UTA_Stops_and_Most_Recent_Ridership.shp") %>%
  st_transform(crs = 26912) %>%
  select(UTA_StopID) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)

schools <- read_sf("data/Utah_Schools_PreK_to_12/Utah_Schools_PreK_to_12.shp") %>%
  st_transform(crs = 26912) %>%
  select(SchoolID) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)

# VISUALIZE

# determine which intersections are 1000 ft from a school
ints_near_schools <- st_join(intersections, schools, join = st_within) %>%
  filter(!is.na(SchoolID))

# detemine which intersections are 1000 ft from a UTA stop
ints_near_UTA <- st_join(intersections, UTA_stops, join = st_within) %>%
  filter(!is.na(UTA_StopID))

plot(ints_near_schools["ID"])
plot(ints_near_UTA["ID"])
plot(intersections["ID"])

# COMBINE

# modify intersections object
intersections <- st_join(intersections, schools, join = st_within)
intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
  st_drop_geometry()

# remove old dfs
rm("crash_sf","df","schools","UTA_stops","ints_near_schools","ints_near_UTA")

# additional intersection data
#ints_extra <- read_csv("data/Intersection_w_ID_MEV_MP.csv")
#intersections <- full_join(intersections, ints_extra)
 
# save csv file
write_csv(intersections, file = "data/Intersections_Compiled.csv")
