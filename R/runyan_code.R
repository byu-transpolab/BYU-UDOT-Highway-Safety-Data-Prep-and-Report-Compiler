# I'm writing my code here instead of directly on the .Rmd because I don't want
# to conflict with Tommy's code.

# INSTALL AND LOAD PACKAGES ################################

# Install pacman ("package manager") if needed
if (!require("pacman")) install.packages("pacman")

# pacman must already be installed; then load contributed
# packages (including pacman) with pacman
pacman::p_load(magrittr, pacman, tidyverse, readxl, sf, zoo)






############################################################
# FILTER CAMS CRASH DATA ###################################
############################################################

# change route 089A to 0011... There seems to be more to this though
crash$ROUTE_ID[crash$ROUTE_ID == '089A'] <- 0011

# change route ID to 194 from 0085 for milepoint less than 2.977 
crash$ROUTE_ID[crash$MILEPOINT < 2.977 & crash$ROUTE_ID == 0085] <- 0194

# eliminate rows with ramps
crash %>%
  as_tibble() %>%
  filter(RAMP_ID == 0 & is.na(RAMP_ID))
  


############################################################
# Combine Additional Variables to Final Output #############
############################################################
# SPATIAL ANALYSIS

#LOAD DATA

crashes <- read_csv("data/CAMS_Crashes_14-21.csv",show_col_types = F)
loc <- read_csv("data/Crash_Data_14-20.csv",show_col_types = F) %>%
  select(crash_id,lat,long)
crashes <- left_join(crashes,loc, by = c("CRASH_ID" = "crash_id")) %>%
  st_as_sf(
    coords = c("long", "lat"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912) #convert to UTM zone 12

intersections <- read_csv("data/Intersection_w_ID_MEV_MP.csv") %>%
  st_as_sf(
    coords = c("BEG_LONG", "BEG_LAT"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912)

UTA_stops <- read_sf("data/UTA_Stops/UTA_Stops_and_Most_Recent_Ridership.shp") %>%
  st_transform(crs = 26912) %>%
  select(UTA_StopID, Mode) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)

bus_stops <- UTA_stops %>%
  filter(Mode == "Bus") %>%
  select(-Mode)

rail_stops <- UTA_stops %>%
  filter(Mode == "Rail") %>%
  select(-Mode)

schools <- read_sf("data/Utah_Schools_PreK_to_12/Utah_Schools_PreK_to_12.shp") %>%
  st_transform(crs = 26912) %>%
  select(SchoolID, SchoolLeve, OnlineScho, SchoolType, TotalK12) %>%
  filter(is.na(OnlineScho), SchoolType == "Vocational" | SchoolType == "Special Education" | SchoolType == "Residential Treatment" | SchoolType == "Regular Education" | SchoolType == "Alternative") %>%
  select(-OnlineScho, -SchoolType, -TotalK12) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)

high_schools <- schools %>%
  filter(SchoolLeve == "HIGH")

mid_schools <- schools %>%
  filter(SchoolLeve == "MID")

elem_schools <- schools %>%
  filter(SchoolLeve == "ELEM")

prek_schools <- schools %>%
  filter(SchoolLeve == "PREK")

ktwelve_schools <- schools %>%
  filter(SchoolLeve == "K12")

# VISUALIZE

# visualize crashes
rm("joinprep","location","vehicle","rollups")

df <- crash %>%
  select(crash_id,lat,long,crash_severity_id,weather_condition_id,manner_collision_id,pedalcycle_involved) %>%
  filter(lat > 30 & long < -95 & !is.na(lat) & !is.na(long))

crash_sf <- st_as_sf(df, 
                     coords = c("long", "lat"), 
                     crs = 4326,
                     remove = F)

crash_sf <- st_transform(crash_sf, crs = 26912) #convert to UTM zone 12

# determine which intersections are 1000 ft from a school
ints_near_schools <- st_join(intersections, schools, join = st_within) %>%
  filter(!is.na(SchoolID))

# determine which intersections are 1000 ft from a UTA stop
ints_near_UTA <- st_join(intersections, UTA_stops, join = st_within) %>%
  filter(!is.na(UTA_StopID))

plot(ints_near_schools["Int_ID"])
plot(ints_near_UTA["Int_ID"])
plot(intersections["Int_ID"])

# COMBINE

# modify intersections object
intersections <- st_join(intersections, schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()

intersections <- st_join(intersections, prek_schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_PREK_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()

intersections <- st_join(intersections, elem_schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_ELEM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()

intersections <- st_join(intersections, mid_schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_MID_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()

intersections <- st_join(intersections, high_schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_HIGH_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()

intersections <- st_join(intersections, ktwelve_schools, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_K12_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID, -SchoolLeve) %>%
  unique()



intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_UTA_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
  select(-UTA_StopID, -Mode) %>%
  unique()

intersections <- st_join(intersections, bus_stops, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_BUS_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
  select(-UTA_StopID) %>%
  unique()

intersections <- st_join(intersections, rail_stops, join = st_within) %>%
  group_by(Int_ID) %>%
  mutate(NUM_RAIL_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
  select(-UTA_StopID) %>%
  unique()

intersections <- unique(intersections)
st_drop_geometry(intersections)
 
# save csv file
write_csv(intersections, file = "data/Intersections_Compiled.csv")


# modify crash data frame
crashes <- st_join(crashes, schools, join = st_within) %>%
  group_by(CRASH_ID) %>%
  mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  select(-SchoolID)

crashes <- st_join(crashes, UTA_stops, join = st_within) %>%
  group_by(CRASH_ID) %>%
  mutate(NUM_UTA = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
  select(-UTA_StopID)

crashes <- unique(crashes)
st_drop_geometry(crashes)

# save csv file
write_csv(crashes, file = "data/Crashes_Compiled_14-21.csv")

# remove old dfs
rm("crash_sf","df","schools", "high_schools", "mid_schools", "elem_schools", "prek_schools", "ktwelve_schools", "UTA_stops","ints_near_schools","ints_near_UTA", "bus_stops", "rail_stops")






# PATCH THE EXCEL OUTPUT
# unless we can move everything to R, this will at least allow us to fix the 
# excel output file so it gives us more variables

# Load excel output file for UICPM
df <- read_csv("data/UICPMinput.csv") #%>%
  #st_as_sf(
  #  coords = c("LONGITUDE", "LATITUDE"), 
  #  crs = 4326,
  #  remove = F) %>%
  #st_transform(crs = 26912)

# Load pavement messages
cw <- read_csv("data/Pavement_Messages.csv") #%>%
  #st_as_sf(
  #  coords = c("BEG_LONG", "BEG_LAT"), 
  #  crs = 4326,
  #  remove = F) %>%
  #st_transform(crs = 26912) %>%
  #st_buffer(dist = 20) %>%       #buffer 20 meters to cover intersection
  select(MILEPOINT = START_ACCUM, TYPE) %>%
  filter(grepl("CROSSWALK",TYPE)) %>%
  mutate(crosswalk = 1) %>%
  select(-TYPE)

# add presence of crosswalks
df <- left_join(df, cw, by = c("UDOT_BMP" = "MILEPOINT"))

#df <- st_join(df, cw, join = st_within) %>%
#  group_by(INT_ID) %>%
#  mutate(CROSSWALK_PRESENT = sum(crosswalk)) %>%
#  select(-crosswalk, -MILEPOINT, -geometry) %>%
#  unique()
#st_drop_geometry(df)





###############################################
# FIX MISSING AADT AND TRUCK DATA #############
###############################################

# Read in file
df <- read_csv("data/CAMSinput.csv",show_col_types = F)

# Convert zeros to NA
df$AADT[df$AADT == 0] <- NA
df$Single_Percent[df$Single_Percent == 0] <- NA
df$Combo_Percent[df$Combo_Percent == 0] <- NA
df$Single_Count[df$Single_Count == 0] <- NA
df$Combo_Count[df$Combo_Count == 0] <- NA
df$Total_Percent[df$Total_Percent == 0] <- NA
df$Total_Count[df$Total_Count == 0] <- NA

# Group by segment id and search for NAs in AADT or Total_Percent Columns
df %>%
  group_by(SEG_ID)%>%
  fill(AADT:Total_Percent, -Single_Count, -Combo_Count, .direction = "up") %>%
  mutate(
    Single_Count = AADT * Single_Percent,
    Combo_Count = AADT * Combo_Percent,
    Total_Count = AADT * Total_Percent
  ) #%>%
  #filter(is.na(AADT))
