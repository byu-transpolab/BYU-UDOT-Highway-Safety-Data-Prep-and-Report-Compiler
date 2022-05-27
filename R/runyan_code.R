# I'm writing my code here instead of directly on the .Rmd because I don't want
# to conflict with Tommy's code.

# INSTALL AND LOAD PACKAGES ################################

# Install pacman ("package manager") if needed
if (!require("pacman")) install.packages("pacman")

# pacman must already be installed; then load contributed
# packages (including pacman) with pacman
pacman::p_load(magrittr, pacman, tidyverse, readxl, sf, zoo, essentials)






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

# Fix that dumb "rurall" problem
df$Urban_Ru_1[df$Urban_Ru_1 == "Rurall"] <- "Rural"

# Group by segment id and search for NAs in AADT or Total_Percent Columns
test <- df
test %>%
  group_by(SEG_ID)%>%
  #fill(AADT:Total_Percent, -Single_Count, -Combo_Count, .direction = "up") %>%
  mutate(
    New_AADT = case_when(
      row_number() == 1L ~ AADT,
      is.na(AADT) & sum(AADT) != rollsum(AADT, k=as.scalar(row_number()), align='left') ~ lead(AADT, n=1L)
      ),
    New_Single_Percent = case_when(
      row_number() == 1L ~ Single_Percent,
      is.na(Single_Percent) & !is.na(slice_tail()) ~ lag(Single_Percent, n=1L)
      ),
    Single_Count = AADT * Single_Percent,
    Combo_Count = AADT * Combo_Percent,
    Total_Count = AADT * Total_Percent
  ) #%>%
  #filter(is.na(AADT))





###############################################
# MEDIANS AND SHOULDER WIDTH ##################
###############################################

# Read in files
df <- read_csv("data/CAMSinput.csv",show_col_types = F)

medians <- read_sf("data/shapefile/medians.shp") %>%
  st_transform(crs = 26912) %>%
  select() %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)

shoulders <- read_sf("data/shapefile/Shoulders.shp") %>%
  st_transform(crs = 26912) %>%
  select() %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)



###############################################
# Fix Spatial Gaps ############################
###############################################

# Use fc file to build a gaps dataset
gaps <- read_filez_csv(fc.filepath, fc.columns)
names(gaps)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
gaps <- gaps %>% filter(grepl("M", ROUTE))
gaps <- gaps %>% filter(grepl("State", RouteType))

gaps <- gaps %>%
  select(ROUTE, BEG_MP, END_MP) %>% 
  arrange(ROUTE, BEG_MP)

gaps$BEG_GAP <- 0
gaps$END_GAP <- 0
for(i in 1:(nrow(gaps)-1)){
  if(gaps$ROUTE[i] == gaps$ROUTE[i+1] & gaps$END_MP[i] < gaps$BEG_MP[i+1]){
    gaps$BEG_GAP[i] <- gaps$END_MP[i]
    gaps$END_GAP[i] <- gaps$BEG_MP[i+1]
  }
}

gaps <- gaps %>%
  select(ROUTE, BEG_GAP, END_GAP) %>%
  filter(BEG_GAP != 0)

# Delete segments which fall within gaps
RC <- RC %>% mutate(MID_MP = BEG_MP + (END_MP - BEG_MP) / 2)

RC$flag <- 0
for(i in 1:nrow(RC)){
  for(j in 1:nrow(gaps)){
    if(RC$MID_MP[i] > gaps$BEG_GAP[j] & 
       RC$MID_MP[i] < gaps$END_GAP[j] &
       RC$BEG_MP[i] >= gaps$BEG_GAP[j] &
       RC$END_MP[i] <= gaps$END_GAP[j] &
       RC$ROUTE[i] == gaps$ROUTE[j]){
      RC$flag[i] <- 1
    }
  }
}

RC <- RC %>% filter(flag == 0) %>% select(-flag, -MID_MP)

# Load exploded CAMS file created in ArcGIS
r_exp.filepath <- "data/shapefile/CAMS_exploded.shp"
r_exp.columns <- c("ROUTE",
                    "BEG_MILEAG",
                    "END_MILEAG",
                    "Match_ID",
                    "Shape_Le_1")
r_exp <- read_filez_shp(r_exp.filepath, r_exp.columns)
names(r_exp)[c(1:5)] <- c("ROUTE", "BEG_MP", "END_MP", "MATCH_ID", "LENGTH")
r_exp <- r_exp %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Convert length to miles
r_exp$LENGTH <- r_exp$LENGTH / 1609.344

# Correct Beg and End milepoints
rt <- ""
for(i in 1:nrow(r_exp)){
  if(r_exp$ROUTE[i] == rt){
    gap_num <- gap_num + 1
    gap <- 0
    gap_chk <- 0
    for(j in 1:nrow(gaps)){
      if(r_exp$ROUTE[i] == gaps$ROUTE[j]){
        gap_chk <- gap_chk + 1
        if(gap_num == gap_chk){
          gap <- gaps$END_GAP[j] - gaps$BEG_GAP[j]
          break
        }
      }
    }
    r_exp$END_MP[i-1] <- r_exp$BEG_MP[i-1] + r_exp$LENGTH[i-1]
    r_exp$BEG_MP[i] <- r_exp$BEG_MP[i-1] + r_exp$LENGTH[i-1] + gap
    r_exp$END_MP[i] <- r_exp$BEG_MP[i] + r_exp$LENGTH[i]
  } else{
    gap_num <- 0
  }
  rt <- r_exp$ROUTE[i]
}

# Write to shapfile for analysis in ArcGIS
st_write(r_exp,"data/shapefile/routes_exploded.shp")

# Read in spatially segmented shapefile from ArcGIS
seg_spatial <- read_sf("data/shapefile/CAMS_spatial.shp")

# Convert crash dataframe to shapefile
crash_sf <- crash %>%
  st_as_sf(
    coords = c("long", "lat"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912) #convert to UTM zone 12

st_write(crash_sf,"data/shapefile/crash.shp")

# Once we are certain of segment IDs we can right join this spatial data into 
# the segments file then we can do a buffer and spatial join with the crashes

# Bad news... the roadway buffer works beautifully, but the crashes don't always
# fall on the current roadway due to roadwork, and there is no way to tell apart
# crashes on an under or overpass. This seemed like a good idea in theory, but 
# there is no way it will work. We have to rely on milepoints. On the bright 
# side, visualizing this in ArcGIS has allowed us to notice details we would have 
# otherwise missed, such as the presence of spatial gaps in the routes.



###############################################
# Determine intersection related crashes ######
###############################################

# Create functional area reference table
FA_ref <- tibble(
  speed = c(5,10,15,20,25,30,35,40,45,50,55,60,65,70,75), 
  d1 = c(75,75,75,75,90,110,130,145,165,185,200,220,240,255,275), 
  d2 = c(70,70,70,70,105,150,225,290,360,440,525,655,755,875,995)
  ) %>%
  mutate(
    d3 = 50,
    total = d1 + d2 + d3
  )

# Create fed routes list
fed <- read_filez_csv(fc.filepath, fc.columns)
names(fed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
fed <- fed %>% filter(grepl("M", ROUTE))
fed <- fed %>% filter(grepl("Fed Aid", RouteType))

fed.routes <- as.character(fed %>% pull(ROUTE) %>% unique() %>% sort())

# Filter intersections
test <- intersection
intersection <- test

intersection <- intersection %>%
  filter(
    SR_SR == "YES" | 
    TRAFFIC_CO == "SIGNAL" | 
    INT_RT_1 %in% substr(fed.routes,1,4) |
    INT_RT_2 %in% substr(fed.routes,1,4) | 
    INT_RT_3 %in% substr(fed.routes,1,4) |
    INT_RT_4 %in% substr(fed.routes,1,4)
  )

# Add speed limit for each intersection approach on intersections file
intersection$INT_RT_1_SL <- NA
intersection$INT_RT_2_SL <- NA
intersection$INT_RT_3_SL <- NA
intersection$INT_RT_4_SL <- NA
for (i in 1:nrow(intersection)){
  int_route1 <- intersection[["INT_RT_1"]][i]
  int_route2 <- intersection[["INT_RT_2"]][i]
  int_route3 <- intersection[["INT_RT_3"]][i]
  int_route4 <- intersection[["INT_RT_4"]][i]
  int_mp <- int[["BEG_MP"]][i]
  speed_row1 <- which(speed$ROUTE == int_route1 & 
                       speed$BEG_MP < int_mp & 
                       speed$END_MP > int_mp)
  speed_row2 <- which(speed$ROUTE == int_route2 & 
                        speed$BEG_MP < int_mp & 
                        speed$END_MP > int_mp)
  speed_row3 <- which(speed$ROUTE == int_route3 & 
                        speed$BEG_MP < int_mp & 
                        speed$END_MP > int_mp)
  speed_row4 <- which(speed$ROUTE == int_route4 & 
                        speed$BEG_MP < int_mp & 
                        speed$END_MP > int_mp)
  intersection[["INT_RT_1_SL"]][i] <- speed$SPEED_LIMIT[speed_row1]
  intersection[["INT_RT_2_SL"]][i] <- speed$SPEED_LIMIT[speed_row2]
  intersection[["INT_RT_3_SL"]][i] <- speed$SPEED_LIMIT[speed_row3]
  intersection[["INT_RT_4_SL"]][i] <- speed$SPEED_LIMIT[speed_row4]
}
