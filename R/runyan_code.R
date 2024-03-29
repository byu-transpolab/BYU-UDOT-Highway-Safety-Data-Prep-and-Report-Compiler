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
# Modify Intersections ########################
###############################################

# Modify intersections to include bus stops and schools near intersections
mod_intersections <- function(intersections,UTA_Stops,schools){
  intersections <- st_join(intersections, schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID)

  intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_UTA = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
    select(-UTA_StopID)

  intersections <- unique(intersections)
  st_drop_geometry(intersections)
}



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



###############################################
# Spatially join segments with crashes ########
###############################################

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

# Write to shapefile for analysis in ArcGIS
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
  speed = c(5,10,15,20,25,30,35,40,45,50,55,60,65,70,75,80), 
  d1 = c(75,75,75,75,90,110,130,145,165,185,200,220,240,255,275,275), 
  d2 = c(70,70,70,70,105,150,225,290,360,440,525,655,755,875,995,995)
  ) %>%
  mutate(
    d3 = 50,
    total = d1 + d2 + d3
  )

# Create fed routes list
fed <- read_csv_file(fc_fp, fc_col)
names(fed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
fed <- fed %>% filter(grepl("M", ROUTE))
fed <- fed %>% filter(grepl("Fed Aid", RouteType))

fed.routes <- as.character(fed %>% pull(ROUTE) %>% unique() %>% sort())

# Read and filter intersections (SR_SR or SR_FedAid or Signalized SR)
intersection_fp <- "data/csv/Intersections_Compiled.csv"
intersection_col <- c("ROUTE",
                          "UDOT_BMP",
                          "UDOT_EMP",
                          "Int_ID",
                          "INT_TYPE",
                          "TRAFFIC_CO",
                          "SR_SR",
                          "INT_RT_1",
                          "INT_RT_2",
                          "INT_RT_3",
                          "INT_RT_4",
                          "INT_RT_1_M",
                          "INT_RT_2_M",
                          "INT_RT_3_M",
                          "INT_RT_4_M",
                          "STATION",
                          "REGION",
                          "BEG_LONG",
                          "BEG_LAT",
                          "BEG_ELEV")

intersection <- read_csv_file(intersection_fp, intersection_col)
names(intersection)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
# Add M to Route Column to Standardize Route Format
intersection$ROUTE <- paste(substr(intersection$ROUTE, 1, 5), "M", sep = "")
# Getting only state routes
intersection <- intersection %>% filter(ROUTE %in% substr(main.routes, 1, 6))
# Filter SR_SR, Fed_Aid, and Signalized
intersection <- intersection %>%
  filter(
    SR_SR == "YES" | 
    TRAFFIC_CO == "SIGNAL" | 
    INT_RT_1 %in% substr(fed.routes,1,4) |
    INT_RT_2 %in% substr(fed.routes,1,4) | 
    INT_RT_3 %in% substr(fed.routes,1,4) |
    INT_RT_4 %in% substr(fed.routes,1,4)
  )
# Sort
intersection <- intersection %>% arrange(ROUTE, BEG_MP)

# Create full speed file for merging with intersections
speed_full <- read_csv_file(speed_fp, speed_col)
names(speed_full)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
speed_full <- speed_full %>% filter(grepl("M", ROUTE))
speed_full <- compress_seg(speed_full)

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
  int_mp_1 <- intersection[["INT_RT_1_M"]][i]
  int_mp_2 <- intersection[["INT_RT_2_M"]][i]
  int_mp_3 <- intersection[["INT_RT_3_M"]][i]
  int_mp_4 <- intersection[["INT_RT_4_M"]][i]
  speed_row1 <- NA
  speed_row2 <- NA
  speed_row3 <- NA
  speed_row4 <- NA
  # get route direction
  suffix = substr(intersection$ROUTE[i],5,6)
  # find rows in speed limit file
  if(!is.na(int_route1) & tolower(int_route1) != "local"){
    int_route1 <- paste0(int_route1, suffix)
    speed_row1 <- which(speed_full$ROUTE == int_route1 & 
                          speed_full$BEG_MP <= int_mp_1 & 
                          speed_full$END_MP >= int_mp_1)
    if(length(speed_row1) == 0 ){
      speed_row1 <- NA
    }
    # print(paste(speed_row1, "at", int_route1, int_mp_1))
  }
  if(!is.na(int_route2) & tolower(int_route2) != "local"){
    int_route2 <- paste0(int_route2, suffix)
    speed_row2 <- which(speed_full$ROUTE == int_route2 & 
                          speed_full$BEG_MP <= int_mp_2 & 
                          speed_full$END_MP >= int_mp_2)
    if(length(speed_row2) == 0 ){
      speed_row2 <- NA
    }
  }
  if(!is.na(int_route3) & tolower(int_route3) != "local"){
    int_route3 <- paste0(int_route3, suffix)
    speed_row3 <- which(speed_full$ROUTE == int_route3 & 
                          speed_full$BEG_MP <= int_mp_3 & 
                          speed_full$END_MP >= int_mp_3)
    if(length(speed_row3) == 0 ){
      speed_row3 <- NA
    }
  }
  if(!is.na(int_route4) & tolower(int_route4) != "local"){
    int_route4 <- paste0(int_route4, suffix)
    speed_row4 <- which(speed_full$ROUTE == int_route4 & 
                          speed_full$BEG_MP <= int_mp_4 & 
                          speed_full$END_MP >= int_mp_4)
    if(length(speed_row4) == 0 ){
      speed_row4 <- NA
    }
  }
  # fill values
  intersection[["INT_RT_1_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row1]
  intersection[["INT_RT_2_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row2]
  intersection[["INT_RT_3_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row3]
  intersection[["INT_RT_4_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row4]
}

# Add Functional Area for each approach
intersection$INT_RT_1_FA <- NA
intersection$INT_RT_2_FA <- NA
intersection$INT_RT_3_FA <- NA
intersection$INT_RT_4_FA <- NA
stopbar <- 60
for(i in 1:nrow(intersection)){
  route1 <- intersection[["INT_RT_1"]][i]
  route2 <- intersection[["INT_RT_2"]][i]
  route3 <- intersection[["INT_RT_3"]][i]
  route4 <- intersection[["INT_RT_4"]][i]
  route1_sl <- intersection[["INT_RT_1_SL"]][i]
  route2_sl <- intersection[["INT_RT_2_SL"]][i]
  route3_sl <- intersection[["INT_RT_3_SL"]][i]
  route4_sl <- intersection[["INT_RT_4_SL"]][i]
  dist1 <- 0
  dist2 <- 0
  dist3 <- 0
  dist4 <- 0
  if(!is.na(route1)){
    if(!is.na(route1_sl)){
      FA_row1 <- which(FA_ref$speed == route1_sl)
      dist1 <- FA_ref$total[FA_row1]
    } else{
      dist1 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_1_FA"]][i] <- dist1 + stopbar
  }
  if(!is.na(route2)){
    if(!is.na(route2_sl)){
      FA_row2 <- which(FA_ref$speed == route2_sl)
      dist2 <- FA_ref$total[FA_row2]
    } else{
      dist2 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_2_FA"]][i] <- dist2 + stopbar
  }
  if(!is.na(route3)){
    if(!is.na(route3_sl)){
      FA_row3 <- which(FA_ref$speed == route3_sl)
      dist3 <- FA_ref$total[FA_row3]
    } else{
      dist3 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_3_FA"]][i] <- dist3 + stopbar
  }
  if(!is.na(route4)){
    if(!is.na(route4_sl)){
      FA_row4 <- which(FA_ref$speed == route4_sl)
      dist4 <- FA_ref$total[FA_row4]
    } else{
      dist4 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_4_FA"]][i] <- dist4 + stopbar
  }
}

# Create FA file which has FA's for each route
ROUTE <- NA
BEG_MP <- NA
END_MP <- NA
row <- 0
for(i in 1:nrow(intersection)){
  for(j in 1:4){
    rt <- intersection[[paste0("INT_RT_",j)]][i]
    mp <- intersection[[paste0("INT_RT_",j,"_M")]][i]
    fa <- intersection[[paste0("INT_RT_",j,"_FA")]][i]
    if(rt %in% substr(main.routes,1,4)){
      row <- row + 1
      ROUTE[row] <- rt
      BEG_MP[row] <- mp - (fa/5280)
      END_MP[row] <- mp + (fa/5280)
    }
  }
}
FA <- tibble(ROUTE, BEG_MP, END_MP)

# Separate crashes into intersection and non-intersection related
crash$int_related <- FALSE
for(i in 1:nrow(crash)){
  rt <- crash$route[i]
  mp <- crash$milepoint[i]
  row <- which(FA$ROUTE == rt &
                 FA$BEG_MP < mp &
                 FA$END_MP > mp)
  if(length(row) > 0 | crash$intersection_related[i] == "Y"){
    crash$int_related[i] <- TRUE
  }
}

# Filter "Intersection Related" crashes
crash_seg <- crash %>% filter(int_related == FALSE)
crash_int <- crash %>% filter(int_related == TRUE)




###############################################
# Join Crash Vehicles to Crashes ##############
###############################################





###############################################
# Join Crash attributes to segments ###########
###############################################

att_name <- "crash_severity_id"
segs <- RC
csh <- crash_seg

# Add Crash Attributes Function
add_crash_attribute <- function(att_name, segs, csh){
  # set up dataframe for joining
  csh <- csh %>% select(seg_id, crash_year, !!att_name)
  att_name <- sym(att_name)
  # pivot wider
  csh <- csh %>%
    group_by(seg_id, crash_year, !!att_name) %>% 
    mutate(n = n()) %>% 
    pivot_wider(id_cols = c(seg_id, crash_year),
                names_from = !!att_name,
                names_prefix = paste0(att_name,"_"),
                values_from = n,
                values_fill = 0,
                values_fn = first)
  # join to segments
  segs <- left_join(segs, csh, by = c("SEG_ID"="seg_id", "YEAR"="crash_year"))
  # return segments
  return(segs)
}







############ random

# # change route 089A to 0011... There seems to be more to this though
# df$ROUTE_ID[df$ROUTE_ID == '089A'] <- 0011
# 
# # change route ID to 194 from 0085 for milepoint less than 2.977 
# df$ROUTE_ID[df$MILEPOINT < 2.977 & df$ROUTE_ID == 0085] <- 0194
# 
# # eliminate rows with ramps
# df %>%
#   as_tibble() %>%
#   filter(RAMP_ID == 0 | is.na(RAMP_ID))

###
## Unused Code
###

# The code below this point is old, but may have some useful parts

# CRASH DATA

```{r Load Crash Data}
location <- read_csv("data/Crash_Data_14-20.csv",show_col_types = F)
vehicle <- read_csv("data/Vehicle_Data_14-20.csv",show_col_types = F)
rollups <- read_csv("data/Rollups_14-20.csv",show_col_types = F)
```

```{r Check Crash Data Headers}
location <- check_location(location)
vehicle <- check_vehicle(vehicle)
rollups <- check_rollups(rollups)
```

```{r Join Crash Data}
crash <- left_join(location,rollups,by='crash_id')
full_crash <- left_join(crash,vehicle,by='crash_id')
rm("location","vehicle","rollups")
# crash <- filter_CAMS(crash)
```

```{r}
# write.csv(crash, "C:/Users/barrigat/Downloads/crashtest2.csv",row.names = F)
```

# Roadway Data

```{r Load Roadway Data}
# load AADT csv
aadt <- read_csv("data/AADT_Rounded.csv",show_col_types = F)
# load functional class csv
functional_class <- read_csv("data/Functional_Class_ALRS.csv",show_col_types = F)
# load intersections csv and convert to spatial frame
intersections <- read_csv("data/Intersection_w_ID_MEV_MP.csv") %>%
  st_as_sf(
    coords = c("BEG_LONG", "BEG_LAT"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912)
# load UTA stops shapefile
UTA_stops <- read_sf("data/UTA_Stops/UTA_Stops_and_Most_Recent_Ridership.shp") %>%
  st_transform(crs = 26912) %>%
  select(UTA_StopID) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)
# load schools (not college) shapefile
schools <- read_sf("data/Utah_Schools_PreK_to_12/Utah_Schools_PreK_to_12.shp") %>%
  st_transform(crs = 26912) %>%
  select(SchoolID) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)
# modify intersections to include bus stops and schools nearby
intersections <- mod_intersections(intersections,UTA_Stops,schools)
# remove unneeded dfs
rm("schools","UTA_stops")
# load lanes csv
lanes <- read_csv("data/Lanes.csv",show_col_types = F)
# load pavement messages csv
pavement_messages <- read_csv("data/Pavement_Messages.csv",show_col_types = F)
# load speed limit csv
speed_limit <- read_csv("data/UDOT_Speed_Limits_2019.csv",show_col_types = F)

# urban_code_large <- read_sf("data/Urban_Boundaries_Large.shp")
# urban_code_small <- read_sf("data/Urban_Boundaries_Small.shp")

# urban_code_test <- shapefile("data/Urban_Boundaries_Large")
```

```{r Check Roadway Data Headers}
aadt <- check_aadt(aadt)
functional_class <- check_functionalclass(functional_class)
intersections <- check_intersections(intersections)
lanes <- check_lanes(lanes)
pavement_messages <- check_pavementmessages(pavement_messages)
speed_limit <- check_speedlimit(speed_limit)
```

```{r Join Roadway Data}
road <- left_join(,,by='')
road <- left_join(,,by='')
road <- left_join(,,by='')
road <- left_join(,,by='')
road <- left_join(road,,by='')
rm("aadt","functional_class","intersections","lanes","pavement_messages","speed_limit")
```

```{r Temporary Load Excel GUI Data}
# The code before this should hopefully get us to this point. In the meanwhile,
# we can just use the file created by the Excel GUI and modify it as needed.
# CAMS_no_spatial <- read_csv("data/CAMSinput.csv")
# ISAM <- read_csv("data/UICPMinput.csv")
test <- RC
RC <- test

# CAMS from the Excel GUI doesn't include spatial data so we can add it in here.
# load sf data
sf <- read_sf("data/shapefile/UDOT_Routes_(ALRS).shp") %>%
  select(ROUTE_ALIA, BEG_MILEAG, END_MILEAG, SHAPE_Leng)
sf$ROUTE_ALIA <- paste(substr(sf$ROUTE_ALIA, 1, 6), "M", sep = "")
# append spatial data
CAMS <- right_join(sf, RC, by = c("ROUTE_ALIA"="ROUTE"), keep = TRUE) %>%
  select(-ROUTE_ALIA) %>%
  st_transform(crs = 26912) #convert to UTM Zone 12 Projection
# remove intermediate data
# rm(CAMS_no_spatial,sf)
# segment the routes based on segment milepoints
#CAMS_lineref <- ogrlineref(
#  l = CAMS$geometry,
#  mb = CAMS$BEG_MILEPOINT*1609.34,
#  me = CAMS$END_MILEPOINT*1609.34
#)

# write CAMS data to shapefile
CAMS_write <- CAMS %>%
  select(ROUTE, BEG_MP, END_MP, BEG_MILEAG, END_MILEAG, SHAPE_Leng) %>%
  group_by(ROUTE) %>%
  mutate(BEG_MP = case_when(
    BEG_MP == min(BEG_MP) ~ BEG_MILEAG,
    BEG_MP != min(BEG_MP) ~ BEG_MP
  ),
  Seg_Length = END_MP - BEG_MP
  ) %>%
  filter(Seg_Length > 0) %>%
  mutate(BEG_MP = case_when(
    BEG_MP == min(BEG_MP) ~ BEG_MILEAG,
    BEG_MP != min(BEG_MP) ~ BEG_MP
  ),
  BEG_MP = min(BEG_MP) + (BEG_MP-min(BEG_MP)) * (END_MILEAG - BEG_MILEAG)/(max(END_MP) - min(BEG_MP)),
  END_MP = min(BEG_MP) + (END_MP-min(BEG_MP)) * (END_MILEAG - BEG_MILEAG)/(max(END_MP) - min(BEG_MP)),
  Seg_Length = END_MP - BEG_MP
  ) %>%
  #mutate(END_MILEAG = END_MILEAG - BEG_MILEAG,
  #      BEG_MILEPOINT = BEG_MILEPOINT - BEG_MILEAG,
  #     END_MILEPOINT = END_MILEPOINT - BEG_MILEAG,
  #    BEG_MILEAG = 0
  #   ) %>%
  unique()
CAMS_write <- CAMS_write %>% rowid_to_column("SEG_ID")
st_write(CAMS_write,"data/shapefile/CAMS.shp")

CAMS_write <- CAMS %>%
  select(ROUTE, BEG_MILEAG, END_MILEAG, SHAPE_Leng) %>%
  #mutate(END_MILEAG = END_MILEAG - BEG_MILEAG,
  #      BEG_MILEAG = 0
  #     ) %>%
  unique()
st_write(CAMS_write,"data/shapefile/CAMS_routes.shp")

rm(CAMS_write)

```

```{python Geospatial Segmentation using Arcpy}

# -*- coding: utf-8 -*-
"""
Generated by ArcGIS ModelBuilder on : 2022-03-17 17:20:36
"""
import arcpy

def Model():  # Model
  
  # To allow overwriting outputs change overwriteOutput option to True.
  arcpy.env.overwriteOutput = False

# Model Environment settings
with arcpy.EnvManager(scratchWorkspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb", workspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb"):
  CAMS_routes_shp = "C:\\Users\\samyan\\Documents\\ArcGIS\\Projects\\test_CAMS-segmentation\\data\\CAMS_routes\\CAMS_routes.shp"
CAMS_shp = "C:\\Users\\samyan\\Documents\\ArcGIS\\Projects\\test_CAMS-segmentation\\data\\CAMS\\CAMS.shp"

# Process: Create Routes (Create Routes) (lr)
CAMS_routes_CreateRoutes = "C:\\Users\\samyan\\Documents\\ArcGIS\\Projects\\test_CAMS-segmentation\\test_CAMS-segmentation.gdb\\CAMS_routes_CreateRoutes"
with arcpy.EnvManager(scratchWorkspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb", workspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb"):
  arcpy.lr.CreateRoutes(in_line_features=CAMS_routes_shp, route_id_field="LABEL", out_feature_class=CAMS_routes_CreateRoutes, measure_source="TWO_FIELDS", from_measure_field="BEG_MIL", to_measure_field="END_MIL", coordinate_priority="UPPER_LEFT", measure_factor=1, measure_offset=0, ignore_gaps="IGNORE", build_index="INDEX")

# Process: Make Route Event Layer (Make Route Event Layer) (lr)
CAMS_Events = "CAMS Events1"
with arcpy.EnvManager(scratchWorkspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb", workspace=r"C:\Users\samyan\Documents\ArcGIS\Projects\test_CAMS-segmentation\test_CAMS-segmentation.gdb"):
  arcpy.lr.MakeRouteEventLayer(in_routes=CAMS_routes_CreateRoutes, route_id_field="LABEL", in_table=CAMS_shp, in_event_properties="LABEL Line BEG_MILEP END_MILEP", out_layer=CAMS_Events, offset_field="", add_error_field="ERROR_FIELD", add_angle_field="NO_ANGLE_FIELD", angle_type="NORMAL", complement_angle="ANGLE", offset_direction="LEFT", point_event_type="POINT")

if __name__ == '__main__':
  Model()

```

