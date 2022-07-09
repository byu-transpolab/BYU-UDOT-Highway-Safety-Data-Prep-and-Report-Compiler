###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Set Filepath and Column Names for each Dataset
###

routes_fp <- "data/csv/UDOT_Routes_ALRS.csv"
routes_col <- c("ROUTE_ID",                          
                "BEG_MILEAGE",                             
                "END_MILEAGE")

aadt_fp <- "data/csv/AADT_Unrounded.csv"
aadt_col <- c("ROUTE_NAME",
              "START_ACCU",
              "END_ACCUM",
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
              "CUTRK2014")

fc_fp <- "data/csv/Functional_Class_ALRS.csv"
fc_col <- c("ROUTE_ID",                          
            "FROM_MEASURE",                             
            "TO_MEASURE",
            "FUNCTIONAL_CLASS",                             
            "RouteDir",                             
            "RouteType")

speed_fp <- "data/csv/UDOT_Speed_Limits_2019.csv"
speed_col <- c("ROUTE_ID",
               "FROM_MEASURE",
               "TO_MEASURE",
               "SPEED_LIMIT")

lane_fp <- "data/csv/Lanes.csv"
lane_col <- c("ROUTE",
              "START_ACCUM",
              "END_ACCUM",
              "THRU_CNT",
              "THRU_WDTH")

urban_fp <- "data/csv/Urban_Code.csv"
urban_col <- c("ROUTE_ID",
               "FROM_MEASURE",
               "TO_MEASURE",
               "URBAN_CODE")

# intersection_fp <- "data/csv/Intersections.csv"
# intersection_col <- c("ROUTE",
#                       "START_ACCUM",
#                       "END_ACCUM",
#                       "ID",
#                       "INT_TYPE",
#                       "TRAFFIC_CO",
#                       "SR_SR",
#                       "INT_RT_1",
#                       "INT_RT_2",
#                       "INT_RT_3",
#                       "INT_RT_4",
#                       "STATION",
#                       "REGION",
#                       "BEG_LONG",
#                       "BEG_LAT",
#                       "BEG_ELEV")

intersection_fp <- "data/csv/Intersection_w_ID_MEV_MP.csv"
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

driveway_fp<- "data/csv/Driveway.csv"
driveway_col<- c("ROUTE",
                 "START_ACCUM",
                 "END_ACCUM",
                 "DIRECTION",
                 "TYPE")

median_fp <- "data/csv/Medians.csv"
median_col <- c("ROUTE_NAME",
                "START_ACCUM",
                "END_ACCUM",
                "MEDIAN_TYP",
                "TRFISL_TYP",
                "MDN_PRTCTN")

shoulder_fp <-"data/csv/Shoulders.csv"
shoulder_col <- c("ROUTE",
                  "START_ACCUM", # these columns are rounded quite a bit.
                  "END_ACCUM",
                  "UTPOSITION",
                  "SHLDR_WDTH")

location_fp <- "data/csv/Crash_Data_14-20.csv"
location_col <- c("crash_id", 
                  "crash_datetime", 
                  "route", 
                  "route_direction", 
                  "ramp_id", 
                  "milepoint", 
                  "lat", 
                  "long",
                  "crash_severity_id", 
                  "light_condition_id", 
                  "weather_condition_id", 
                  "manner_collision_id",
                  "roadway_surf_condition_id",
                  "roadway_junct_feature_id", 
                  "horizontal_alignment_id",
                  "vertical_alignment_id", 
                  "roadway_contrib_circum_id",
                  "first_harmful_event_id")

rollups_fp <- "data/csv/Rollups_14-20.csv"
rollups_col <- c("crash_id", 
                 "number_fatalities", 
                 "number_four_injuries",
                 "number_three_injuries", 
                 "number_two_injuries",
                 "number_one_injuries",
                 "pedalcycle_involved", 
                 "pedalcycle_involved_level4_tot",	
                 "pedalcycle_involved_fatal_tot",
                 "pedestrian_involved",	
                 "pedestrian_involved_level4_tot",	
                 "pedestrian_involved_fatal_tot",
                 "motorcycle_involved",	
                 "motorcycle_involved_level4_tot",	
                 "motorcycle_involved_fatal_tot",
                 "unrestrained", 
                 "dui", 
                 "aggressive_driving",	
                 "distracted_driving",	
                 "drowsy_driving",	
                 "speed_related",
                 "intersection_related",	
                 "adverse_weather",	
                 "adverse_roadway_surf_condition",	
                 "roadway_geometry_related",
                 "wild_animal_related",	
                 "domestic_animal_related",	
                 "roadway_departure",	
                 "overturn_rollover",	
                 "commercial_motor_veh_involved",	
                 "interstate_highway",	
                 "teen_driver_involved",	
                 "older_driver_involved",	
                 "route_type",	
                 "night_dark_condition",	
                 "single_vehicle",	
                 "train_involved",	
                 "railroad_crossing",
                 "transit_vehicle_involved",	
                 "collision_with_fixed_object")


vehicle_fp <- "data/csv/Vehicle_Data_14-20.csv"
vehicle_col <-  c("crash_id", 
                  "vehicle_num", 
                  "travel_direction_id", 
                  "event_sequence_1_id", 
                  "event_sequence_2_id", 
                  "event_sequence_3_id", 
                  "event_sequence_4_id",
                  "most_harmful_event_id", 
                  "vehicle_maneuver_id")

read_csv_file <- function(filepath, columns) {
  if (str_detect(filepath, ".csv")) {
    print("reading csv")
    read_csv(filepath) %>% select(all_of(columns))
  }
  else {
    print("Error reading in:")
    print(filepath)
  }
}
###
## Routes Data Prep
###

# Read in routes File
routes <- read_csv_file(routes_fp, routes_col)

# Standardize Column Names
names(routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
routes <- routes %>% filter(grepl("M", ROUTE))

# Round milepoints to 6 decimal places
routes <- routes %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Find Number of Unique Routes in the routes File
num.routes.routes <- routes %>% pull(ROUTE) %>% unique() %>% length()

# load divided routes data
div_routes <- read_csv("data/csv/DividedRoutesList_20220307_Adjusted.csv") 
names(div_routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
div_routes <- div_routes %>%
  mutate(LENGTH = END_MP - BEG_MP) %>%
  arrange(ROUTE, BEG_MP)

###
## Functional Class Data Prep
###

# Read in fc File
fc <- read_csv_file(fc_fp, fc_col)

# Standardize Column Names
names(fc)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
fc <- fc %>% filter(grepl("M", ROUTE))

# Create fed routes list
fed <- fc %>% filter(grepl("Fed Aid", RouteType))
fed.routes <- as.character(fed %>% pull(ROUTE) %>% unique() %>% sort())

# Select only State Routes
fc <- fc %>% filter(grepl("State", RouteType))

# Find Number of Unique Routes in the fc File
num.fc.routes <- fc %>% pull(ROUTE) %>% unique() %>% length()
main.routes <- as.character(fc %>% pull(ROUTE) %>% unique() %>% sort())

# Compress fc
fc <- compress_seg(fc)

# # Create routes dataframe
# routes <- fc %>% select(ROUTE, BEG_MP, END_MP)
# routes <- compress_seg(fc, c("ROUTE", "BEG_MP", "END_MP"), c("ROUTE"))

# fix ending endpoints
fc <- fix_endpoints(fc, routes)

####
## AADT Data Prep
####

# Read in aadt File
aadt <- read_csv_file(aadt_fp, aadt_col)

# Standardize Column Names
names(aadt)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Get Only Main Routes
aadt <- aadt %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in aadt file
num.aadt.routes <- aadt %>% pull(ROUTE) %>% unique() %>% length()

# Convert aadt to numeric
aadt <- aadt_numeric(aadt, aadt_col)

# Compress aadt
aadt <- compress_seg(aadt)

# fix negative aadt values
aadt <- aadt_neg(aadt, routes, div_routes)

# fix ending endpoints
aadt <- fix_endpoints(aadt, routes)

# fill in missing data
aadt <- fill_missing_aadt(aadt, aadt_col)

###
## Speed Limits Data Prep
###

# Read in speed File
speed <- read_csv_file(speed_fp, speed_col)

# Standardizing Column Names
names(speed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Getting Only Main Routes
speed <- speed %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in speed file
num.speed.routes <- speed %>% pull(ROUTE) %>% unique() %>% length()

# Compress speed
speed <- compress_seg(speed)
# speed <- compress_seg_alt(speed)

# fix ending endpoints
speed <- fix_endpoints(speed, routes)

###
## Lanes Data Prep
###

# Read in lane file
lane <- read_csv_file(lane_fp, lane_col)

# Standardize Column Names
names(lane)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Add M to Route Column to Standardize Route Format
lane$ROUTE <- paste(substr(lane$ROUTE, 1, 6), "M", sep = "")

# Get Only Main Routes
lane <- lane %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in lane file
num.lane.routes <- lane %>% pull(ROUTE) %>% unique() %>% length()

# Compress lanes
lane <- compress_seg(lane)     # note: the alt function is much slower for lanes

# fix ending endpoints
lane <- fix_endpoints(lane, routes)

###
## Urban Code Data Prep
###

# Read in urban code file
urban <- read_csv_file(urban_fp, urban_col)

# Standardize Column Names
names(urban)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
urban <- urban %>% filter(grepl("M", ROUTE))

# Round milepoints to 6 decimal places
urban <- urban %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Get Only Main Routes
urban <- urban %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in urban code file
num.urban.routes <- urban %>% pull(ROUTE) %>% unique() %>% length()

# Compress urban code
urban <- compress_seg(urban)     

# fix ending endpoints
urban <- fix_endpoints(urban, routes)



###
## Intersection Data Prep
###
#Intersections are labeled at a point, not a segment

# Read in intersection file
intersection <- read_csv_file(intersection_fp, intersection_col)

# Standardize Column Names
names(intersection)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Add M to Route Column to Standardize Route Format
intersection$ROUTE <- paste(substr(intersection$ROUTE, 1, 5), "M", sep = "")

# Getting only state routes
intersection <- intersection %>% filter(ROUTE %in% substr(main.routes, 1, 6))

# Find Number of Unique Routes in shoulder file
num.intersection.routes <- intersection %>% pull(ROUTE) %>% unique() %>% length()

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

# Sort intersections
intersection <- intersection %>% arrange(ROUTE, BEG_MP)

# Make sure Intersection IDs are unique
n <- 1
for(i in 2:nrow(intersection)){
  if(intersection$Int_ID[i] == intersection$Int_ID[i-n]){
    intersection$Int_ID[i] <- paste0(intersection$Int_ID[i],"_(",n,")")
    n <- n + 1
  } else{
    n <- 1
  }
}

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
MP <- NA
Int_ID <- NA
row <- 0
for(i in 1:nrow(intersection)){
  id <- intersection[["Int_ID"]][i]
  for(j in 1:4){
    rt <- intersection[[paste0("INT_RT_",j)]][i]
    mp <- intersection[[paste0("INT_RT_",j,"_M")]][i]
    fa <- intersection[[paste0("INT_RT_",j,"_FA")]][i]
    if(rt %in% substr(main.routes,1,4)){
      row <- row + 1
      if(rt == substr(intersection$ROUTE[i],0,4)){
        ROUTE[row] <- intersection$ROUTE[i]
      } else{
        ROUTE[row] <- paste0(rt,"PM")   # assuming if direction not specified it is positive
      }
      BEG_MP[row] <- mp - (fa/5280)
      END_MP[row] <- mp + (fa/5280)
      MP[row] <- mp
      Int_ID[row] <- id
    }
  }
}
FA <- tibble(ROUTE, BEG_MP, END_MP, MP, Int_ID)




###
## Shoulder Data Prep
###

# Read in shoulder file
shoulder <- read_csv_file(shoulder_fp, shoulder_col)

# Standardize Column Names
names(shoulder)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Getting rid of ramps
shoulder <- shoulder %>% filter(nchar(ROUTE) == 5)

# Add M to Route Column to Standardize Route Format
shoulder$ROUTE <- paste(substr(shoulder$ROUTE, 1, 6), "M", sep = "")

# Getting only state routes
shoulder <- shoulder %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in shoulder file
num.shoulder.routes <- shoulder %>% pull(ROUTE) %>% unique() %>% length()

# Create Point to Reference Shoulders
shoulder <- shoulder %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))

# filter out seemingly duplicated shoulders for observation
shd_disc <- shoulder %>%
  group_by(ROUTE, BEG_MP, UTPOSITION) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)

###
## Median Data Prep
###

# Read in median file
median <- read_csv_file(median_fp, median_col)

# Standardize Column Names
names(median)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

#Getting rid of ramps
median <- median %>% filter(nchar(ROUTE) == 5)

# Add M to Route Column to Standardize Route Format
median$ROUTE <- paste(substr(median$ROUTE, 1, 6), "M", sep = "")

# Getting only state routes
median <- median %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in median file
num.median.routes <- median %>% pull(ROUTE) %>% unique() %>% length()

# Compress Medians
median <- compress_seg(median, variables = c("MEDIAN_TYP", "TRFISL_TYP", "MDN_PRTCTN"))

# Create Point to Reference Driveways
median <- median %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))

###
## Driveway Data Prep
###

# Read in driveway file
driveway <- read_csv_file(driveway_fp, driveway_col)

# Standardize Column Names
names(driveway)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Adding M to Route
driveway$ROUTE <- paste(substr(driveway$ROUTE, 1, 6), "M", sep = "")

# Create Point to Reference Driveways
driveway <- driveway %>% mutate(MP = (BEG_MP+END_MP)/2)

# compress points
driveway <- driveway %>% unique()

#filter out seemingly duplicated driveways for observation
drv_disc <- driveway %>%
  group_by(ROUTE, BEG_MP) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)

###
## Read in crash Files
###

location <- read_csv_file(location_fp, location_col)
rollups <- read_csv_file(rollups_fp, rollups_col)
vehicle <- read_csv_file(vehicle_fp, vehicle_col)

###
## Use fc file to build a gaps dataset
###

gaps <- read_csv_file(fc_fp, fc_col)
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