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
            "RouteType",
            "COUNTY_CODE")

speed_fp <- "data/csv/2021_Statewide_Speed_Limits.csv"
speed_col <- c("Route",
               "Beg_MP",
               "End_MP",
               "Speed Limit")

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
                      "INT_RT_1",
                      "INT_RT_2",
                      "INT_RT_3",
                      "INT_RT_4",
                      "UDOT_BMP",
                      "INT_RT_1_M",
                      "INT_RT_2_M",
                      "INT_RT_3_M",
                      "INT_RT_4_M",
                      "Int_ID",
                      "INT_TYPE",
                      "TRAFFIC_CO",
                      "SR_SR",
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

regions_fp <- "data/csv/UDOT_Regions.csv"
regions_col <- c("COUNTY_CODE", 
                 "UDOT_Region")

###
## Routes Data Prep
###

# Read in routes File
routes <- read_csv_file(routes_fp, routes_col)

# Standardize Column Names
names(routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select Only Main Routes
routes <- routes %>% filter(grepl("M", ROUTE))

# Round Milepoints to 6 Decimal Places
routes <- routes %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Find Number of Unique Routes in the routes File
num_routes_routes <- routes %>% pull(ROUTE) %>% unique() %>% length()

# Load divided routes Data
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

# Compress fc
fc <- compress_seg(fc)

# Fix Missing Data
fc <- fc %>% mutate(COUNTY_CODE = if_else(ROUTE == "0131PM","Salt Lake", COUNTY_CODE),
                    COUNTY_CODE = if_else(ROUTE == "0190PM","Salt Lake", COUNTY_CODE),
                    COUNTY_CODE = if_else(ROUTE == "0209PM","Salt Lake", COUNTY_CODE))

# Create Full fc File (Including Fed Routes) For Merging with Intersections
fc_full <- fc

# Select Only Main Routes
fc <- fc %>% filter(grepl("M", ROUTE))

# Create fed routes List
fed <- fc %>% filter(grepl("Fed Aid", RouteType))
fed_routes <- as.character(fed %>% pull(ROUTE) %>% unique() %>% sort())

# Select only State Routes
fc <- fc %>% filter(grepl("State", RouteType))

# Find Number of Unique Routes in the fc File
num_fc_routes <- fc %>% pull(ROUTE) %>% unique() %>% length()
state_routes <- as.character(fc %>% pull(ROUTE) %>% unique() %>% sort())  

# Fix Endpoints
fc <- fix_endpoints(fc, routes)


###
## Regions Data Prep
###

# Read in regions data set
regions <- read_csv_file(regions_fp, regions_col)

# Left Join to fc
fc <- left_join(fc, regions, "COUNTY_CODE")
fc_full <- left_join(fc_full, regions, "COUNTY_CODE")


####
## AADT Data Prep
####

# Read in aadt File
aadt <- read_csv_file(aadt_fp, aadt_col)

# Standardize Column Names
names(aadt)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Convert aadt to Numeric
aadt <- aadt_numeric(aadt, aadt_col)

# Compress aadt
aadt <- compress_seg(aadt)

# Fill in Missing aadt
aadt <- fill_missing_aadt(aadt, aadt_col)

# Create Full aadt File (Including Fed Routes) For Merging with Intersections
aadt_full <- aadt

# Select Only Main Routes
aadt <- aadt %>% 
  filter(grepl("M", ROUTE)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in aadt file
num_aadt_routes <- aadt %>% pull(ROUTE) %>% unique() %>% length()

# Fix Negative aadt Values
aadt <- aadt_neg(aadt, routes, div_routes)

# Fix Endpoints
aadt <- fix_endpoints(aadt, routes)

# Get Only State Routes
aadt <- aadt %>% filter(ROUTE %in% substr(state_routes, 1, 6))


###
## Speed Limits Data Prep
###

# Read in speed File
speed <- read_csv_file(speed_fp, speed_col)

# Standardizing Column Names
names(speed)[c(1:4)] <- c("ROUTE", "BEG_MP", "END_MP", "SPEED_LIMIT")

# Compress speed
speed <- compress_seg(speed)

# Create Full speed File (Including Fed Routes) For Merging with Intersections
speed_full <- speed

# Find Number of Unique Routes in speed File
num_speed_routes <- speed %>% pull(ROUTE) %>% unique() %>% length()

# Fix Endpoints
speed <- fix_endpoints(speed, routes)


###
## Lanes Data Prep
###

# Read in lane File
lane <- read_csv_file(lane_fp, lane_col)

# Standardize Column Names
names(lane)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Compress lanes
lane <- compress_seg(lane)

# Add M to Route Column to Standardize Route Format
lane$ROUTE <- paste(substr(lane$ROUTE, 1, 6), "M", sep = "")

# Create Full lane File (Including Fed Routes) For Merging with Intersections
lane_full <- lane

# Select Only Main Routes
lane <- lane %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in lanes File
num_lane_routes <- lane %>% pull(ROUTE) %>% unique() %>% length()

# Get Only State Routes
lane <- lane %>% filter(ROUTE %in% substr(state_routes, 1, 6))

# Fix Endpoints
lane <- fix_endpoints(lane, routes)


###
## Urban Code Data Prep
###

# Read in urban code File
urban <- read_csv_file(urban_fp, urban_col)

# Standardize Column Names
names(urban)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Compress urban code
urban <- compress_seg(urban)

# Create Full urban code File (Including Fed Routes) For Merging with Intersections
urban_full <- urban

# Select Only Main Routes
urban <- urban %>% filter(grepl("M", ROUTE)) %>%
  filter(BEG_MP < END_MP)

# Round Milepoints to 6 Decimal Places
urban <- urban %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Find Number of Unique Routes in urban code File
num_urban_routes <- urban %>% pull(ROUTE) %>% unique() %>% length()

# Get Only State Routes
urban <- urban %>% filter(ROUTE %in% substr(state_routes, 1, 6))

# Fix Endpoints
urban <- fix_endpoints(urban, routes)


###
## Intersection Data Prep
###

# Intersections are labeled at a point, not a segment

# Read in intersection File
intersection <- read_csv_file(intersection_fp, intersection_col)

# Standardize Column Names
names(intersection)[c(1,6)] <- c("INT_RT_0", "INT_RT_0_M")

# Organize Columns
intersection <- intersection %>% select(Int_ID:BEG_ELEV, everything())

# Add M to Primary Route Column to Standardize Route Format
intersection$INT_RT_0 <- paste(substr(intersection$INT_RT_0, 1, 5), "M", sep = "")

# Get Only State Routes  
intersection <- intersection %>% 
  filter(
    INT_RT_0 %in% substr(state_routes, 1,4) |
    INT_RT_1 %in% substr(state_routes, 1,4) |
    INT_RT_2 %in% substr(state_routes, 1,4) |
    INT_RT_3 %in% substr(state_routes, 1,4) |
    INT_RT_4 %in% substr(state_routes, 1,4) 
  )

# Find Number of Unique primary Routes in intersections File
num_intersection_routes <- intersection %>% pull(INT_RT_0) %>% unique() %>% length()

# Filter SR_SR, Fed_Aid, and Signalized
intersection <- intersection %>%
  filter(
    SR_SR == "YES" | 
    TRAFFIC_CO == "SIGNAL" | 
    INT_RT_0 %in% substr(fed_routes,1,4) |
    INT_RT_1 %in% substr(fed_routes,1,4) |
    INT_RT_2 %in% substr(fed_routes,1,4) | 
    INT_RT_3 %in% substr(fed_routes,1,4) |
    INT_RT_4 %in% substr(fed_routes,1,4)
  )

# Ensure Intersection IDs are Unique
intersection <- intersection %>% arrange(Int_ID)
n <- 1
for(i in 2:nrow(intersection)){
  if(intersection$Int_ID[i] == intersection$Int_ID[i-n]){
    intersection$Int_ID[i] <- paste0(intersection$Int_ID[i],"_(",n,")")
    n <- n + 1
  } else{
    n <- 1
  }
}

# Sort intersections
intersection <- intersection %>% arrange(INT_RT_0, INT_RT_0_M)

# Create Functional Area Reference Table
# FA_ref <- tibble(
#   speed = c(5,10,15,20,25,30,35,40,45,50,55,60,65,70,75,80), 
#   d1 = c(75,75,75,75,90,110,130,145,165,185,200,220,240,255,275,275), 
#   d2 = c(70,70,70,70,105,150,225,290,360,440,525,655,755,875,995,995)
# ) %>%
#   mutate(
#     d3 = 50,
#     total = d1 + d2 + d3
#   )
FA_ref <- tibble(
  INT_TYPE = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,"ROUNDABOUT","CFI OFFSET LEFT TURN",
               "THRU TURN CENTRAL","THRU TURN OFFSET U-TURN","SPUI","DDI",
               "MID-BLOCK"),
  TRAFFIC_CO = c("SIGNAL","STOP SIGN - SIDE STREET","STOP SIGN - ALL WAY",
                 "STOP SIGN","YIELD SIGN","YIELD SIGN - ALL WAY",
                 "YIELD SIGN - SIDE STREET","UNCONTROLLED","OTHER",
                 NA,NA,NA,NA,NA,NA,NA),                     
  d = list(300,150,100,150,100,100,100,100,150,300,c(400,300),c(400,200),c(400,300),500,
           400,100)
)

# Add Speed Limits for Each Intersection Approach on intersections File
intersection$INT_RT_0_SL <- NA
intersection$INT_RT_1_SL <- NA
intersection$INT_RT_2_SL <- NA
intersection$INT_RT_3_SL <- NA
intersection$INT_RT_4_SL <- NA
for (i in 1:nrow(intersection)){
  int_route0 <- intersection[["INT_RT_0"]][i]
  int_route1 <- intersection[["INT_RT_1"]][i]
  int_route2 <- intersection[["INT_RT_2"]][i]
  int_route3 <- intersection[["INT_RT_3"]][i]
  int_route4 <- intersection[["INT_RT_4"]][i]
  int_mp_0 <- intersection[["INT_RT_0_M"]][i]
  int_mp_1 <- intersection[["INT_RT_1_M"]][i]
  int_mp_2 <- intersection[["INT_RT_2_M"]][i]
  int_mp_3 <- intersection[["INT_RT_3_M"]][i]
  int_mp_4 <- intersection[["INT_RT_4_M"]][i]
  speed_row0 <- NA
  speed_row1 <- NA
  speed_row2 <- NA
  speed_row3 <- NA
  speed_row4 <- NA
  # Get Route Direction
  suffix = substr(intersection$INT_RT_0[i],5,6)
  # Find Rows in speed File
  if(!is.na(int_route0) & tolower(int_route0) != "local"){
    speed_row0 <- which(speed_full$ROUTE == int_route0 & 
                          speed_full$BEG_MP <= int_mp_0 & 
                          speed_full$END_MP >= int_mp_0)
    if(length(speed_row0) == 0 ){
      speed_row0 <- NA
    }
    # print(paste(speed_row1, "at", int_route1, int_mp_1))
  }
  if(!is.na(int_route1) & tolower(int_route1) != "local"){
    # Determine the Direction
    # Assume Same as Primary Route if Same Route Number and "PM" Otherwise
    if(int_route1 == substr(int_route0,1,4)){
      int_route1 <- paste0(int_route1, suffix)
    } else{int_route1 <- paste0(int_route1, "PM")}
    speed_row1 <- which(speed_full$ROUTE == int_route1 & 
                          speed_full$BEG_MP <= int_mp_1 & 
                          speed_full$END_MP >= int_mp_1)
    if(length(speed_row1) == 0 ){
      speed_row1 <- NA
    }
  }
  if(!is.na(int_route2) & tolower(int_route2) != "local"){
    if(int_route2 == substr(int_route0,1,4)){
      int_route2 <- paste0(int_route2, suffix)
    } else{int_route2 <- paste0(int_route2, "PM")}
    speed_row2 <- which(speed_full$ROUTE == int_route2 & 
                          speed_full$BEG_MP <= int_mp_2 & 
                          speed_full$END_MP >= int_mp_2)
    if(length(speed_row2) == 0 ){
      speed_row2 <- NA
    }
  }
  if(!is.na(int_route3) & tolower(int_route3) != "local"){
    if(int_route3 == substr(int_route0,1,4)){
      int_route3 <- paste0(int_route3, suffix)
    } else{int_route3 <- paste0(int_route3, "PM")}
    speed_row3 <- which(speed_full$ROUTE == int_route3 & 
                          speed_full$BEG_MP <= int_mp_3 & 
                          speed_full$END_MP >= int_mp_3)
    if(length(speed_row3) == 0 ){
      speed_row3 <- NA
    }
  }
  if(!is.na(int_route4) & tolower(int_route4) != "local"){
    if(int_route4 == substr(int_route0,1,4)){
      int_route4 <- paste0(int_route4, suffix)
    } else{int_route4 <- paste0(int_route4, "PM")}
    speed_row4 <- which(speed_full$ROUTE == int_route4 & 
                          speed_full$BEG_MP <= int_mp_4 & 
                          speed_full$END_MP >= int_mp_4)
    if(length(speed_row4) == 0 ){
      speed_row4 <- NA
    }
  }
  # fill values
  intersection[["INT_RT_0_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row0]
  intersection[["INT_RT_1_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row1]
  intersection[["INT_RT_2_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row2]
  intersection[["INT_RT_3_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row3]
  intersection[["INT_RT_4_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row4]
}

# Add Functional Area for Each Approach
intersection$INT_RT_0_FA <- NA
intersection$INT_RT_1_FA <- NA
intersection$INT_RT_2_FA <- NA
intersection$INT_RT_3_FA <- NA
intersection$INT_RT_4_FA <- NA
# stopbar <- 60
for(i in 1:nrow(intersection)){
  # Assign values to variables
  route0 <- intersection[["INT_RT_0"]][i]
  route1 <- intersection[["INT_RT_1"]][i]
  route2 <- intersection[["INT_RT_2"]][i]
  route3 <- intersection[["INT_RT_3"]][i]
  route4 <- intersection[["INT_RT_4"]][i]
  type <- intersection$INT_TYPE[i]
  traffic_co <- intersection$TRAFFIC_CO[i]
  # Determine functional area distance from FA_ref table
  fa_row <- which(FA_ref$INT_TYPE == type)
  if(length(fa_row) == 0){
    fa_row <- which(FA_ref$TRAFFIC_CO == traffic_co)
  }
  dist <- as.numeric(unlist(FA_ref$d[fa_row]))
  # Assign functional area distance to each applicable leg
  if(!is.na(route0)){
    intersection$INT_RT_0_FA[i] <- dist[1]
  }
  if(!is.na(route1)){
    if(length(dist) > 1 & route1 != substrMinusRight(route0,2)){
      intersection$INT_RT_1_FA[i] <- dist[2]
    } else{
      intersection$INT_RT_1_FA[i] <- dist[1]
    }
  }
  if(!is.na(route2)){
    if(length(dist) > 1 & route2 != substrMinusRight(route0,2)){
      intersection$INT_RT_2_FA[i] <- dist[2]
    } else{
      intersection$INT_RT_2_FA[i] <- dist[1]
    }
  }
  if(!is.na(route3)){
    if(length(dist) > 1 & route3 != substrMinusRight(route0,2)){
      intersection$INT_RT_3_FA[i] <- dist[2]
    } else{
      intersection$INT_RT_3_FA[i] <- dist[1]
    }
  }
  if(!is.na(route4)){
    if(length(dist) > 1 & route4 != substrMinusRight(route0,2)){
      intersection$INT_RT_4_FA[i] <- dist[2]
    } else{
      intersection$INT_RT_4_FA[i] <- dist[1]
    }
  }
  
  # route0 <- intersection[["INT_RT_0"]][i]
  # route1 <- intersection[["INT_RT_1"]][i]
  # route2 <- intersection[["INT_RT_2"]][i]
  # route3 <- intersection[["INT_RT_3"]][i]
  # route4 <- intersection[["INT_RT_4"]][i]
  # route0_sl <- intersection[["INT_RT_0_SL"]][i]
  # route1_sl <- intersection[["INT_RT_1_SL"]][i]
  # route2_sl <- intersection[["INT_RT_2_SL"]][i]
  # route3_sl <- intersection[["INT_RT_3_SL"]][i]
  # route4_sl <- intersection[["INT_RT_4_SL"]][i]
  # dist0 <- 0
  # dist1 <- 0
  # dist2 <- 0
  # dist3 <- 0
  # dist4 <- 0
  # if(!is.na(route0)){
  #   if(!is.na(route0_sl)){
  #     FA_row0 <- which(FA_ref$speed == route0_sl)
  #     dist0 <- FA_ref$total[FA_row0]
  #   } else{
  #     dist0 <- 250   # This is the default if there is no Speed Limit
  #   }
  #   intersection[["INT_RT_0_FA"]][i] <- dist0 + stopbar
  # }
  # if(!is.na(route1)){
  #   if(!is.na(route1_sl)){
  #     FA_row1 <- which(FA_ref$speed == route1_sl)
  #     dist1 <- FA_ref$total[FA_row1]
  #   } else{
  #     dist1 <- 250   # This is the default if there is no Speed Limit
  #   }
  #   intersection[["INT_RT_1_FA"]][i] <- dist1 + stopbar
  # }
  # if(!is.na(route2)){
  #   if(!is.na(route2_sl)){
  #     FA_row2 <- which(FA_ref$speed == route2_sl)
  #     dist2 <- FA_ref$total[FA_row2]
  #   } else{
  #     dist2 <- 250   # This is the default if there is no Speed Limit
  #   }
  #   intersection[["INT_RT_2_FA"]][i] <- dist2 + stopbar
  # }
  # if(!is.na(route3)){
  #   if(!is.na(route3_sl)){
  #     FA_row3 <- which(FA_ref$speed == route3_sl)
  #     dist3 <- FA_ref$total[FA_row3]
  #   } else{
  #     dist3 <- 250   # This is the default if there is no Speed Limit
  #   }
  #   intersection[["INT_RT_3_FA"]][i] <- dist3 + stopbar
  # }
  # if(!is.na(route4)){
  #   if(!is.na(route4_sl)){
  #     FA_row4 <- which(FA_ref$speed == route4_sl)
  #     dist4 <- FA_ref$total[FA_row4]
  #   } else{
  #     dist4 <- 250   # This is the default if there is no Speed Limit
  #   }
  #   intersection[["INT_RT_4_FA"]][i] <- dist4 + stopbar
  # }
}

# Create FA File which has FA for Each Route
ROUTE <- NA
BEG_MP <- NA
END_MP <- NA
MP <- NA
Int_ID <- NA
row <- 0
for(i in 1:nrow(intersection)){
  id <- intersection[["Int_ID"]][i]
  for(j in 0:4){
    rt <- intersection[[paste0("INT_RT_",j)]][i]
    mp <- intersection[[paste0("INT_RT_",j,"_M")]][i]
    fa <- intersection[[paste0("INT_RT_",j,"_FA")]][i]
    if(!is.na(rt) & tolower(rt) != "local"){   #old condition:  rt %in% substr(state_routes,1,4)
      row <- row + 1
      if(substr(rt,0,4) == substr(intersection$INT_RT_0[i],0,4)){
        ROUTE[row] <- intersection$INT_RT_0[i]
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

# Create Functional Area Table
FA <- tibble(ROUTE, BEG_MP, END_MP, MP, Int_ID) %>% unique()


###
## Shoulder Data Prep
###

# Read in shoulder File
shoulder <- read_csv_file(shoulder_fp, shoulder_col)

# Standardize Column Names
names(shoulder)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Remove Ramps
shoulder <- shoulder %>% filter(nchar(ROUTE) == 5)

# Add M to Route Column to Standardize Route Format
shoulder$ROUTE <- paste(substr(shoulder$ROUTE, 1, 6), "M", sep = "")

# Get Only State Routes
shoulder <- shoulder %>% filter(ROUTE %in% substr(state_routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in shoulder File
num_shoulder_routes <- shoulder %>% pull(ROUTE) %>% unique() %>% length()

# Create Point to Reference Shoulders
shoulder <- shoulder %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))

# Filter Out Duplicated Shoulders for Observation
shd_disc <- shoulder %>%
  group_by(ROUTE, BEG_MP, UTPOSITION) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)


###
## Median Data Prep
###

# Read in median File
median <- read_csv_file(median_fp, median_col)

# Standardize Column Names
names(median)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Remove Ramps
median <- median %>% filter(nchar(ROUTE) == 5)

# Add M to Route Column to Standardize Route Format
median$ROUTE <- paste(substr(median$ROUTE, 1, 6), "M", sep = "")

# Get Only State Routes
median <- median %>% filter(ROUTE %in% substr(state_routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in median File
num_median_routes <- median %>% pull(ROUTE) %>% unique() %>% length()

# Compress Medians
median <- compress_seg(median, variables = c("MEDIAN_TYP", "TRFISL_TYP", "MDN_PRTCTN"))

# Create Point to Reference Driveways
median <- median %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))


###
## Driveway Data Prep
###

# Read in driveway File
driveway <- read_csv_file(driveway_fp, driveway_col)

# Standardize Column Names
names(driveway)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Adding M to Route
driveway$ROUTE <- paste(substr(driveway$ROUTE, 1, 6), "M", sep = "")

# Create Point to Reference Driveways
driveway <- driveway %>% mutate(MP = (BEG_MP+END_MP)/2)

# Compress Points
driveway <- driveway %>% unique()

# Filter Out Duplicated Driveways for Observation
drv_disc <- driveway %>%
  group_by(ROUTE, BEG_MP) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)


###
## Read in Crash Files
###

location <- read_csv_file(location_fp, location_col)
rollups <- read_csv_file(rollups_fp, rollups_col)
vehicle <- read_csv_file(vehicle_fp, vehicle_col)


###
## Build Gaps Dataset
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


###
## Load school and UTA files
###

# Load UTA Stops shapefile
UTA_stops <- read_sf("data/shapefile/UTA_Stops_and_Most_Recent_Ridership.shp") %>%
  st_transform(crs = 26912) %>%
  select(UTA_StopID, Mode) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)
# Load Schools (K-12) shapefile
schools <- read_sf("data/shapefile/Utah_Schools_PreK_to_12.shp") %>%
  st_transform(crs = 26912) %>%
  select(SchoolID, SchoolLeve, OnlineScho, SchoolType, TotalK12) %>%
  filter(is.na(OnlineScho), SchoolType == "Vocational" | SchoolType == "Special Education" | SchoolType == "Residential Treatment" | SchoolType == "Regular Education" | SchoolType == "Alternative") %>%
  select(-OnlineScho, -SchoolType, -TotalK12) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)