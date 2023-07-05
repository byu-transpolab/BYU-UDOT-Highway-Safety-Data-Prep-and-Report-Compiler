library(tidyverse)
library(openxlsx)
source("R/RGUI_Functions.R")


# Read in Files
TOMS_seg <- read_csv("data/csv/EWRS_Seg_rank.csv")
TOMS_int <- read_csv("data/csv/EWRS_Int_rank.csv")
CAMS <- read_csv("data/output/CAMS_07Mar23_22_21_MA.csv")
ISAM <- read_csv("data/output/ISAM_07Mar23_22_21_MA.csv")
seg_pred <- read_csv("data/csv/PredictedSegCrashes.csv")
int_pred <- read_csv("data/csv/PredictedIntCrashes.csv")
crash_seg <- read_csv("data/temp/crash_seg.csv")
crash_int <- read_csv("data/temp/crash_int.csv")
vehicle <- read_csv("data/csv/Vehicle_File.csv")


# Create rank columns
TOMS_seg <- TOMS_seg %>% rowid_to_column("rank") %>% rename(SEG_ID = Seg_ID)
TOMS_int <- TOMS_int %>% rowid_to_column("rank") %>% rename(INT_ID = Int_ID)


# Join data to output files
TOMS_seg <- left_join(TOMS_seg, CAMS, by = "SEG_ID") 
TOMS_int <- left_join(TOMS_int, ISAM, by = c("INT_ID" = "Int_ID"))
TOMS_seg <- left_join(TOMS_seg, seg_pred, by = c("SEG_ID" = "Seg_ID")) 
TOMS_int <- left_join(TOMS_int, int_pred, by = c("INT_ID" = "Int_ID"))


# Format statistical output files
# Format segments results
latest_year <- max(TOMS_seg$YEAR, na.rm = TRUE) # finds latest year
TOMS_seg <- TOMS_seg %>%
  mutate(
    ROUTE_ID = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE)), # numeric route number
    Route_Name = ROUTE_ID, # not sure why this needs to be here, but they had it
    ROUTE = substr(ROUTE, 1, 5), # remove the M
    FC_Code = case_when( # assign FC Codes
      FUNCTIONAL_CLASS == "Interstate" ~ 1L,
      FUNCTIONAL_CLASS == "Other Freeways and Expressways" ~ 2L,
      FUNCTIONAL_CLASS == "Other Principal Arterial" ~ 3L,
      FUNCTIONAL_CLASS == "Minor Arterial" ~ 4L,
      FUNCTIONAL_CLASS == "Major Collector" ~ 5L,
      FUNCTIONAL_CLASS == "Minor Collector" ~ 6L,
      FUNCTIONAL_CLASS == "Local" ~ 7L
    ),
    Total_Percent = NUM_TRUCKS / AADT, # percent trucks
    Urban_Ru_1 = case_when( # urban code names
      URBAN_CODE == 99999L ~ "Rural",
      URBAN_CODE == 99998L ~ "Small Urban",
      URBAN_CODE == 78499L ~ "Salt Lake City",
      URBAN_CODE == 77446L ~ "St. George",
      URBAN_CODE == 72559L ~ "Provo-Orem",
      URBAN_CODE == 64945L ~ "Ogden - Layton",
      URBAN_CODE == 50959L ~ "Logan"
    ),
    Severe_Crashes = crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5,
    Sum_Total_Crashes = total_crashes,
    VMT = AADT * LENGTH_MILES,
    logVMT = log(VMT),
    hier = NA, # not sure what this is
    Percentile = percent_rank(rank),
    State_Rank = rank,
    Region_Rank = NA, # these require grouping so we'll do it later
    County_Rank = NA,
    Predicted_Total = Sev1+Sev2+Sev3+Sev4+Sev5
  ) %>%
  select(
    SEG_ID,
    LABEL = ROUTE,
    BEG_MILEPOINT = BEG_MP,
    END_MILEPOINT = END_MP,
    Seg_Length = LENGTH_MILES,
    YEAR,
    Route_Name,
    ROUTE_ID,
    DIRECTION = RouteDir,
    FC_Code,
    FC_Type = FUNCTIONAL_CLASS,
    COUNTY = COUNTY_CODE,
    REGION = UDOT_Region,
    AADT,
    # Single_Percent = , # We didn't care about these for our analysis
    # Combo_Percent = ,
    # Single_Count = ,
    # Combo_Count = ,
    Total_Percent,
    Total_Count = NUM_TRUCKS,
    SPEED_LIMIT,
    Num_Lanes = THRU_CNT,
    Urban_Rural = URBAN_CODE,
    Urban_Ru_1,
    Total_Crashes = total_crashes,
    Severe_Crashes,
    Sum_Total_Crashes,
    Sev_5_Crashes = crash_severity_id_5,
    Sev_4_Crashes = crash_severity_id_4,
    Sev_3_Crashes = crash_severity_id_3,
    Sev_2_Crashes = crash_severity_id_2,
    Sev_1_Crashes = crash_severity_id_1,
    ADVERSE_ROADWAY_SURF_CONDITION = adverse_roadway_surf_condition_crashes,
    ADVERSE_WEATHER = adverse_weather_crashes,
    AGGRESSIVE_DRIVING = aggressive_driving_crashes,
    BICYCLIST_INVOLVED = pedalcycle_involved_crashes,
    COLLISION_WITH_FIXED_OBJECT = collision_with_fixed_object_crashes,
    COMMERCIAL_MOTOR_VEH_INVOLVED = commercial_motor_veh_involved_crashes,
    DISTRACTED_DRIVING = distracted_driving_crashes,
    DOMESTIC_ANIMAL_RELATED = domestic_animal_related_crashes,
    DROWSY_DRIVING = drowsy_driving_crashes,
    DUI = dui_crashes,
    HEADON_COLLISION = manner_collision_id_3,
    # IMPROPER_RESTRAINT = improper_restraint_crashes,
    # INTERSECTION_RELATED = intersection_related_crashes,
    INTERSTATE_HIGHWAY = INTERSTATE,
    MOTORCYCLE_INVOLVED = motorcycle_involved_crashes,
    NIGHT_DARK_CONDITION = night_dark_condition_crashes,
    OLDER_DRIVER_INVOLVED = older_driver_involved_crashes,
    OVERTURN_ROLLOVER = overturn_rollover_crashes,
    PEDESTRIAN_INVOLVED = pedestrian_involved_crashes,
    RAILROAD_CROSSING = railroad_crossing_crashes,
    ROADWAY_DEPARTURE = roadway_departure_crashes,
    ROADWAY_GEOMETRY_RELATED = roadway_geometry_related_crashes,
    SINGLE_VEHICLE = single_vehicle_crashes,
    SPEED_RELATED = speed_related_crashes,
    TEENAGE_DRIVER_INVOLVED = teen_driver_involved_crashes,
    TRAIN_INVOLVED = train_involved_crashes,
    TRANSIT_VEHICLE_INVOLVED = transit_vehicle_involved_crashes,
    UNRESTRAINED = unrestrained_crashes,
    # URBAN_COUNTY,
    WILD_ANIMAL_RELATED = wild_animal_related_crashes,
    # WORK_ZONE_RELATED_YNU = work_zone_related_ynu_crashes,
    # (light_condition_id_1:collision_with_fixed_object_crashes),
    VMT,
    logVMT,
    hier,
    !!sym(paste0("Crashes",substr(latest_year,3,4))) := total_crashes,
    Predicted_Total,
    Percentile,
    State_Rank,
    Region_Rank,
    County_Rank,
    Predicted_Sev1 = Sev1,
    Predicted_Sev2 = Sev2,
    Predicted_Sev3 = Sev3,
    Predicted_Sev4 = Sev4,
    Predicted_Sev5 = Sev5,
    Median_Type,
    Right_Shoulder,
    THRU_CNT,
    THRU_WDTH,
    Driveway_Density_dpm
  )

# Format intersections results
latest_year <- max(TOMS_int$YEAR, na.rm = TRUE) # finds latest year
TOMS_int <- TOMS_int %>%
  mutate(
    ORIGINAL_INT_ID = MANDLI_ID, # not included in our latest files
    ROUTE = as.integer(gsub(".*?([0-9]+).*", "\\1", INT_RT_0)), # numeric route number
    INT_RT_1 = ifelse(INT_RT_1 == "Local", "Local", as.integer(gsub(".*?([0-9]+).*", "\\1", INT_RT_1))),
    INT_RT_2 = ifelse(INT_RT_2 == "Local", "Local", as.integer(gsub(".*?([0-9]+).*", "\\1", INT_RT_2))),
    INT_RT_3 = ifelse(INT_RT_3 == "Local", "Local", as.integer(gsub(".*?([0-9]+).*", "\\1", INT_RT_3))),
    INT_RT_4 = ifelse(INT_RT_4 == "Local", "Local", as.integer(gsub(".*?([0-9]+).*", "\\1", INT_RT_4))),
    CITY = sub("^([[:alpha:]]*).*", "\\1", STATION), # extract alpha-beta characters
    Group_Num = 1,
    MAX.FC_CODE = case_when( # assign FC Codes
      PRINCIPAL_FUNCTIONAL_CLASS == "Interstate" ~ 1L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Other Freeways and Expressways" ~ 2L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Other Principal Arterial" ~ 3L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Minor Arterial" ~ 4L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Major Collector" ~ 5L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Minor Collector" ~ 6L,
      PRINCIPAL_FUNCTIONAL_CLASS == "Local" ~ 7L
    ),
    PERCENT_TRUCKS = ENT_TRUCKS / ENT_VEH, # percent trucks
    MAX_ROAD_WIDTH = MAX_THRU_CNT * MAX_THRU_WDTH,
    MIN_ROAD_WIDTH = MIN_THRU_CNT * MIN_THRU_WDTH,
    URBAN_DESC = case_when( # urban code names
      URBAN_CODE == 99999L ~ "Rural",
      URBAN_CODE == 99998L ~ "Small Urban",
      URBAN_CODE == 78499L ~ "Salt Lake City",
      URBAN_CODE == 77446L ~ "St. George",
      URBAN_CODE == 72559L ~ "Provo-Orem",
      URBAN_CODE == 64945L ~ "Ogden - Layton",
      URBAN_CODE == 50959L ~ "Logan"
    ),
    Severe_Crashes = crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5,
    Sum_Total_Crashes = total_crashes,
    hier = NA, # not sure what this is
    Percentile = percent_rank(rank),
    cover = NA, # don't know what this is
    Region_Rank = NA, # these require grouping so we'll do it later
    County_Rank = NA,
    MIN.FC_TYPE = NA,  # we did principals, not max/min
    Predicted_Total = Sev1+Sev2+Sev3+Sev4+Sev5,
    INTERSTATE = ifelse(FUNCTIONAL_CLASS_Interstate==0,0,1)
  ) %>%
  select(
    INT_ID,
    ORIGINAL_INT_ID = MANDLI_ID,
    LATITUDE = lat,
    LONGITUDE = long,
    ELEVATION = BEG_ELEV,
    YEAR,
    TRAFFIC_CO,
    SR_SR,
    ROUTE,
    INT_RT_1,
    INT_RT_2,
    INT_RT_3,
    INT_RT_4,
    UDOT_BMP = INT_RT_0_M,
    INT_RT_1_M,
    INT_RT_2_M,
    INT_RT_3_M,
    INT_RT_4_M,
    INT_TYPE,
    CITY,
    REGION = UDOT_Region,
    Group_Num, # no idea what this is
    MAX.FC_CODE,
    # MIN.FC_CODE = , # we did principals, not max/min
    MAX.FC_TYPE = PRINCIPAL_FUNCTIONAL_CLASS,
    MIN.FC_TYPE,
    # INT_DIST_0 = , # functional area on a different file right now
    # INT_DIST_4 = ,
    PERCENT_TRUCKS,
    COUNTY = COUNTY_CODE,
    REGION.1 = UDOT_Region, # not sure why we need this again
    ENT_VEH,
    NUM_LEGS,
    MAX_NUM_LANES = MAX_THRU_CNT,
    MIN_NUM_LANES = MIN_THRU_CNT,
    MAX_ROAD_WIDTH,
    MIN_ROAD_WIDTH,
    MAX_SPEED_LIMIT,
    MIN_SPEED_LIMIT,
    URBAN_CODE,
    URBAN_DESC,
    Total_Crashes = total_crashes,
    Severe_Crashes,
    Sum_Total_Crashes,
    Sev_5_Crashes = crash_severity_id_5,
    Sev_4_Crashes = crash_severity_id_4,
    Sev_3_Crashes = crash_severity_id_3,
    Sev_2_Crashes = crash_severity_id_2,
    Sev_1_Crashes = crash_severity_id_1,
    ADVERSE_ROADWAY_SURF_CONDITION = adverse_roadway_surf_condition_crashes,
    ADVERSE_WEATHER = adverse_weather_crashes,
    AGGRESSIVE_DRIVING = aggressive_driving_crashes,
    BICYCLIST_INVOLVED = pedalcycle_involved_crashes,
    COLLISION_WITH_FIXED_OBJECT = collision_with_fixed_object_crashes,
    COMMERCIAL_MOTOR_VEH_INVOLVED = commercial_motor_veh_involved_crashes,
    DISTRACTED_DRIVING = distracted_driving_crashes,
    DOMESTIC_ANIMAL_RELATED = domestic_animal_related_crashes,
    DROWSY_DRIVING = drowsy_driving_crashes,
    DUI = dui_crashes,
    HEADON_COLLISION = manner_collision_id_3,
    # IMPROPER_RESTRAINT = improper_restraint_crashes,
    # INTERSECTION_RELATED = intersection_related_crashes,
    INTERSTATE_HIGHWAY = INTERSTATE,
    MOTORCYCLE_INVOLVED = motorcycle_involved_crashes,
    NIGHT_DARK_CONDITION = night_dark_condition_crashes,
    OLDER_DRIVER_INVOLVED = older_driver_involved_crashes,
    OVERTURN_ROLLOVER = overturn_rollover_crashes,
    PEDESTRIAN_INVOLVED = pedestrian_involved_crashes,
    # RAILROAD_CROSSING = railroad_crossing_crashes,
    ROADWAY_DEPARTURE = roadway_departure_crashes,
    ROADWAY_GEOMETRY_RELATED = roadway_geometry_related_crashes,
    SINGLE_VEHICLE = single_vehicle_crashes,
    SPEED_RELATED = speed_related_crashes,
    TEENAGE_DRIVER_INVOLVED = teen_driver_involved_crashes,
    TRAIN_INVOLVED = train_involved_crashes,
    TRANSIT_VEHICLE_INVOLVED = transit_vehicle_involved_crashes,
    UNRESTRAINED = unrestrained_crashes,
    # URBAN_COUNTY,
    WILD_ANIMAL_RELATED = wild_animal_related_crashes,
    # WORK_ZONE_RELATED_YNU = work_zone_related_ynu_crashes,
    # (light_condition_id_6:collision_with_fixed_object_crashes),
    hier,
    !!sym(paste0("Crashes",substr(latest_year,3,4))) := total_crashes,
    Predicted_Total,
    Percentile,
    cover,
    State_Rank = rank,
    Region_Rank,
    County_Rank,
    Predicted_Sev1 = Sev1,
    Predicted_Sev2 = Sev2,
    Predicted_Sev3 = Sev3,
    Predicted_Sev4 = Sev4,
    Predicted_Sev5 = Sev5
  )


# Reduce Data
TOMS_seg <- TOMS_seg %>%
  group_by(SEG_ID) %>%
  mutate(across(Total_Crashes:WILD_ANIMAL_RELATED, sum)) %>%
  ungroup() %>%
  filter(YEAR == latest_year)

TOMS_int <- TOMS_int %>%
  group_by(INT_ID) %>%
  mutate(across(Total_Crashes:WILD_ANIMAL_RELATED, sum)) %>%
  ungroup() %>%
  filter(YEAR == latest_year)


# Rank Data
TOMS_seg <- TOMS_seg %>%
  group_by(REGION) %>%
  mutate(Region_Rank = order(State_Rank)) %>%
  group_by(COUNTY) %>%
  mutate(County_Rank = order(State_Rank)) %>%
  ungroup()

TOMS_int <- TOMS_int %>%
  group_by(REGION) %>%
  mutate(Region_Rank = order(State_Rank)) %>%
  group_by(COUNTY) %>%
  mutate(County_Rank = order(State_Rank)) %>%
  ungroup()


# Determine filepath for stats files
CAMSstats <- paste0("CAMS_stats_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")
ISAMstats <- paste0("ISAM_stats_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")


# Save stats workbooks
wb <- createWorkbook()
addWorksheet(wb, sheetName = CAMSstats)
writeData(wb, sheet = 1, x = TOMS_seg, colNames = TRUE)
saveWorkbook(wb, file = paste0("data/output/",CAMSstats))

wb <- createWorkbook()
addWorksheet(wb, sheetName = ISAMstats)
writeData(wb, sheet = 1, x = TOMS_int, colNames = TRUE)
saveWorkbook(wb, file = paste0("data/output/",ISAMstats))






# Create Segments Parameters File
param_seg <- CAMS %>% select(SEG_ID:INTERSTATE) %>% unique()


# identify segments for each crash
crash_seg$seg_id <- NA
for (i in 1:nrow(crash_seg)){
  rt <- crash_seg$route[i]
  mp <- crash_seg$milepoint[i]
  seg_row <- which(param_seg$ROUTE == rt & 
                     param_seg$BEG_MP < mp & 
                     param_seg$END_MP > mp)
  if(length(seg_row) > 0){
    crash_seg[["seg_id"]][i] <- param_seg$SEG_ID[seg_row]
  }
}


# Conform Parameters File to necessary formatting
param_seg <- crash_seg %>%
  mutate(
    ROUTE_ID = as.integer(gsub(".*?([0-9]+).*", "\\1", route)),
    ramp_id = ifelse(ramp_id == 0, NA, ramp_id),
    route = substr(route, 1, 5)
  ) %>%
  select(
    SEG_ID = seg_id,
    CRASH_ID = crash_id,
    CRASH_DATETIME = crash_datetime, #needs reformatting in excel
    ROUTE_ID, #just the numeric portion
    DIRECTION = route_direction,
    LABEL = route, #remove the M
    RAMP_ID = ramp_id, #convert zeros to blanks
    MILEPOINT = milepoint,
    LATITUDE = lat,
    LONGITUDE = long,
    (number_vehicles_involved:collision_with_fixed_object),
    (crash_severity_id:roadway_contrib_circum_id),
    first_harmful_event_id
  ) %>%
  arrange(LABEL, MILEPOINT)


# Create Intersection Parameters File
param_int <- ISAM %>% select(Int_ID:TOTAL_THRU_CNT) %>% unique()
crash_int <- left_join(crash_int, param_int, by = c("int_id" = "Int_ID"))


# Conform Parameters File to necessary formatting
param_int <- crash_int %>%
  mutate(
    ROUTE_ID = as.integer(gsub(".*?([0-9]+).*", "\\1", route)),
    ramp_id = ifelse(ramp_id == 0, NA, ramp_id),
    route = substr(route, 1, 5)
  ) %>%
  select(
    INT_ID = int_id,
    CRASH_ID = crash_id,
    CRASH_DATETIME = crash_datetime, #needs reformatting in excel
    (number_vehicles_involved:collision_with_fixed_object),
    (crash_severity_id:roadway_contrib_circum_id),
    first_harmful_event_id,
    ROUTE_ID, #just the numeric portion
    LABEL = route, #remove the M
    DIRECTION = route_direction,
    RAMP_ID = ramp_id, #convert zeros to blanks
    MILEPOINT = milepoint,
    LATITUDE = lat.x,
    LONGITUDE = long.x
  ) %>%
  arrange(LABEL, MILEPOINT)


# Summarize vehicle details
vehicle <- vehicle %>%
  group_by(crash_id) %>%
  arrange(crash_id, vehicle_num) %>%
  summarize(
    event_sequence_1_id = case_when(
      nth(event_sequence_1_id, 1) < 88 ~ nth(event_sequence_1_id, 1),
      nth(event_sequence_1_id, 2) < 88 ~ nth(event_sequence_1_id, 2),
      nth(event_sequence_1_id, 3) < 88 ~ nth(event_sequence_1_id, 3),
      nth(event_sequence_1_id, 4) < 88 ~ nth(event_sequence_1_id, 4),
      nth(event_sequence_1_id, 5) < 88 ~ nth(event_sequence_1_id, 5),
      TRUE ~ nth(event_sequence_1_id, 1)
    ),
    event_sequence_2_id = case_when(
      nth(event_sequence_2_id, 1) < 88 ~ nth(event_sequence_2_id, 1),
      nth(event_sequence_2_id, 2) < 88 ~ nth(event_sequence_2_id, 2),
      nth(event_sequence_2_id, 3) < 88 ~ nth(event_sequence_2_id, 3),
      nth(event_sequence_2_id, 4) < 88 ~ nth(event_sequence_2_id, 4),
      nth(event_sequence_2_id, 5) < 88 ~ nth(event_sequence_2_id, 5),
      TRUE ~ nth(event_sequence_2_id, 1)
    ),
    event_sequence_3_id = case_when(
      nth(event_sequence_3_id, 1) < 88 ~ nth(event_sequence_3_id, 1),
      nth(event_sequence_3_id, 2) < 88 ~ nth(event_sequence_3_id, 2),
      nth(event_sequence_3_id, 3) < 88 ~ nth(event_sequence_3_id, 3),
      nth(event_sequence_3_id, 4) < 88 ~ nth(event_sequence_3_id, 4),
      nth(event_sequence_3_id, 5) < 88 ~ nth(event_sequence_3_id, 5),
      TRUE ~ nth(event_sequence_3_id, 1)
    ),
    event_sequence_4_id = case_when(
      nth(event_sequence_4_id, 1) < 88 ~ nth(event_sequence_4_id, 1),
      nth(event_sequence_4_id, 2) < 88 ~ nth(event_sequence_4_id, 2),
      nth(event_sequence_4_id, 3) < 88 ~ nth(event_sequence_4_id, 3),
      nth(event_sequence_4_id, 4) < 88 ~ nth(event_sequence_4_id, 4),
      nth(event_sequence_4_id, 5) < 88 ~ nth(event_sequence_4_id, 5),
      TRUE ~ nth(event_sequence_4_id, 1)
    ),
    most_harmful_event_id = case_when(
      nth(most_harmful_event_id, 1) < 88 ~ nth(most_harmful_event_id, 1),
      nth(most_harmful_event_id, 2) < 88 ~ nth(most_harmful_event_id, 2),
      nth(most_harmful_event_id, 3) < 88 ~ nth(most_harmful_event_id, 3),
      nth(most_harmful_event_id, 4) < 88 ~ nth(most_harmful_event_id, 4),
      nth(most_harmful_event_id, 5) < 88 ~ nth(most_harmful_event_id, 5),
      TRUE ~ nth(most_harmful_event_id, 1)
    ),
    vehicle_maneuver_id = case_when(
      nth(vehicle_maneuver_id, 1) < 88 ~ nth(vehicle_maneuver_id, 1),
      nth(vehicle_maneuver_id, 2) < 88 ~ nth(vehicle_maneuver_id, 2),
      nth(vehicle_maneuver_id, 3) < 88 ~ nth(vehicle_maneuver_id, 3),
      nth(vehicle_maneuver_id, 4) < 88 ~ nth(vehicle_maneuver_id, 4),
      nth(vehicle_maneuver_id, 5) < 88 ~ nth(vehicle_maneuver_id, 5),
      TRUE ~ nth(vehicle_maneuver_id, 1)
    )
  )


# Add vehicle details to parameters files
param_seg <- left_join(param_seg, vehicle, by = c("CRASH_ID" = "crash_id"))
param_int <- left_join(param_int, vehicle, by = c("CRASH_ID" = "crash_id"))


# Paste header details into parameter files
latest_year <- max(ISAM$YEAR) # finds latest year
earliest_year <- min(ISAM$YEAR) # finds earliest year
param_header_seg <- tibble(c("Severities", "Functional Area Definition", "Selected Years:"),
                           c(12345, "UDOT", paste0(earliest_year,"-",latest_year)))
param_header_int <- tibble(c("Severities:", "Functional Area Type:", "Selected Years:"),
                           c(12345, "UDOT", paste0(earliest_year,"-",latest_year)))


# Determine filepath for parameter files
CAMSparameters <- paste0("data/output/CAMS_parameters_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")
ISAMparameters <- paste0("data/output/ISAM_parameters_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")


# Save parameters workbooks
wb <- createWorkbook()
addWorksheet(wb, sheetName = "Parameters")

writeData(wb, sheet = 1, x = param_header_seg, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, sheet = 1, x = param_seg, startRow = 4, colNames = TRUE)

saveWorkbook(wb, file = CAMSparameters)

wb <- createWorkbook()
addWorksheet(wb, sheetName = "Parameters")

writeData(wb, sheet = 1, x = param_header_int, startRow = 1, colNames = FALSE, rowNames = FALSE)
writeData(wb, sheet = 1, x = param_int, startRow = 4, colNames = TRUE)

saveWorkbook(wb, file = ISAMparameters)

