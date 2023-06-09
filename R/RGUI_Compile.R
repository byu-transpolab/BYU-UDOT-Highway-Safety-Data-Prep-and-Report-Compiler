###
## Load in Libraries
###

library(tidyverse)
library(dplyr)
library(openxlsx)

###
## Read in temp files
###

IC <- read_csv("data/temp/IC.csv")
IC_byint <- read_csv("data/temp/IC_byint.csv")
RC <- read_csv("data/temp/RC.csv")
RC_byseg <- read_csv("data/temp/RC_byseg.csv")
crash_seg <- read_csv("data/temp/crash_seg.csv")
crash_int <- read_csv("data/temp/crash_int.csv")

###
## Compile Segment Crash & Roadway Data
###

# identify segments for each crash
crash_seg$seg_id <- NA
for (i in 1:nrow(crash_seg)){
  rt <- crash_seg$route[i]
  mp <- crash_seg$milepoint[i]
  seg_row <- which(RC_byseg$ROUTE == rt & 
                     RC_byseg$BEG_MP < mp & 
                     RC_byseg$END_MP > mp)
  if(length(seg_row) > 0){
    crash_seg[["seg_id"]][i] <- RC_byseg$SEG_ID[seg_row]
  }
}

# REPLACE WITH ADDITION METHOD BELOW
# # Add Segment Crashes
# RC$TotalCrashes <- 0
# for (i in 1:nrow(RC)){
#   RCroute <- RC[["ROUTE"]][i]
#   RCbeg <- RC[["BEG_MP"]][i]
#   RCend <- RC[["END_MP"]][i]
#   RCyear <- RC[["YEAR"]][i]
#   crash_row <- which(crash_seg$route == RCroute & 
#                        crash_seg$milepoint > RCbeg & 
#                        crash_seg$milepoint < RCend &
#                        crash_seg$crash_year == RCyear)
#   RC[["TotalCrashes"]][i] <- length(crash_row)
# }

# Add crash Severity
RC <- add_crash_attribute("crash_severity_id", RC, crash_seg)
# Calculate total crashes
RC <- RC %>% 
  group_by(SEG_ID, YEAR) %>%
  mutate(total_crashes = crash_severity_id_1 + crash_severity_id_2 + crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5)
# Add Segment Crash Attributes
RC <- add_crash_attribute("num_veh", RC, crash_seg, TRUE)
RC <- add_crash_attribute("light_condition_id", RC, crash_seg)
RC <- add_crash_attribute("weather_condition_id", RC, crash_seg)
RC <- add_crash_attribute("manner_collision_id", RC, crash_seg)
RC <- add_crash_attribute("horizontal_alignment_id", RC, crash_seg)
RC <- add_crash_attribute("vertical_alignment_id", RC, crash_seg)
RC <- add_crash_attribute("pedalcycle_involved", RC, crash_seg) %>% 
  select(-pedalcycle_involved_N) %>%
  rename(pedalcycle_involved_crashes = pedalcycle_involved_Y)
RC <- add_crash_attribute("pedestrian_involved", RC, crash_seg) %>% 
  select(-pedestrian_involved_N) %>%
  rename(pedestrian_involved_crashes = pedestrian_involved_Y)
RC <- add_crash_attribute("motorcycle_involved", RC, crash_seg) %>% 
  select(-motorcycle_involved_N) %>%
  rename(motorcycle_involved_crashes = motorcycle_involved_Y)
RC <- add_crash_attribute("unrestrained", RC, crash_seg) %>% 
  select(-unrestrained_N) %>%
  rename(unrestrained_crashes = unrestrained_Y)
RC <- add_crash_attribute("dui", RC, crash_seg) %>% 
  select(-dui_N) %>%
  rename(dui_crashes = dui_Y)
RC <- add_crash_attribute("aggressive_driving", RC, crash_seg) %>% 
  select(-aggressive_driving_N) %>%
  rename(aggressive_driving_crashes = aggressive_driving_Y)
RC <- add_crash_attribute("distracted_driving", RC, crash_seg) %>% 
  select(-distracted_driving_N) %>%
  rename(distracted_driving_crashes = distracted_driving_Y)
RC <- add_crash_attribute("drowsy_driving", RC, crash_seg) %>% 
  select(-drowsy_driving_N) %>%
  rename(drowsy_driving_crashes = drowsy_driving_Y)
RC <- add_crash_attribute("speed_related", RC, crash_seg) %>% 
  select(-speed_related_N) %>%
  rename(speed_related_crashes = speed_related_Y)
RC <- add_crash_attribute("adverse_weather", RC, crash_seg) %>% 
  select(-adverse_weather_N) %>%
  rename(adverse_weather_crashes = adverse_weather_Y)
RC <- add_crash_attribute("adverse_roadway_surf_condition", RC, crash_seg) %>% 
  select(-adverse_roadway_surf_condition_N) %>%
  rename(adverse_roadway_surf_condition_crashes = adverse_roadway_surf_condition_Y)
RC <- add_crash_attribute("roadway_geometry_related", RC, crash_seg) %>% 
  select(-roadway_geometry_related_N) %>%
  rename(roadway_geometry_related_crashes = roadway_geometry_related_Y)
RC <- add_crash_attribute("night_dark_condition", RC, crash_seg) %>% 
  select(-night_dark_condition_N) %>%
  rename(night_dark_condition_crashes = night_dark_condition_Y)
RC <- add_crash_attribute("wild_animal_related", RC, crash_seg) %>% 
  select(-wild_animal_related_N) %>%
  rename(wild_animal_related_crashes = wild_animal_related_Y)
RC <- add_crash_attribute("domestic_animal_related", RC, crash_seg) %>% 
  select(-domestic_animal_related_N) %>%
  rename(domestic_animal_related_crashes = domestic_animal_related_Y)
RC <- add_crash_attribute("roadway_departure", RC, crash_seg) %>% 
  select(-roadway_departure_N) %>%
  rename(roadway_departure_crashes = roadway_departure_Y)
RC <- add_crash_attribute("overturn_rollover", RC, crash_seg) %>% 
  select(-overturn_rollover_N) %>%
  rename(overturn_rollover_crashes = overturn_rollover_Y)
RC <- add_crash_attribute("commercial_motor_veh_involved", RC, crash_seg) %>% 
  select(-commercial_motor_veh_involved_N) %>%
  rename(commercial_motor_veh_involved_crashes = commercial_motor_veh_involved_Y)
# RC <- add_crash_attribute("interstate_highway", RC, crash_seg) %>% 
#   select(-interstate_highway_N) %>%
#   rename(interstate_highway_crashes = interstate_highway_Y)
RC <- add_crash_attribute("teen_driver_involved", RC, crash_seg) %>% 
  select(-teen_driver_involved_N) %>%
  rename(teen_driver_involved_crashes = teen_driver_involved_Y)
RC <- add_crash_attribute("older_driver_involved", RC, crash_seg) %>% 
  select(-older_driver_involved_N) %>%
  rename(older_driver_involved_crashes = older_driver_involved_Y)
RC <- add_crash_attribute("single_vehicle", RC, crash_seg) %>% 
  select(-single_vehicle_N) %>%
  rename(single_vehicle_crashes = single_vehicle_Y)
RC <- add_crash_attribute("train_involved", RC, crash_seg) %>% 
  select(-train_involved_N) %>%
  rename(train_involved_crashes = train_involved_Y)
RC <- add_crash_attribute("railroad_crossing", RC, crash_seg) %>% 
  select(-railroad_crossing_N) %>%
  rename(railroad_crossing_crashes = railroad_crossing_Y)
RC <- add_crash_attribute("transit_vehicle_involved", RC, crash_seg) %>% 
  select(-transit_vehicle_involved_N) %>%
  rename(transit_vehicle_involved_crashes = transit_vehicle_involved_Y)
RC <- add_crash_attribute("collision_with_fixed_object", RC, crash_seg) %>% 
  select(-collision_with_fixed_object_N) %>%
  rename(collision_with_fixed_object_crashes = collision_with_fixed_object_Y)

# # Create Parameters File
# crash_seg <- left_join(crash_seg, RC_byseg, by = c("seg_id" = "SEG_ID"))
# 
# # Conform Parameters File to necessary formatting
# param_seg <- crash_seg %>%
#   mutate(
#     ROUTE_ID = as.integer(gsub(".*?([0-9]+).*", "\\1", route)),
#     ramp_id = ifelse(ramp_id == 0, NA, ramp_id),
#     route = substr(route, 1, 5)
#   ) %>%
# select(
#   SEG_ID = seg_id,
#   CRASH_ID = crash_id,
#   CRASH_DATETIME = crash_datetime, #needs reformatting in excel
#   ROUTE_ID, #just the numeric portion
#   DIRECTION = route_direction,
#   LABEL = route, #remove the M
#   RAMP_ID = ramp_id, #convert zeros to blanks
#   MILEPOINT = milepoint,
#   LATITUDE = lat,
#   LONGITUDE = long,
#   (number_vehicles_involved:collision_with_fixed_object),
#   (crash_severity_id:roadway_contrib_circum_id),
#   first_harmful_event_id
#   # NUMBER_VEHICLES_INVOLVED = number_vehicles_involved,
#   # NUMBER_FATALITIES = number_fatalities,
#   # NUMBER_FOUR_INJURIES = number_four_injuries,
#   # NUMBER_THREE_INJURIES = number_three_injuries,
#   # NUMBER_TWO_INJURIES = number_two_injuries,
#   # NUMBER_ONE_INJURIES = number_one_injuries,
#   # PEDESTRIAN_INVOLVED = pedestrian_involved,
#   # BICYCLIST_INVOLVED = pedalcycle_involved,
#   # MOTORCYCLE_INVOLVED = motorcycle_involved,
#   # # IMPROPER_RESTRAINT = , #not in the current dataset
#   # UNRESTRAINED = unrestrained,
#   # DUI = dui,
#   # AGGRESSIVE_DRIVING = aggressive_driving,
#   # DISTRACTED_DRIVING = distracted_driving,
#   # DROWSY_DRIVING = drowsy_driving,
#   # SPEED_RELATED = speed_related,
#   # INTERSECTION_RELATED = intersection_related,
#   # ADVERSE_WEATHER = adverse_weather,
#   # ADVERSE_ROADWAY_SURF_CONDITION = adverse_roadway_surf_condition,
#   # ROADWAY_GEOMETRY_RELATED = roadway_geometry_related,
#   # WILD_ANIMAL_RELATED = wild_animal_related,
#   # DOMESTIC_ANIMAL_RELATED = domestic_animal_related,
#   # ROADWAY_DEPARTURE = roadway_departure,
#   # OVERTURN_ROLLOVER = overturn_rollover,
#   # COMMERCIAL_MOTOR_VEH_INVOLVED = commercial_motor_veh_involved,
#   # INTERSTATE_HIGHWAY = interstate_highway,
#   # TEENAGE_DRIVER_INVOLVED = ,
#   # OLDER_DRIVER_INVOLVED = ,
#   # URBAN_COUNTY = ,
#   # ROUTE_TYPE = ,
#   # NIGHT_DARK_CONDITION = ,
#   # SINGLE_VEHICLE = ,
#   # TRAIN_INVOLVED = ,
#   # RAILROAD_CROSSING = ,
#   # TRANSIT_VEHICLE_INVOLVED = ,
#   # COLLISION_WITH_FIXED_OBJECT = collision_with_fixed_object,
#   # CRASH_SEVERITY_ID = crash_severity_id,
#   # LIGHT_CONDITION_ID = ,
#   # WEATHER_CONDITION_ID = ,
#   # MANNER_COLLISION_ID = ,
#   # # PAVEMENT_ID = , #Not in current dataset
#   # ROADWAY_SURF_CONDITION_ID = ,
#   # ROADWAY_JUNCT_FEATURE_ID = ,
#   # # WORK_ZONE_RELATED_YNU = , #Not in current dataset
#   # # WORK_ZONE_WORKER_PRESENT_YNU = , #Not in current dataset
#   # HORIZONTAL_ALIGNMENT_ID = ,
#   # VERTICAL_ALIGNMENT_ID = ,
#   # ROADWAY_CONTRIB_CIRCUM_ID = roadway_contrib_circum_id,
#   # FIRST_HARMFUL_EVENT_ID = first_harmful_event_id
#   # # VEHICLE_NUM = , #We haven't written code to join vehicle data to crash data
#   # # TRAVEL_DIRECTION_ID = , #We haven't written code to join vehicle data to crash data
#   # # EVENT_SEQUENCE_1_ID = , #We haven't written code to join vehicle data to crash data
#   # # EVENT_SEQUENCE_2_ID = , #We haven't written code to join vehicle data to crash data
#   # # EVENT_SEQUENCE_3_ID = , #We haven't written code to join vehicle data to crash data
#   # # EVENT_SEQUENCE_4_ID = , #We haven't written code to join vehicle data to crash data
#   # # MOST_HARMFUL_EVENT_ID = , #We haven't written code to join vehicle data to crash data
#   # # VEHICLE_MANEUVER_ID = , #We haven't written code to join vehicle data to crash data
#   # # VEHICLE_DETAIL_ID =  #We haven't written code to join vehicle data to crash data
# ) %>%
#   arrange(LABEL, MILEPOINT)

# Last Minute Changes to the Data

# WARNING: MAKE SURE THIS DOESN'T REMOVE IMPORTANT INFORMATION
# Remove rows where AADT is NA because we are assuming the route didn't exist during that year
RC <- RC %>% filter(!is.na(AADT))
RC <- RC %>% rename(Driveway_Density_dpm = Driveway_Freq) # already fixed in roadwayprep. just renaming for clarity.


###
## Compile Intersection Crash & Roadway Data
###

# Add crash severity
IC <- add_crash_attribute_int("crash_severity_id", IC, crash_int)
# Calculate total crashes
IC <- IC %>% 
  group_by(Int_ID, YEAR) %>%
  mutate(total_crashes = crash_severity_id_1 + crash_severity_id_2 + crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5)
# Add Intersection Crash Attributes
IC <- add_crash_attribute_int("num_veh", IC, crash_int, TRUE)
IC <- add_crash_attribute_int("light_condition_id", IC, crash_int)
IC <- add_crash_attribute_int("weather_condition_id", IC, crash_int)
IC <- add_crash_attribute_int("manner_collision_id", IC, crash_int)
IC <- add_crash_attribute_int("horizontal_alignment_id", IC, crash_int)
IC <- add_crash_attribute_int("vertical_alignment_id", IC, crash_int)
IC <- add_crash_attribute_int("pedalcycle_involved", IC, crash_int) %>% 
  select(-pedalcycle_involved_N) %>%
  rename(pedalcycle_involved_crashes = pedalcycle_involved_Y)
IC <- add_crash_attribute_int("pedestrian_involved", IC, crash_int) %>% 
  select(-pedestrian_involved_N) %>%
  rename(pedestrian_involved_crashes = pedestrian_involved_Y)
IC <- add_crash_attribute_int("motorcycle_involved", IC, crash_int) %>% 
  select(-motorcycle_involved_N) %>%
  rename(motorcycle_involved_crashes = motorcycle_involved_Y)
IC <- add_crash_attribute_int("unrestrained", IC, crash_int) %>% 
  select(-unrestrained_N) %>%
  rename(unrestrained_crashes = unrestrained_Y)
IC <- add_crash_attribute_int("dui", IC, crash_int) %>% 
  select(-dui_N) %>%
  rename(dui_crashes = dui_Y)
IC <- add_crash_attribute_int("aggressive_driving", IC, crash_int) %>% 
  select(-aggressive_driving_N) %>%
  rename(aggressive_driving_crashes = aggressive_driving_Y)
IC <- add_crash_attribute_int("distracted_driving", IC, crash_int) %>% 
  select(-distracted_driving_N) %>%
  rename(distracted_driving_crashes = distracted_driving_Y)
IC <- add_crash_attribute_int("drowsy_driving", IC, crash_int) %>% 
  select(-drowsy_driving_N) %>%
  rename(drowsy_driving_crashes = drowsy_driving_Y)
IC <- add_crash_attribute_int("speed_related", IC, crash_int) %>% 
  select(-speed_related_N) %>%
  rename(speed_related_crashes = speed_related_Y)
IC <- add_crash_attribute_int("adverse_weather", IC, crash_int) %>% 
  select(-adverse_weather_N) %>%
  rename(adverse_weather_crashes = adverse_weather_Y)
IC <- add_crash_attribute_int("adverse_roadway_surf_condition", IC, crash_int) %>% 
  select(-adverse_roadway_surf_condition_N) %>%
  rename(adverse_roadway_surf_condition_crashes = adverse_roadway_surf_condition_Y)
IC <- add_crash_attribute_int("roadway_geometry_related", IC, crash_int) %>% 
  select(-roadway_geometry_related_N) %>%
  rename(roadway_geometry_related_crashes = roadway_geometry_related_Y)
IC <- add_crash_attribute_int("night_dark_condition", IC, crash_int) %>% 
  select(-night_dark_condition_N) %>%
  rename(night_dark_condition_crashes = night_dark_condition_Y)
IC <- add_crash_attribute_int("wild_animal_related", IC, crash_int) %>% 
  select(-wild_animal_related_N) %>%
  rename(wild_animal_related_crashes = wild_animal_related_Y)
IC <- add_crash_attribute_int("domestic_animal_related", IC, crash_int) %>% 
  select(-domestic_animal_related_N) %>%
  rename(domestic_animal_related_crashes = domestic_animal_related_Y)
IC <- add_crash_attribute_int("roadway_departure", IC, crash_int) %>% 
  select(-roadway_departure_N) %>%
  rename(roadway_departure_crashes = roadway_departure_Y)
IC <- add_crash_attribute_int("overturn_rollover", IC, crash_int) %>% 
  select(-overturn_rollover_N) %>%
  rename(overturn_rollover_crashes = overturn_rollover_Y)
IC <- add_crash_attribute_int("commercial_motor_veh_involved", IC, crash_int) %>% 
  select(-commercial_motor_veh_involved_N) %>%
  rename(commercial_motor_veh_involved_crashes = commercial_motor_veh_involved_Y)
# IC <- add_crash_attribute_int("interstate_highway", IC, crash_int) %>% 
#   select(-interstate_highway_N) %>%
#   rename(interstate_highway_crashes = interstate_highway_Y)
IC <- add_crash_attribute_int("teen_driver_involved", IC, crash_int) %>% 
  select(-teen_driver_involved_N) %>%
  rename(teen_driver_involved_crashes = teen_driver_involved_Y)
IC <- add_crash_attribute_int("older_driver_involved", IC, crash_int) %>% 
  select(-older_driver_involved_N) %>%
  rename(older_driver_involved_crashes = older_driver_involved_Y)
IC <- add_crash_attribute_int("single_vehicle", IC, crash_int) %>% 
  select(-single_vehicle_N) %>%
  rename(single_vehicle_crashes = single_vehicle_Y)
IC <- add_crash_attribute_int("train_involved", IC, crash_int) %>% 
  select(-train_involved_N) %>%
  rename(train_involved_crashes = train_involved_Y)
IC <- add_crash_attribute_int("railroad_crossing", IC, crash_int) %>% 
  select(-railroad_crossing_N) %>%
  rename(railroad_crossing_crashes = railroad_crossing_Y)
IC <- add_crash_attribute_int("transit_vehicle_involved", IC, crash_int) %>% 
  select(-transit_vehicle_involved_N) %>%
  rename(transit_vehicle_involved_crashes = transit_vehicle_involved_Y)
IC <- add_crash_attribute_int("collision_with_fixed_object", IC, crash_int) %>% 
  select(-collision_with_fixed_object_N) %>%
  rename(collision_with_fixed_object_crashes = collision_with_fixed_object_Y)

# # Create Parameters File
# crash_int <- left_join(crash_int, IC_byint, by = c("int_id" = "Int_ID"))
# 
# # Conform Parameters File to necessary formatting
# param_int <- crash_int %>%
#   mutate(
#     ROUTE_ID = as.integer(gsub(".*?([0-9]+).*", "\\1", route)),
#     ramp_id = ifelse(ramp_id == 0, NA, ramp_id),
#     route = substr(route, 1, 5)
#   ) %>%
#   select(
#     INT_ID = int_id,
#     CRASH_ID = crash_id,
#     CRASH_DATETIME = crash_datetime, #needs reformatting in excel
#     (number_vehicles_involved:collision_with_fixed_object),
#     (crash_severity_id:roadway_contrib_circum_id),
#     first_harmful_event_id,
#     ROUTE_ID, #just the numeric portion
#     LABEL = route, #remove the M
#     DIRECTION = route_direction,
#     RAMP_ID = ramp_id, #convert zeros to blanks
#     MILEPOINT = milepoint,
#     LATITUDE = lat.x,
#     LONGITUDE = long.x
#   ) %>%
#   arrange(LABEL, MILEPOINT)

# Last Minute Changes to the Data

# WARNING: MAKE SURE THIS DOESN'T REMOVE IMPORTANT INFORMATION
# # Remove rows where AADT is zero because we are assuming the route didn't exist during that year
# IC <- IC %>% filter(!is.na(ENT_VEH))


###
## Write to output
###

CAMSoutput <- paste0("data/output/CAMS_",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(RC, file = CAMSoutput)
write_csv(RC, "data/temp/CAMS.csv")
ISAMoutput <- paste0("data/output/ISAM_",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(IC, file = ISAMoutput)
write_csv(IC, "data/temp/ISAM.csv")

# # Paste header details into parameter files
# latest_year <- max(IC$YEAR) # finds latest year
# earliest_year <- min(IC$YEAR) # finds earliest year
# param_header_seg <- tibble(c("Severities", "Functional Area Definition", "Selected Years:"),
#                            c(12345, "UDOT", paste0(earliest_year,"-",latest_year)))
# param_header_int <- tibble(c("Severities", "Functional Area Type", "Selected Years:"),
#                            c(12345, "UDOT", paste0(earliest_year,"-",latest_year)))
# 
# # Determine filepath for parameter files
# CAMSparameters <- paste0("data/output/CAMS_parameters_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")
# ISAMparameters <- paste0("data/output/ISAM_parameters_",format(Sys.time(),"%d%b%y_%H_%M"),".xlsx")
# 
# # Save parameters workbooks
# wb <- createWorkbook()
# addWorksheet(wb, sheetName = "Parameters")
# 
# writeData(wb, sheet = 1, x = param_header_seg, startRow = 1, colNames = FALSE, rowNames = FALSE)
# writeData(wb, sheet = 1, x = param_seg, startRow = 4, colNames = TRUE)
# 
# saveWorkbook(wb, file = CAMSparameters)
# 
# wb <- createWorkbook()
# addWorksheet(wb, sheetName = "Parameters")
# 
# writeData(wb, sheet = 1, x = param_header_int, startRow = 1, colNames = FALSE, rowNames = FALSE)
# writeData(wb, sheet = 1, x = param_int, startRow = 4, colNames = TRUE)
# 
# saveWorkbook(wb, file = ISAMparameters)

# # # reset RC and IC
# RC <- RC %>% select(SEG_ID:CUTRK)
# IC <- IC %>% select(ROUTE:CUTRK)
