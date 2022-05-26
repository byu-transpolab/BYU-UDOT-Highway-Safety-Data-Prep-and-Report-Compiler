###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Set Filepath and Column Names for each Dataset
###

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

location <- read_csv_file(location_fp, location_col)
rollups <- read_csv_file(rollups_fp, rollups_col)
vehicle <- read_csv_file(vehicle_fp, vehicle_col)

crash <- left_join(location,rollups,by='crash_id')
fullcrash <- left_join(crash,vehicle,by='crash_id')

crashfilter <- function(crash){
  # Filter Out all Ramp Crashes
  crash <- crash %>% filter(is.na(ramp_id))
  # Split Date Time Column
  crash$crash_date <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 1)
  crash$crash_time <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 2)
  crash$crash_year <- sapply(strsplit(as.character(crash$crash_date), "/"), "[", 3)
  # Format Routes to match other Datasets
  crash$route <- paste(substr(crash$route, 1, 5), crash$route_direction, sep = "")
  crash$route <- paste(substr(crash$route, 1, 6), "M", sep = "")
  crash$route <- paste0("000", crash$route)
  crash$route <- substr(crash$route, nchar(crash$route)-6+1, nchar(crash$route))
  crash <- crash %>% filter(route %in% substr(main.routes, 1, 6))
}
