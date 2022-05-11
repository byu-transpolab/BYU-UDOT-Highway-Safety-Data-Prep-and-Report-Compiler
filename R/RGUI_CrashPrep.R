# The following is the functions that prepare the crash data to be joined

# Keeps the following columns from the location file
check_location <- function(df){
  df %>%
    select(crash_id = crash_id, 
           crash_datetime = crash_datetime, 
           route = route, 
           direction = route_direction, 
           ramp_id = ramp_id, 
           milepoint = milepoint, 
           lat = lat, 
           long = long, 
           crash_severity_id = crash_severity_id, 
           light_condition_id = light_condition_id, 
           weather_condition_id = weather_condition_id, 
           manner_collision_id = manner_collision_id,
           roadway_surf_condition_id = roadway_surf_condition_id,
           roadway_junct_feature_id = roadway_junct_feature_id, 
           horizontal_alignment_id = horizontal_alignment_id,
           vertical_alignment_id = vertical_alignment_id, 
           roadway_contrib_circum_id = roadway_contrib_circum_id,
           first_harmful_event_id = first_harmful_event_id)
}

# Keeps the following columns from the vehicle file
check_vehicle <- function(df){
  df %>%
    select(crash_id = crash_id, 
           vehicle_num = vehicle_num, 
           travel_direction_id = travel_direction_id, 
           event_1_id = event_sequence_1_id, 
           event_2_id = event_sequence_2_id, 
           event_3_id = event_sequence_3_id, 
           event_4_id = event_sequence_4_id,
           most_harmful_event_id = most_harmful_event_id, 
           vehicle_maneuver_id = vehicle_maneuver_id)
}

# Keeps the following columns from the rollups file
check_rollups <- function(df){
  df %>%
    select(crash_id = crash_id, 
           number_fatalities = number_fatalities, 
           number_four_injuries = number_four_injuries,
           number_three_injuries = number_three_injuries, 
           number_two_injuries = number_two_injuries,
           number_one_injuries = number_one_injuries,
           pedalcycle_involved = pedalcycle_involved, 
           pedalcycle_involved_level4_tot = pedalcycle_involved_level4_tot,	
           pedalcycle_involved_fatal_tot = pedalcycle_involved_fatal_tot,
           pedestrian_involved = pedestrian_involved,	
           pedestrian_involved_level4_tot = pedestrian_involved_level4_tot,	
           pedestrian_involved_fatal_tot = pedestrian_involved_fatal_tot,
           motorcycle_involved = motorcycle_involved,	
           motorcycle_involved_level4_tot = motorcycle_involved_level4_tot,	
           motorcycle_involved_fatal_tot = motorcycle_involved_fatal_tot,
           unrestrained = unrestrained, 
           dui = dui, 
           aggressive_driving = aggressive_driving,	
           distracted_driving = distracted_driving,	
           drowsy_driving = drowsy_driving,	
           speed_related = speed_related,
           intersection_related = intersection_related,	
           adverse_weather = adverse_weather,	
           adverse_roadway_surf_condition = adverse_roadway_surf_condition,	
           roadway_geometry_related = roadway_geometry_related,
           wild_animal_related = wild_animal_related,	
           domestic_animal_related = domestic_animal_related,	
           roadway_departure = roadway_departure,	
           overturn_rollover = overturn_rollover,	
           commercial_motor_veh_involved = commercial_motor_veh_involved,	
           interstate_highway = interstate_highway,	
           teen_driver_involved = teen_driver_involved,	
           older_driver_involved = older_driver_involved,	
           route_type = route_type,	
           night_dark_condition = night_dark_condition,	
           single_vehicle = single_vehicle,	
           train_involved = train_involved,	
           railroad_crossing = railroad_crossing,
           transit_vehicle_involved = transit_vehicle_involved,	
           collision_with_fixed_object = collision_with_fixed_object)
}

location <- read.csv("data/csv/Crash_Data_14-20.csv")
location_fp <- "data/csv/Crash_Data_14-20.csv"
location_col <- c("crash_id", 
                  "crash_datetime", 
                  "route", 
                  "direction", 
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

rollups <- read.csv("data/csv/Rollups_14-20.csv")
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
crash <- left_join(crash,vehicle,by='crash_id')

# FILTER CAMS CRASH DATA
filter_CAMS <- function(df){
  # change route 089A to 0011... There seems to be more to this though
  df$ROUTE_ID[df$ROUTE_ID == '089A'] <- 0011
  
  # change route ID to 194 from 0085 for milepoint less than 2.977 
  df$ROUTE_ID[df$MILEPOINT < 2.977 & df$ROUTE_ID == 0085] <- 0194
  
  # eliminate rows with ramps
  df %>%
    as_tibble() %>%
    filter(RAMP_ID == 0 & is.na(RAMP_ID))
}
