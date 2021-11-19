# The following is the functions that prepare the roadway data to be joined

# Keeps the following columns from the aadt file
check_aadt <- function(df){
  df %>%
    select(rt_num, 
           start_accu,
           end_accum, 
           station, 
           starts_with('aadt'),
           #aadt2020, aadt2019, aadt2018, aadt2017, aadt2016,
           starts_with('sutrk'),
           starts_with('cutrk'),
           #sutrk2020, cutrk2020, sutrk2019, cutrk2019, sutrk2018, cutrk2018, sutrk2017, cutrk2017, sutrk2016, cutrk2016,
           rt_type)
}

# Keeps the following columns from the functional class file
check_functionalclass <- function(df){
  df %>%
    select(from_measure, 
           to_measure, 
           functional_class, 
           )
}

# Keeps the following columns from the intersections file
check_intersections <- function(df){
  df %>%
    select(orig_int_id = ID,
           lat = BEG_LAT,
           long = BEG_LONG,
           elevation = BEG_ELEV,
           traffic_ctrl = TRAFFIC_CO,
           sr_sr = SR_SR,
           route = ROUTE,
           int_rt_1 = INT_RT_1,
           int_rt_2 = INT_RT_2,
           int_rt_3 = INT_RT_3,
           int_rt_4 = INT_RT_4,
           udot_bmp = MANDLI_BMP,
           int_type = INT_TYPE,
           station = STATION,
           region = REGION,
           #group_num = GROUP_NUM  -can't find this one
           num_schools = NUM_SCHOOLS,
           num_uta = NUM_UTA)
}

# Keeps the following columns from the lanes file
check_lanes <- function(df){
  df %>%
    select(route = ROUTE,
           beg_milepoint = START_ACCUM,
           end_milepoint = END_ACCUM,
           
           )
}

# Keeps the following columns from the pavement messages file
check_pavementmessages <- function(df){
  df %>%
    select()
}

# Keeps the following columns from the speed limits file
check_speedlimits <- function(df){
  df %>%
    select()
}


# Modify intersections to include bus stops and schools near intersections
mod_intersections <- function(intersections,UTA_Stops,schools){
  intersections <- st_join(intersections, schools, join = st_within) %>%
    group_by(OBJECTID) %>%
    mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID)
  
  intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
    group_by(OBJECTID) %>%
    mutate(NUM_UTA = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
    select(-UTA_StopID)
  
  st_drop_geometry(intersections)
}