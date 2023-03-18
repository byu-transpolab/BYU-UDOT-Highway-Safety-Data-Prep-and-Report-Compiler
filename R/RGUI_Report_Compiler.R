library(tidyverse)
library(openxlsx)


# Read in Statistical Files
TOMS_seg <- read_csv("data/csv/Seg_Out_WRS.csv")
TOMS_int <- read_csv("data/csv/Int_Out_WRS.csv")
CAMS <- read_csv("data/temp/CAMS.csv")
ISAM <- read_csv("data/temp/ISAM.csv")


# Join data to output files
TOMS_seg <- left_join(TOMS_seg, CAMS, by = "SEG_ID")
TOMS_int <- left_join(TOMS_int, ISAM, by = c("INT_ID" = "Int_ID"))


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
    Percentile = percent_rank(`Ranking w/WRS`),
    Region_Rank = NA, # these require grouping so we'll do it later
    County_Rank = NA 
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
    (light_condition_id_1:collision_with_fixed_object_crashes),
    VMT,
    logVMT,
    hier,
    !!sym(paste0("Crashes",substr(latest_year,3,4))) := total_crashes,
    Predicted_Total = final_cost,
    Percentile,
    State_Rank = `Ranking w/WRS`,
    Region_Rank,
    County_Rank
  )

# Format intersections results
latest_year <- max(TOMS_int$YEAR, na.rm = TRUE) # finds latest year
TOMS_int <- TOMS_int %>%
  mutate(
    ORIGINAL_INT_ID = NA, # not included in our latest files
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
    Percentile = percent_rank(`Ranking w/WRS`),
    cover = NA, # don't know what this is
    Region_Rank = NA, # these require grouping so we'll do it later
    County_Rank = NA 
  ) %>%
  select(
    INT_ID,
    ORIGINAL_INT_ID,
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
    # MIN.FC_TYPE = ,
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
    (light_condition_id_6:collision_with_fixed_object_crashes),
    hier,
    !!sym(paste0("Crashes",substr(latest_year,3,4))) := total_crashes,
    Predicted_Total = final_cost,
    Percentile,
    cover,
    State_Rank = `Ranking w/WRS`,
    Region_Rank,
    County_Rank
  )


# Reduce Data
TOMS_seg <- TOMS_seg %>%
  group_by(SEG_ID) %>%
  mutate(across(Total_Crashes:collision_with_fixed_object_crashes, sum)) %>%
  ungroup() %>%
  filter(YEAR == latest_year)

TOMS_int <- TOMS_int %>%
  group_by(INT_ID) %>%
  mutate(across(Total_Crashes:collision_with_fixed_object_crashes, sum)) %>%
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


