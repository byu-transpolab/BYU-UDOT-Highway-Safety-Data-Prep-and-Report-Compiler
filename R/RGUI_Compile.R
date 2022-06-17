###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

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

# Add Segment Crash Attributes
RC <- add_crash_attribute("crash_severity_id", RC, crash_seg)
RC <- add_crash_attribute("light_condition_id", RC, crash_seg)
RC <- add_crash_attribute("weather_condition_id", RC, crash_seg)
RC <- add_crash_attribute("light_condition_id", RC, crash_seg)
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
RC <- add_crash_attribute("speed_related", RC, crash_seg) %>% 
  select(-speed_related_N) %>%
  rename(speed_related_crashes = speed_related_Y)
RC <- add_crash_attribute("adverse_roadway_surf_condition", RC, crash_seg) %>% 
  select(-adverse_roadway_surf_condition_N) %>%
  rename(adverse_roadway_surf_condition_crashes = adverse_roadway_surf_condition_Y)
RC <- add_crash_attribute("roadway_geometry_related", RC, crash_seg) %>% 
  select(-roadway_geometry_related_N) %>%
  rename(roadway_geometry_related_crashes = roadway_geometry_related_Y)
RC <- add_crash_attribute("wild_animal_related", RC, crash_seg) %>% 
  select(-wild_animal_related_N) %>%
  rename(wild_animal_related_crashes = wild_animal_related_Y)
RC <- add_crash_attribute("roadway_departure", RC, crash_seg) %>% 
  select(-roadway_departure_N) %>%
  rename(roadway_departure_crashes = roadway_departure_Y)
RC <- add_crash_attribute("overturn_rollover", RC, crash_seg) %>% 
  select(-overturn_rollover_N) %>%
  rename(overturn_rollover_crashes = overturn_rollover_Y)

# Calculate total crashes
RC <- RC %>% 
  group_by(SEG_ID, YEAR) %>%
  mutate(TotalCrashes = crash_severity_id_1 + crash_severity_id_2 + crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5)


###
## Compile Intersection Crash & Roadway Data
###

# Create shell for intersection file
IC <- intersection %>% 
  rowwise() %>%
  mutate(
    MP = BEG_MP,
    MAX_SPEED_LIMIT = max(c_across(INT_RT_1_SL:INT_RT_4_SL), na.rm = TRUE),
    MIN_SPEED_LIMIT = min(c_across(INT_RT_1_SL:INT_RT_4_SL), na.rm = TRUE)
  ) %>%
  select(ROUTE, MP, Int_ID, everything()) %>%
  select(-BEG_MP, -END_MP, -contains("_SL"), -contains("_FA"))

# append "PM" or "NM" to routes
for(i in 1:nrow(IC)){
  id <- IC[["Int_ID"]][i]
  for(j in 1:4){
    rt <- IC[[paste0("INT_RT_",j)]][i]
    if(rt %in% substr(main.routes,1,4)){
      if(rt == substr(IC$ROUTE[i],0,4)){
        IC[[paste0("INT_RT_",j)]][i] <- IC$ROUTE[i]
      } else{
        IC[[paste0("INT_RT_",j)]][i] <- paste0(rt,"PM")   # assuming if direction not specified it is positive
      }
    }
  }
}

# Create a Num_Legs column
IC <- IC %>%
  mutate(
    NUM_LEGS = as.integer(gsub(".*?([0-9]+).*", "\\1", INT_TYPE)),
    NUM_LEGS = case_when(
      INT_TYPE == "DDI" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "CFI OFFSET LEFT TURN" |
      INT_TYPE == "ROUNDABOUT" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "SPUI" |
      INT_TYPE == "THRU TURN CENTRAL" |
      INT_TYPE == "THRU TURN OFFSET U-TURN"
      ~ 4,
      INT_TYPE == "DDI" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "CFI OFFSET LEFT TURN" |
      INT_TYPE == "ROUNDABOUT" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "SPUI" |
      INT_TYPE == "THRU TURN CENTRAL" |
      INT_TYPE == "THRU TURN OFFSET U-TURN"
      ~ 3,
      INT_TYPE == "DDI" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "CFI OFFSET LEFT TURN" |
      INT_TYPE == "ROUNDABOUT" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "SPUI" |
      INT_TYPE == "THRU TURN CENTRAL" |
      INT_TYPE == "THRU TURN OFFSET U-TURN"
      ~ 2
    )
  )

# Add roadway data
IC <- add_int_att(IC, urban) %>% 
  select(-(MIN_URBAN_CODE:URBAN_CODE_4)) %>%
  rename(URBAN_CODE = MAX_URBAN_CODE)
IC <- add_int_att(IC, fc %>% select(-RouteDir,-RouteType)) %>% 
  select(-(MAX_FUNCTIONAL_CLASS:AVG_FUNCTIONAL_CLASS))

IC <- add_int_att(IC, aadt)  
IC <- IC %>%
  select(-contains("MAX_AADT"),-contains("MIN_AADT"),-contains("AVG_AADT"),
         -contains("MAX_SUTRK"),-contains("MIN_SUTRK"),-contains("AVG_SUTRK"),
         -contains("MAX_CUTRK"),-contains("MIN_CUTRK"),-contains("AVG_CUTRK")) %>%
  

IC <- add_int_att(IC, lane) %>%
  select(-(AVG_THRU_CNT:THRU_CNT_4),-(AVG_THRU_WDTH:THRU_WDTH_4))

# Add years to intersection shell
yrs <- crash_int %>% select(crash_year) %>% unique()
IC <- IC %>%
  group_by(Int_ID) %>%
  slice(rep(row_number(), times = nrow(yrs))) %>%
  mutate(YEAR = min(yrs$crash_year):max(yrs$crash_year))

# Add Intersection Crash Attributes
IC <- add_crash_attribute_int("crash_severity_id", IC, crash_int)
IC <- add_crash_attribute_int("light_condition_id", IC, crash_int)
IC <- add_crash_attribute_int("weather_condition_id", IC, crash_int)
IC <- add_crash_attribute_int("light_condition_id", IC, crash_int)
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
IC <- add_crash_attribute_int("speed_related", IC, crash_int) %>% 
  select(-speed_related_N) %>%
  rename(speed_related_crashes = speed_related_Y)
IC <- add_crash_attribute_int("adverse_roadway_surf_condition", IC, crash_int) %>% 
  select(-adverse_roadway_surf_condition_N) %>%
  rename(adverse_roadway_surf_condition_crashes = adverse_roadway_surf_condition_Y)
IC <- add_crash_attribute_int("roadway_geometry_related", IC, crash_int) %>% 
  select(-roadway_geometry_related_N) %>%
  rename(roadway_geometry_related_crashes = roadway_geometry_related_Y)
IC <- add_crash_attribute_int("wild_animal_related", IC, crash_int) %>% 
  select(-wild_animal_related_N) %>%
  rename(wild_animal_related_crashes = wild_animal_related_Y)
IC <- add_crash_attribute_int("roadway_departure", IC, crash_int) %>% 
  select(-roadway_departure_N) %>%
  rename(roadway_departure_crashes = roadway_departure_Y)
IC <- add_crash_attribute_int("overturn_rollover", IC, crash_int) %>% 
  select(-overturn_rollover_N) %>%
  rename(overturn_rollover_crashes = overturn_rollover_Y)

# Calculate total crashes
IC <- IC %>% 
  group_by(Int_ID, YEAR) %>%
  mutate(TotalCrashes = crash_severity_id_1 + crash_severity_id_2 + crash_severity_id_3 + crash_severity_id_4 + crash_severity_id_5)


###
## Write to output
###

output <- paste0("data/output/",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(RC, file = paste0("CAMS",output))
write_csv(IC, file = paste0("ISAM",output))