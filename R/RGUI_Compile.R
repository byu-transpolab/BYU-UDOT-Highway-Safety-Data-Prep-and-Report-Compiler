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
    NUM_LEGS = as.integer(gsub(".*?([0-9]+).*", "\\1", INT_TYPE))
  ) %>%
  mutate(
    NUM_LEGS = case_when(
      INT_TYPE == "DDI" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "ROUNDABOUT" |
      INT_TYPE == "CFI CENTRAL" |
      INT_TYPE == "SPUI" |
      INT_TYPE == "THRU TURN CENTRAL" |
      Int_ID == "0173P-7.167-0173" |
      Int_ID == "0173P-7.369-0173" |
      Int_ID == "0265P-0.65-0265" |
      Int_ID == "0265P-0.823-0256" |
      Int_ID == "0071P-5.207-0071" |
      Int_ID == "0126P-1.803-0126"
      ~ 4L,
      Int_ID == "0154P-5.602-None" |
      Int_ID == "0154P-5.884-0154" |
      Int_ID == "0154P-14.756-None" |
      Int_ID == "0154P-15.072-0154" |
      Int_ID == "0154P-15.801-None" |
      Int_ID == "0154P-16.822-None" |
      Int_ID == "0154P-17.054-0154" |
      Int_ID == "0154P-17.817-None" |
      Int_ID == "0154P-18.062-0154" |
      Int_ID == "0154P-19.024-0154" |
      Int_ID == "0154P-19.31-None" |
      Int_ID == "0154P-19.6-0154" |
      Int_ID == "0232P-0.39-None" |
      Int_ID == "0266P-7.665-None" |
      Int_ID == "0126P-1.61-0126"
      ~ 3L,
      Int_ID == "0172P-0.456-None" |
      Int_ID == "0172P-2.775-None" |
      Int_ID == "0289P-0.458-None" |
      Int_ID == "0089P-366.513-None" |
      Int_ID == "0089P-380.099-None" |
      Int_ID == "0089P-362.667-None" |
      Int_ID == "0126P-1.797-None" |
      Int_ID == "0154P-18.863-None"
      ~ 2L,
      TRUE ~ NUM_LEGS
    )
  )

# Add roadway data
IC <- add_int_att(IC, urban) %>% 
  select(-(MIN_URBAN_CODE:URBAN_CODE_4)) %>%
  rename(URBAN_CODE = MAX_URBAN_CODE)
IC <- add_int_att(IC, fc %>% select(-RouteDir,-RouteType)) %>% 
  select(-(MAX_FUNCTIONAL_CLASS:AVG_FUNCTIONAL_CLASS))
IC <- add_int_att(IC, aadt, TRUE)
  # select(-contains("MAX_AADT"),-contains("MIN_AADT"),-contains("AVG_AADT"),
  #        -contains("MAX_SUTRK"),-contains("MIN_SUTRK"),-contains("AVG_SUTRK"),
  #        -contains("MAX_CUTRK"),-contains("MIN_CUTRK"),-contains("AVG_CUTRK"))
IC <- add_int_att(IC, lane) %>%
  select(-(AVG_THRU_CNT:THRU_CNT_4),-(AVG_THRU_WDTH:THRU_WDTH_4))

# Add years to intersection shell
# yrs <- crash_int %>% select(crash_year) %>% unique()
# IC <- IC %>%
#   group_by(Int_ID) %>%
#   slice(rep(row_number(), times = nrow(yrs))) %>%
#   mutate(YEAR = min(yrs$crash_year):max(yrs$crash_year))
IC <- pivot_aadt(IC) %<% rename(ENT_VEH = AADT)

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

CAMSoutput <- paste0("data/output/CAMS_",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(RC, file = CAMSoutput)
ISAMoutput <- paste0("data/output/ISAM_",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(IC, file = ISAMoutput)
