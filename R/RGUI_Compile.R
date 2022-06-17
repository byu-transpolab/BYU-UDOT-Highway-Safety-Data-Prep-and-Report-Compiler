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

# Add roadway data
int <- IC
att <- urban
add_int_att <- function(int, att){
  # create row id column for attribute data
  att <- att %>% rownames_to_column("att_rw")
  # identify attribute segments on intersection file
  for(k in 1:4){               # since we can expect intersections to have overlapping attributes, we need to identify all of these
    int[[paste0("att_rw_",k)]] <- NA
  }
  for (i in 1:nrow(int)){
    rt1 <- int$INT_RT_1[i]
    mp1 <- int$INT_RT_1_MP[i]
    rt2 <- int$INT_RT_2[i]
    mp2 <- int$INT_RT_2_MP[i]
    rt3 <- int$INT_RT_3[i]
    mp3 <- int$INT_RT_3_MP[i]
    rt4 <- int$INT_RT_4[i]
    mp4 <- int$INT_RT_4_MP[i]
    if(!is.na(rt1)){
      rw1 <- which(att$ROUTE == rt1 & 
                     att$BEG_MP < mp1 & 
                     att$END_MP > mp1)
    }
    if(!is.na(rt2)){
      rw2 <- which(att$ROUTE == rt2 & 
                     att$BEG_MP < mp2 & 
                     att$END_MP > mp2)
    }
    if(!is.na(rt3)){
      rw3 <- which(att$ROUTE == rt3 & 
                     att$BEG_MP < mp3 & 
                     att$END_MP > mp3)
    }
    if(!is.na(rt4)){
      rw4 <- which(att$ROUTE == rt4 & 
                     att$BEG_MP < mp4 & 
                     att$END_MP > mp4)
    }
    if(length(rw1) > 0 | length(rw2) > 0 | length(rw3) > 0 | length(rw4) > 0){
      j <- 0
      for(n in 1:4){
        rw <- sym(paste0("rw",n))
        for(m in 1:length(rw)){
          j <- j + 1
          int[[paste0("att_rw_",j)]][i] <- rw[m]
        }
      }
    }
  }
  # join to segments
  ints <- left_join(int, att, by = c("att_rw"="att_rw"))
  # remove att_rw and return segments
  int <- int %>% select(-contains("att_rw"))
  return(int)
}

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