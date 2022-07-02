###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Roadway Data Prep
###

# List Data needed for Aggregation

sdtms <- list(aadt, fc, speed, lane, urban)
sdtms <- lapply(sdtms, as_tibble)

# Generating Merge Cell

shell <- shell_creator(sdtms)

joined_populated <- lapply(sdtms, shell_join)

###
## Merging joined and populated sdtms together; formatting
###

RC <- c(list(shell), joined_populated) %>% 
  reduce(left_join, by = c("ROUTE", "startpoints")) %>% 
  mutate(BEG_MP = startpoints,
         END_MP = endpoints) %>% 
  select(-startpoints, -endpoints) %>% 
  select(ROUTE, BEG_MP, END_MP, everything()) %>% 
  arrange(ROUTE, BEG_MP)

###
## Final Data Compilation
###

# Compressing by combining rows identical except for identifying information 

RC <- compress_seg_alt(RC)
RC <- RC %>% unique()
RC <- compress_seg_ovr(RC)


###
## Fix Spatial Gaps
###

# Use fc file to build a gaps dataset
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

# Delete segments which fall within gaps
RC <- RC %>% mutate(MID_MP = BEG_MP + (END_MP - BEG_MP) / 2)

RC$flag <- 0
for(i in 1:nrow(RC)){
  for(j in 1:nrow(gaps)){
    if(RC$MID_MP[i] > gaps$BEG_GAP[j] & 
       RC$MID_MP[i] < gaps$END_GAP[j] &
       RC$BEG_MP[i] >= gaps$BEG_GAP[j] &
       RC$END_MP[i] <= gaps$END_GAP[j] &
       RC$ROUTE[i] == gaps$ROUTE[j]){
      RC$flag[i] <- 1
    }
  }
}

RC <- RC %>% filter(flag == 0) %>% select(-flag, -MID_MP)

# Compressing segments by length
test <- compress_seg_len(RC, 0.1)

for(i in 2:nrow(test)){
  if(test$BEG_MP[i] > test$END_MP[i-1] &
     test$ROUTE[i] == test$ROUTE[i-1]){
    chk <- 0
    for(j in 1:nrow(gaps)){
      if(test$ROUTE[i] == gaps$ROUTE[j] &
         round(test$BEG_MP[i],2) == round(gaps$END_GAP[j],2)){
        chk <- 1
      }
    }
    if(chk == 0){
      print(paste("unwanted gap at",test$ROUTE[i],test$BEG_MP[i]))
    }
  }
  if(test$BEG_MP[i] < test$END_MP[i-1]){
    print(paste("overlap at",test$ROUTE[i],test$BEG_MP[i]))
  }
}


###
## Adding Other Data
###

# Add Driveways
RC$Driveway_Freq <- 0
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  drive_row <- which(driveway$ROUTE == RCroute & 
                       driveway$MP > RCbeg & 
                       driveway$MP < RCend)
  RC[["Driveway_Freq"]][i] <- length(drive_row)
}

# Add Medians
RC$Median_Freq <- 0
RC$Median_Type <- 0
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  med_row <- which(median$ROUTE == RCroute & 
                     median$MP > RCbeg  & 
                     median$MP < RCend)
  med_type <- max(median[["MEDIAN_TYP"]][med_row])
  RC[["Median_Freq"]][i] <- length(med_row)
  RC[["Median_Type"]][i] <- if_else(RC[["Median_Freq"]][i] == 0,"NA", med_type)
}

# Add Shoulders
RC$Right_Shoulder_Freq <- 0
RC$Left_Shoulder_Freq <- 0
RC$Right_Shoulder_Max <- 0
RC$Right_Shoulder_Min <- 0
RC$Left_Shoulder_Max <- 0
RC$Left_Shoulder_Min <- 0
# RC$Right_Shoulder_Avg <- 0
# RC$Left_Shoulder_Avg <- 0
# shoulder <- shoulder %>% arrange(ROUTE, MP)
# RC <- RC %>% arrange(ROUTE, BEG_MP, END_MP)
# row <- 1
for (i in 1:nrow(RC)){
  # start timer
  start.time <- Sys.time()
  
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  
  # r_sho_row <- NA
  # l_sho_row <- NA
  # r_idx <- 1
  # l_idx <- 1
  # repeat{
  #   if(shoulder[["ROUTE"]][row] == RCroute & 
  #      shoulder[["MP"]][row] >= RCbeg  & 
  #      shoulder[["MP"]][row] <= RCend){
  #     if(shoulder[["UTPOSITION"]][row] == "RIGHT"){
  #       r_sho_row[r_idx] <- row
  #       r_idx <- r_idx + 1
  #     } else if(shoulder[["UTPOSITION"]][row] =="LEFT"){
  #       l_sho_row[l_idx] <- row
  #       l_idx <- l_idx + 1
  #     }
  #   } else{
  #     if(shoulder[["MP"]][row] > ){
  #       row <- row + 1
  #       print(paste("Route",shoulder[["ROUTE"]][row],"at MP",shoulder[["MP"]][row],"not on RC"))
  #     }
  #     break #exit the loop once we are no longer on the segment
  #   }
  #   row <- row + 1 #increment row
  # }
  # print(paste(row,RCroute,shoulder[["ROUTE"]][row],RCbeg,RCend,shoulder[["MP"]][row]))
  
  r_sho_row <- which(shoulder$ROUTE == RCroute &
                       shoulder$MP >= RCbeg  &
                       shoulder$MP <= RCend &
                       shoulder$UTPOSITION== "RIGHT")
  l_sho_row <- which(shoulder$ROUTE == RCroute &
                       shoulder$MP >= RCbeg  &
                       shoulder$MP <= RCend &
                       shoulder$UTPOSITION =="LEFT")
  
  RC[["Right_Shoulder_Freq"]][i] <- length(r_sho_row)
  RC[["Left_Shoulder_Freq"]][i] <- length(l_sho_row)
  RC[["Right_Shoulder_Max"]][i] <- if_else(length(r_sho_row) == 0, "NA", as.character(max(shoulder[["SHLDR_WDTH"]][r_sho_row])))
  RC[["Right_Shoulder_Min"]][i] <- if_else(length(r_sho_row) == 0, "NA", as.character(min(shoulder[["SHLDR_WDTH"]][r_sho_row])))
  RC[["Left_Shoulder_Max"]][i] <- if_else(length(l_sho_row) == 0, "NA", as.character(max(shoulder[["SHLDR_WDTH"]][l_sho_row])))
  RC[["Left_Shoulder_Min"]][i] <- if_else(length(l_sho_row) == 0, "NA", as.character(min(shoulder[["SHLDR_WDTH"]][l_sho_row])))
  # RC[["Right_Shoulder_Avg"]][i] <- shoulder[["SHLDR_WDTH"]][r_sho_row]
  # RC[["Left_Shoulder_Avg"]][i] <- ((shoulder[["SHLDR_WDTH"]][l_sho_row]*shoulder[["Length"]][l_sho_row])/(RC[["END_MP"]][i]-RC[["BEG_MP"]][i]))
}
# record time
end.time <- Sys.time()
time.taken <- end.time - start.time
print(paste("Time taken to add shoulders data to segments:", time.taken))

# Add segment id column
RC <- RC %>% rowid_to_column("SEG_ID")

# Save a copy for future use
RC_byseg <- RC

# Pivot AADT (add years)
RC <- pivot_aadt(RC)

# Fixed this with the values_fn parameter of pivot_wider()
# # Unlist AADT for output
# for (i in 1:nrow(RC)){
#   if(length(unlist(RC$AADT[i])) > 1 | length(unlist(RC$SUTRK[i])) > 1 | length(unlist(RC$CUTRK[i])) > 1){
#     RC$AADT[i] <- unlist(RC$AADT[i])[1]
#     RC$SUTRK[i] <- unlist(RC$SUTRK[i])[1]
#     RC$CUTRK[i] <- unlist(RC$CUTRK[i])[1]
#   }
# }
# RC$AADT <- as.numeric(RC$AADT)
# RC$SUTRK <- as.numeric(RC$SUTRK)
# RC$CUTRK <- as.numeric(RC$CUTRK)




###
## Create Intersection Roadway data shell
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

# Add spatial data
IC <- IC %>%
  st_as_sf(
    coords = c("BEG_LONG", "BEG_LAT"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912)
# load UTA stops shapefile
UTA_stops <- read_sf("data/shapefile/UTA_Stops_and_Most_Recent_Ridership.shp") %>%
  st_transform(crs = 26912) %>%
  select(UTA_StopID, Mode) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)
# load schools (not college) shapefile
schools <- read_sf("data/shapefile/Utah_Schools_PreK_to_12.shp") %>%
  st_transform(crs = 26912) %>%
  select(SchoolID, SchoolLeve, OnlineScho, SchoolType, TotalK12) %>%
  filter(is.na(OnlineScho), SchoolType == "Vocational" | SchoolType == "Special Education" | SchoolType == "Residential Treatment" | SchoolType == "Regular Education" | SchoolType == "Alternative") %>%
  select(-OnlineScho, -SchoolType, -TotalK12) %>%
  st_buffer(dist = 304.8) #buffer 1000 ft (units converted to meters)
# modify intersections to include bus stops and schools nearby
IC <- mod_intersections(IC,UTA_Stops,schools)
# remove spatial info
IC <- st_drop_geometry(IC)

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
IC <- pivot_aadt(IC) 
IC <- IC %>% rename(ENT_VEH = AADT)
