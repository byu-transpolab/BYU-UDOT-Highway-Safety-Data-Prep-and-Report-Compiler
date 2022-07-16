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

RC <- RC %>% unique()
RC <- compress_seg_ovr(RC)
RC <- compress_seg_alt(RC)

###
## Fix Spatial Gaps
###

# Delete segments which fall within gaps
RC <- RC %>% mutate(MID_MP = BEG_MP + (END_MP - BEG_MP) / 2)

count <- 0
RC$flag <- 0
for(i in 1:nrow(RC)){
  for(j in 1:nrow(gaps)){
    if(RC$MID_MP[i] > gaps$BEG_GAP[j] & 
       RC$MID_MP[i] < gaps$END_GAP[j] &
       RC$BEG_MP[i] >= gaps$BEG_GAP[j] &
       RC$END_MP[i] <= gaps$END_GAP[j] &
       RC$ROUTE[i] == gaps$ROUTE[j]){
      RC$flag[i] <- 1
      count <- count + 1
    }
  }
}

RC <- RC %>% filter(flag == 0) %>% select(-flag, -MID_MP)
print(paste("deleted",count,"segments within gaps"))

# Compressing segments by length
RC <- compress_seg_len(RC, 0.1)

# Add segment id column
RC <- RC %>% rowid_to_column("SEG_ID")

# # Check segments for issues
# test <- RC
# err1 <- 0
# err2 <- 0
# for(i in 2:nrow(test)){
#   if(test$BEG_MP[i] > test$END_MP[i-1] &
#      test$ROUTE[i] == test$ROUTE[i-1]){
#     chk <- 0
#     for(j in 1:nrow(gaps)){
#       if(test$ROUTE[i] == gaps$ROUTE[j] &
#          round(test$BEG_MP[i],2) == round(gaps$END_GAP[j],2)){
#         chk <- 1
#       }
#     }
#     if(chk == 0){
#       print(paste("unwanted gap at",test$ROUTE[i],test$BEG_MP[i]))
#       err1 <- err1 + 1
#     }
#   }
#   if(test$BEG_MP[i] < test$END_MP[i-1] &
#      test$ROUTE[i] == test$ROUTE[i-1]){
#     print(paste("overlap at",test$ROUTE[i],test$BEG_MP[i]))
#     err2 <- err2 + 1
#   }
# }
# print(paste("there were",err1+err2,"errors in the segments."))
# print(paste("unwanted gaps:",err1))
# print(paste("overlaps:",err2))


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
RC$Median_Type <- NA
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  med_row <- which(median$ROUTE == RCroute & 
                     median$BEG_MP < RCend  & 
                     median$END_MP > RCbeg)
  # get the median types and milepoints on this segment
  med_freq <- length(med_row)
  med_type <- median[["MEDIAN_TYP"]][med_row]
  med_beg <- median[["BEG_MP"]][med_row]
  med_end <- median[["END_MP"]][med_row]
  med_len <- median[["Length"]][med_row]
  if(med_freq > 0){
    # determine internal length of each median
    for(j in 1:length(med_row)){
      if(med_beg[j] < RCbeg){med_beg[j] <- RCbeg}
      if(med_end[j] > RCend){med_end[j] <- RCend}
      med_len[j] <- med_end[j] - med_beg[j]
    }
    # accumulate medians of the same type
    if(med_freq > 1 & j < med_freq){
      for(j in 1:length(med_row)){
        for(k in (j+1):length(med_row)){
          if(med_type[j] == med_type[k]){
            med_len[j] <- med_len[j] + med_len[k]
            med_type[k] <- k  #give it a unique name so it won't be accumulated again
          }
        }
      }
    }
    # return the median by max median length
    med_row <- which.max(med_len)
  }
  # print(med_freq)
  # print(med_type[med_row])
  RC[["Median_Type"]][i] <- ifelse(med_freq == 0L, NA, med_type[med_row])
}

RC <- add_medians(RC, median, 0.001)

# start timer
start.time <- Sys.time()
# Add Shoulders
RC$Right_Shoulder <- NA
RC$Right_Shoulder_Max <- NA
RC$Right_Shoulder_Min <- NA
RC$Right_Shoulder_Avg <- NA
RC$Left_Shoulder <- NA
RC$Left_Shoulder_Max <- NA
RC$Left_Shoulder_Min <- NA
RC$Left_Shoulder_Avg <- NA
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  r_sho_row <- which(shoulder$ROUTE == RCroute &
                       shoulder$END_MP > RCbeg  &
                       shoulder$BEG_MP < RCend &
                       shoulder$UTPOSITION== "RIGHT")
  l_sho_row <- which(shoulder$ROUTE == RCroute &
                       shoulder$END_MP > RCbeg  &
                       shoulder$BEG_MP < RCend &
                       shoulder$UTPOSITION =="LEFT")
  # get the shoulder widths and milepoints on this segment
  r_sho_freq <- length(r_sho_row)
  r_sho_wid <- shoulder[["SHLDR_WDTH"]][r_sho_row]
  r_sho_beg <- shoulder[["BEG_MP"]][r_sho_row]
  r_sho_end <- shoulder[["END_MP"]][r_sho_row]
  r_sho_len <- shoulder[["Length"]][r_sho_row]
  l_sho_freq <- length(l_sho_row)
  l_sho_wid <- shoulder[["SHLDR_WDTH"]][l_sho_row]
  l_sho_beg <- shoulder[["BEG_MP"]][l_sho_row]
  l_sho_end <- shoulder[["END_MP"]][l_sho_row]
  l_sho_len <- shoulder[["Length"]][l_sho_row]
  if(r_sho_freq > 0){
    # determine internal length of each shoulder
    for(j in 1:length(r_sho_row)){
      if(r_sho_beg[j] < RCbeg){r_sho_beg[j] <- RCbeg}
      if(r_sho_end[j] > RCend){r_sho_end[j] <- RCend}
      r_sho_len[j] <- r_sho_end[j] - r_sho_beg[j]
    }
    # determine the length weighted average width
    r_sho_avg <- weighted.mean(r_sho_wid, r_sho_len, na.rm = TRUE)
    # accumulate shoulders of the same width
    if(r_sho_freq > 1 & j < r_sho_freq){
      for(j in 1:length(r_sho_row)){
        for(k in (j+1):length(r_sho_row)){
          if(r_sho_wid[j] == r_sho_wid[k]){
            r_sho_len[j] <- r_sho_len[j] + r_sho_len[k]
            r_sho_wid[k] <- paste0("~",k)  #give the width a unique string so it won't be accumulated again
          }
        }
      }
    }
    # return the shoulder width by max shoulder length
    r_sho_row <- which.max(r_sho_len)
  }
  if(l_sho_freq > 0){
    # determine internal length of each shoulder
    for(j in 1:length(l_sho_row)){
      if(l_sho_beg[j] < RCbeg){l_sho_beg[j] <- RCbeg}
      if(l_sho_end[j] > RCend){l_sho_end[j] <- RCend}
      l_sho_len[j] <- l_sho_end[j] - l_sho_beg[j]
    }
    # determine the length weighted average width
    l_sho_avg <- weighted.mean(l_sho_wid, l_sho_len, na.rm = TRUE)
    # accumulate shoulders of the same width
    if(l_sho_freq > 1 & j < l_sho_freq){
      for(j in 1:length(l_sho_row)){
        for(k in (j+1):length(l_sho_row)){
          if(l_sho_wid[j] == l_sho_wid[k]){
            l_sho_len[j] <- l_sho_len[j] + l_sho_len[k]
            l_sho_wid[k] <- paste0("~",k)  #give the width a unique string so it won't be accumulated again
          }
        }
      }
    }
    # return the shoulder width by max shoulder length
    l_sho_row <- which.max(l_sho_len)
  }
  RC[["Right_Shoulder"]][i] <- ifelse(r_sho_freq == 0, NA, r_sho_wid[r_sho_row])
  RC[["Right_Shoulder_Max"]][i] <- ifelse(r_sho_freq == 0, NA, max(r_sho_wid))
  RC[["Right_Shoulder_Min"]][i] <- ifelse(r_sho_freq == 0, NA, min(r_sho_wid))
  RC[["Right_Shoulder_Avg"]][i] <- ifelse(r_sho_freq == 0, NA, r_sho_avg)
  RC[["Left_Shoulder"]][i] <- ifelse(l_sho_freq == 0, NA, l_sho_wid[l_sho_row])
  RC[["Left_Shoulder_Max"]][i] <- ifelse(l_sho_freq == 0, NA, max(l_sho_wid))
  RC[["Left_Shoulder_Min"]][i] <- ifelse(l_sho_freq == 0, NA, min(l_sho_wid))
  RC[["Left_Shoulder_Avg"]][i] <- ifelse(l_sho_freq == 0, NA, l_sho_avg)
}
# record time
end.time <- Sys.time()
time.taken <- end.time - start.time
print(paste("Time taken to add shoulders data to segments:", time.taken))

# Determine segment length
RC <- RC %>% 
  mutate(
    LENGTH_MILES = END_MP - BEG_MP,
    LENGTH_FEET = LENGTH_MILES * 5280
  ) %>%
  select(SEG_ID:END_MP, LENGTH_MILES, LENGTH_FEET, everything())

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

# Remove infinity values
IC <- IC %>%
  mutate(
    MAX_SPEED_LIMIT = ifelse(is.infinite(MAX_SPEED_LIMIT), NA, MAX_SPEED_LIMIT),
    MIN_SPEED_LIMIT = ifelse(is.infinite(MIN_SPEED_LIMIT), NA, MIN_SPEED_LIMIT)
  )

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
    # take numeric. coerces to NA if there is no number, but not a problem
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

# Add roadway data (disclaimer. Make sure column names don't have a "." in them)
IC <- add_int_att(IC, urban) %>% 
  select(-(MIN_URBAN_CODE:URBAN_CODE_4)) %>%
  rename(URBAN_CODE = MAX_URBAN_CODE)
IC <- add_int_att(IC, fc %>% select(-RouteDir,-RouteType)) %>% 
  select(-(MAX_FUNCTIONAL_CLASS:AVG_FUNCTIONAL_CLASS))
IC <- add_int_att(IC, aadt, TRUE) %>% 
  select(-matches("[0-9]{4,}_[0-9]+"))
IC <- add_int_att(IC, lane) %>%
  select(-(AVG_THRU_CNT:THRU_CNT_4),-(AVG_THRU_WDTH:THRU_WDTH_4))

# Add years to intersection shell and pivot aadt
# yrs <- crash_int %>% select(crash_year) %>% unique()
# IC <- IC %>%
#   group_by(Int_ID) %>%
#   slice(rep(row_number(), times = nrow(yrs))) %>%
#   mutate(YEAR = min(yrs$crash_year):max(yrs$crash_year))
IC <- pivot_aadt_int(IC) %>% 
  mutate(
    NUM_ENT_TRUCKS = (SUTRK + CUTRK) * AADT,
    MAX_NUM_TRUCKS = (MAX_SUTRK + MAX_CUTRK) * MAX_AADT,
    MIN_NUM_TRUCKS = (MIN_SUTRK + MIN_CUTRK) * MIN_AADT,
    AVG_NUM_TRUCKS = (AVG_SUTRK + AVG_CUTRK) * AVG_AADT
  ) %>%
  select(-contains("SUTRK"),-contains("CUTRK")) %>% 
  rename(DAILY_ENT_VEH = AADT)
