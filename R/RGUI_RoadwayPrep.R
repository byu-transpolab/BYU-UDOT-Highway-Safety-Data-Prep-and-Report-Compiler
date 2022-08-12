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
## Merging Joined and Populated sdtms Together; Formatting
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

# Delete Segments Which Fall Within Gaps
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

# Compressing Segments by Length
RC <- compress_seg_len(RC, 0.1)

# Add Segment ID Column
RC <- RC %>% rowid_to_column("SEG_ID")


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
  # Get the Median Types and Milepoints on This Segment
  med_freq <- length(med_row)
  med_type <- median[["MEDIAN_TYP"]][med_row]
  med_beg <- median[["BEG_MP"]][med_row]
  med_end <- median[["END_MP"]][med_row]
  med_len <- median[["Length"]][med_row]
  if(med_freq > 0){
    # Determine internal length of each median
    for(j in 1:length(med_row)){
      if(med_beg[j] < RCbeg){med_beg[j] <- RCbeg}
      if(med_end[j] > RCend){med_end[j] <- RCend}
      med_len[j] <- med_end[j] - med_beg[j]
    }
    # Accumulate medians of the same type
    if(med_freq > 1 & j < med_freq){
      for(j in 1:length(med_row)){
        for(k in (j+1):length(med_row)){
          if(med_type[j] == med_type[k]){
            med_len[j] <- med_len[j] + med_len[k]
            med_type[k] <- k  # give it a unique name so it won't be accumulated again
          }
        }
      }
    }
    # Return the median by max median length
    med_row <- which.max(med_len)
  }
  RC[["Median_Type"]][i] <- ifelse(med_freq == 0L, NA, med_type[med_row])
}

RC <- add_medians(RC, median, 0.001)

# Start Timer
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
  # Get the shoulder widths and milepoints on this segment
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
    # Determine internal length of each shoulder
    for(j in 1:length(r_sho_row)){
      if(r_sho_beg[j] < RCbeg){r_sho_beg[j] <- RCbeg}
      if(r_sho_end[j] > RCend){r_sho_end[j] <- RCend}
      r_sho_len[j] <- r_sho_end[j] - r_sho_beg[j]
    }
    # Determine the length weighted average width
    r_sho_avg <- weighted.mean(r_sho_wid, r_sho_len, na.rm = TRUE)
    # Accumulate shoulders of the same width
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
    # Return the shoulder width by max shoulder length
    r_sho_row <- which.max(r_sho_len)
  }
  if(l_sho_freq > 0){
    # Determine internal length of each shoulder
    for(j in 1:length(l_sho_row)){
      if(l_sho_beg[j] < RCbeg){l_sho_beg[j] <- RCbeg}
      if(l_sho_end[j] > RCend){l_sho_end[j] <- RCend}
      l_sho_len[j] <- l_sho_end[j] - l_sho_beg[j]
    }
    # Determine the length weighted average width
    l_sho_avg <- weighted.mean(l_sho_wid, l_sho_len, na.rm = TRUE)
    # Accumulate shoulders of the same width
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
    # Return the shoulder width by max shoulder length
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
# Record time
end.time <- Sys.time()
time.taken <- end.time - start.time
print(paste("Time taken to add shoulders data to segments:", time.taken))

# Determine Segment Length
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


###
## Create Intersection Roadway data shell
###

# Create Shell for intersection File
IC <- intersection %>% 
  rowwise() %>%
  mutate(
    MAX_SPEED_LIMIT = max(c_across(INT_RT_0_SL:INT_RT_4_SL), na.rm = TRUE),
    MIN_SPEED_LIMIT = min(c_across(INT_RT_0_SL:INT_RT_4_SL), na.rm = TRUE),
    AVG_SPEED_LIMIT = mean(c_across(INT_RT_0_SL:INT_RT_4_SL), na.rm = TRUE)
  ) %>%
  select(Int_ID, everything()) %>%
  select(-contains("_SL"), -contains("_FA"))

# Remove Infinity Values
IC <- IC %>%
  mutate(
    MAX_SPEED_LIMIT = ifelse(is.infinite(MAX_SPEED_LIMIT), NA, MAX_SPEED_LIMIT),
    MIN_SPEED_LIMIT = ifelse(is.infinite(MIN_SPEED_LIMIT), NA, MIN_SPEED_LIMIT),
    AVG_SPEED_LIMIT = ifelse(is.nan(MIN_SPEED_LIMIT), NA, MIN_SPEED_LIMIT)
  )

# append "PM" or "NM" to routes
for(i in 1:nrow(IC)){
  id <- IC[["Int_ID"]][i]
  for(j in 1:4){
    rt <- IC[[paste0("INT_RT_",j)]][i]
    if(!is.na(rt) & tolower(rt) != "local"){    # old condition: rt %in% substr(main.routes,1,4)
      if(rt == substr(IC$INT_RT_0[i],0,4)){
        IC[[paste0("INT_RT_",j)]][i] <- IC$INT_RT_0[i]
      } else{
        IC[[paste0("INT_RT_",j)]][i] <- paste0(rt,"PM")   # assuming if direction not specified it is positive
      }
    }
  }
}

# Add spatial data
IC <- IC %>%
  st_as_sf(
    coords = c("BEG_LONG", "BEG_LAT"), 
    crs = 4326,
    remove = F) %>%
  st_transform(crs = 26912)

# Modify intersections to include bus stops and schools nearby
IC <- mod_intersections(IC,UTA_Stops,schools)

# Remove spatial info
IC <- st_drop_geometry(IC)

# Add roadway data (disclaimer. Make sure column names don't have a "." in them)
# Note: There will be some warnings, but these are resolved within the function so don't mind them.
IC <- add_int_att(IC, urban_full) %>% 
  select(-MAX_URBAN_CODE, -AVG_URBAN_CODE)

IC <- expand_int_att(IC, "URBAN_CODE")

IC <- add_int_att(IC, fc_full %>% select(-RouteDir,-RouteType), is_fc = TRUE) %>% 
  select(-(MAX_FUNCTIONAL_CLASS:AVG_FUNCTIONAL_CLASS))

IC <- expand_int_att(IC, "FUNCTIONAL_CLASS")

IC <- add_int_att(IC, aadt_full, is_aadt = TRUE) %>% 
  select(-matches("[0-9]{4,}_[0-9]+"))  # this is called a regex expression. 
                                        # It helps me find strings that follow certain patterns. 
                                        # In this case, a number followed by an underscore, then another number

IC <- add_int_att(IC, lane_full) %>%
  select(-(THRU_CNT_0:THRU_CNT_4),-(THRU_WDTH_0:THRU_WDTH_4))

IC <- pivot_aadt_int(IC) %>% 
  mutate(
    NUM_ENT_TRUCKS = (SUTRK + CUTRK),
    MAX_NUM_TRUCKS = (MAX_SUTRK + MAX_CUTRK) * MAX_AADT,
    MIN_NUM_TRUCKS = (MIN_SUTRK + MIN_CUTRK) * MIN_AADT,
    AVG_NUM_TRUCKS = (AVG_SUTRK + AVG_CUTRK) * AVG_AADT
  ) %>%
  select(-contains("SUTRK"),-contains("CUTRK")) %>% 
  rename(DAILY_ENT_VEH = AADT)
