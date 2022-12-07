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
  RC[["Driveway_Freq"]][i] <- length(drive_row)/(abs(RCend-RCbeg))
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

# Fill all missing data

# First fill missing RouteDir with letter from route name and fill missing 
# RouteType with "State"
RC <- RC %>% 
  mutate(
    RouteDir = if_else(is.na(RouteDir),substr(ROUTE,5,5),RouteDir),
    RouteType = if_else(is.na(RouteType),"State",RouteType)
  )

# Create a list of columns that have missing data to be filled.
missing <- RC %>% 
  ungroup() %>% 
  select(FUNCTIONAL_CLASS:Left_Shoulder_Avg) %>%
  select(-contains("MEDIAN_TYP_"), -RouteDir, -RouteType, -Driveway_Freq) %>%
  colnames()
subsetting_var <- c("Median_Type")
subset_cols <- list(RC %>% 
  ungroup() %>% 
  select(contains("MEDIAN_TYP_")) %>% 
  colnames())

# Change all NA shoulders to 0
# RC$Right_Shoulder[is.na(RC$Right_Shoulder)] <- 0
# RC$Right_Shoulder_Max[is.na(RC$Right_Shoulder_Max)] <- 0 
# RC$Right_Shoulder_Min[is.na(RC$Right_Shoulder_Min)] <- 0 
# RC$Right_Shoulder_Avg[is.na(RC$Right_Shoulder_Avg)] <- 0 
RC$Left_Shoulder[is.na(RC$Left_Shoulder)] <- 0
RC$Left_Shoulder_Max[is.na(RC$Left_Shoulder_Max)] <- 0 
RC$Left_Shoulder_Min[is.na(RC$Left_Shoulder_Min)] <- 0
RC$Left_Shoulder_Avg[is.na(RC$Left_Shoulder_Avg)] <- 0 

# Run the fill_all_missing function to interpolate remaining data
RC <- fill_all_missing(RC, missing, "ROUTE", subsetting_var, subset_cols)

# Replace "no sign" with NA on speed limits
RC <- RC %>%
  mutate(SPEED_LIMIT = ifelse(SPEED_LIMIT == "No Sign", NA, SPEED_LIMIT))

# Manually fill in remaining missing data if they haven't been filled already
RC <- RC %>% 
  mutate(
    COUNTY_CODE = if_else(ROUTE == "0231PM" & is.na(COUNTY_CODE), "Sanpete", COUNTY_CODE),
    UDOT_Region = if_else(ROUTE == "0231PM" & is.na(UDOT_Region), 4, UDOT_Region),
    FUNCTIONAL_CLASS = if_else(ROUTE == "0231PM" & is.na(FUNCTIONAL_CLASS), "Major Collector", FUNCTIONAL_CLASS),
    SPEED_LIMIT = if_else(ROUTE == "0231PM" & is.na(SPEED_LIMIT), 30, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0231PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0231PM" & is.na(THRU_WDTH), 14, THRU_WDTH),
    URBAN_CODE = if_else(ROUTE == "0231PM" & is.na(URBAN_CODE), 99999, URBAN_CODE),
                    
    SPEED_LIMIT = ifelse(ROUTE == "0136PM" & is.na(SPEED_LIMIT), 65, SPEED_LIMIT),
    
    SPEED_LIMIT = ifelse(ROUTE == "0140PM" & is.na(SPEED_LIMIT), 45, SPEED_LIMIT),
    
    SPEED_LIMIT = ifelse(ROUTE == "0168PM" & is.na(SPEED_LIMIT), 35, SPEED_LIMIT),
    
    SPEED_LIMIT = ifelse(ROUTE == "0191NM" & is.na(SPEED_LIMIT), 65, SPEED_LIMIT),
    
    SPEED_LIMIT = ifelse(ROUTE == "0290PM" & is.na(SPEED_LIMIT), 35, SPEED_LIMIT), # hard to tell. Just the road around Snow College
    
    SPEED_LIMIT = ifelse(ROUTE == "0291PM" & is.na(SPEED_LIMIT), 3, SPEED_LIMIT), # school for the deaf. Therefore very low speed limit
    THRU_CNT = if_else(ROUTE == "0291PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0291PM" & is.na(THRU_WDTH), 12, THRU_WDTH),
    Median_Type = if_else(ROUTE == "0291PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    Right_Shoulder = if_else(ROUTE == "0291PM" & is.na(Right_Shoulder), 0, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0291PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0303PM" & is.na(SPEED_LIMIT), 45, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0303PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0303PM" & is.na(THRU_WDTH), 14, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0303PM" & is.na(Driveway_Freq), 1, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0303PM" & is.na(Median_Type), "PAINTED MEDIAN", Median_Type),
    `MEDIAN_TYP_PAINTED MEDIAN` = if_else(ROUTE == "0303PM" & is.na(`MEDIAN_TYP_PAINTED MEDIAN`), 1, `MEDIAN_TYP_PAINTED MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0303PM" & is.na(Right_Shoulder), 6, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0303PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0304PM" & is.na(SPEED_LIMIT), 15, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0304PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0304PM" & is.na(THRU_WDTH), 12, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0304PM" & is.na(Driveway_Freq), 28, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0304PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    `MEDIAN_TYP_NO MEDIAN` = if_else(ROUTE == "0304PM" & is.na(`MEDIAN_TYP_NO MEDIAN`), 1, `MEDIAN_TYP_NO MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0304PM" & is.na(Right_Shoulder), 0, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0304PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0306PM" & is.na(SPEED_LIMIT), 15, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0306PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0306PM" & is.na(THRU_WDTH), 10, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0306PM" & is.na(Driveway_Freq), 1, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0306PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    `MEDIAN_TYP_NO MEDIAN` = if_else(ROUTE == "0306PM" & is.na(`MEDIAN_TYP_NO MEDIAN`), 1, `MEDIAN_TYP_NO MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0306PM" & is.na(Right_Shoulder), 0, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0306PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0309PM" & is.na(SPEED_LIMIT), 25, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0309PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0309PM" & is.na(THRU_WDTH), 10, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0309PM" & is.na(Driveway_Freq), 3, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0309PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    `MEDIAN_TYP_NO MEDIAN` = if_else(ROUTE == "0309PM" & is.na(`MEDIAN_TYP_NO MEDIAN`), 1, `MEDIAN_TYP_NO MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0309PM" & is.na(Right_Shoulder), 2, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0309PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0310PM" & is.na(SPEED_LIMIT), 15, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0310PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0310PM" & is.na(THRU_WDTH), 10, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0310PM" & is.na(Driveway_Freq), 2, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0310PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    `MEDIAN_TYP_NO MEDIAN` = if_else(ROUTE == "0310PM" & is.na(`MEDIAN_TYP_NO MEDIAN`), 1, `MEDIAN_TYP_NO MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0310PM" & is.na(Right_Shoulder), 2, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0310PM" & is.na(Left_Shoulder), 0, Left_Shoulder),
    
    SPEED_LIMIT = ifelse(ROUTE == "0317PM" & is.na(SPEED_LIMIT), 15, SPEED_LIMIT),
    THRU_CNT = if_else(ROUTE == "0317PM" & is.na(THRU_CNT), 2, THRU_CNT),
    THRU_WDTH = if_else(ROUTE == "0317PM" & is.na(THRU_WDTH), 12, THRU_WDTH),
    Driveway_Freq = if_else(ROUTE == "0317PM" & is.na(Driveway_Freq), 1, Driveway_Freq),
    Median_Type = if_else(ROUTE == "0317PM" & is.na(Median_Type), "NO MEDIAN", Median_Type),
    `MEDIAN_TYP_NO MEDIAN` = if_else(ROUTE == "0317PM" & is.na(`MEDIAN_TYP_NO MEDIAN`), 1, `MEDIAN_TYP_NO MEDIAN`),
    Right_Shoulder = if_else(ROUTE == "0317PM" & is.na(Right_Shoulder), 0, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0317PM" & is.na(Left_Shoulder), 0, Left_Shoulder),

    Median_Type = if_else(ROUTE == "0319PM" & is.na(Median_Type), "TWO WAY LEFT TURN LANE", Median_Type),
    `MEDIAN_TYP_TWO WAY LEFT TURN LANE` = if_else(ROUTE == "0319PM" & is.na(`MEDIAN_TYP_TWO WAY LEFT TURN LANE`), 1, `MEDIAN_TYP_TWO WAY LEFT TURN LANE`),
    Right_Shoulder = if_else(ROUTE == "0319PM" & is.na(Right_Shoulder), 12, Right_Shoulder),
    Left_Shoulder = if_else(ROUTE == "0319PM" & is.na(Left_Shoulder), 12, Left_Shoulder)
  )

# Add Interstate Column
RC <- RC %>% mutate(INTERSTATE = case_when(FUNCTIONAL_CLASS=="Interstate"~"1",FUNCTIONAL_CLASS!="Interstate"~"0"))

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
    INT_RT_0_SL = as.integer(INT_RT_0_SL),
    INT_RT_1_SL = as.integer(INT_RT_1_SL),
    INT_RT_2_SL = as.integer(INT_RT_2_SL),
    INT_RT_3_SL = as.integer(INT_RT_3_SL),
    INT_RT_4_SL = as.integer(INT_RT_4_SL),
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
    AVG_SPEED_LIMIT = ifelse(is.nan(AVG_SPEED_LIMIT), NA, AVG_SPEED_LIMIT)
  )

# append "PM" or "NM" to routes
for(i in 1:nrow(IC)){
  id <- IC[["Int_ID"]][i]
  for(j in 1:4){
    rt <- IC[[paste0("INT_RT_",j)]][i]
    if(!is.na(rt) & is.numeric(rt)){
      if(as.integer(rt) == as.integer(gsub(".*?([0-9]+).*", "\\1", IC$INT_RT_0[i]))){
        IC[[paste0("INT_RT_",j)]][i] <- IC$INT_RT_0[i]
      } else{
        # assuming if direction not specified it is positive
        if(as.integer(rt)<10){IC[[paste0("INT_RT_",j)]][i] <- paste0("000",rt,"PM")}
        else if(as.integer(rt)<100){IC[[paste0("INT_RT_",j)]][i] <- paste0("00",rt,"PM")}
        else if(as.integer(rt)<1000){IC[[paste0("INT_RT_",j)]][i] <- paste0("0",rt,"PM")}
        else{IC[[paste0("INT_RT_",j)]][i] <- paste0(rt,"PM")}
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

IC <- IC %>% rename(URBAN_CODE = MIN_URBAN_CODE)

IC <- add_int_att(IC, fc_full %>% select(-RouteDir,-RouteType), is_fc = TRUE) %>% 
  select(-MIN_FUNCTIONAL_CLASS, -AVG_FUNCTIONAL_CLASS, -MAX_FUNCTIONAL_CLASS,
         -(COUNTY_CODE_0:COUNTY_CODE_4), -(UDOT_Region_0:UDOT_Region_4),
         -MAX_COUNTY_CODE, -AVG_COUNTY_CODE,
         -MAX_UDOT_Region, -AVG_UDOT_Region) %>%
  mutate(PRINCIPAL_FUNCTIONAL_CLASS = FUNCTIONAL_CLASS_0) %>%
  rename(COUNTY_CODE = MIN_COUNTY_CODE,
         UDOT_Region = MIN_UDOT_Region)

IC <- expand_int_att(IC, "FUNCTIONAL_CLASS")

IC <- add_int_att(IC, aadt_full, is_aadt = TRUE) %>% 
  select(-matches("[0-9]{4,}_[0-9]+"))  # this is called a regex expression. 
                                        # It helps me find strings that follow certain patterns. 
                                        # In this case, a number followed by an underscore, then another number

IC <- add_int_att(IC, lane_full) %>%
  select(-(THRU_WDTH_0:THRU_WDTH_4))

# Add Total Thru Lane Column
IC <- IC %>%
  rowwise() %>%
  mutate(
    KNOWN_LEGS = 5L - sum(is.na(c(THRU_CNT_0,
                                  THRU_CNT_1,
                                  THRU_CNT_2,
                                  THRU_CNT_3,
                                  THRU_CNT_4))),
    TOTAL_THRU_CNT = sum(c_across(THRU_CNT_0:THRU_CNT_4), na.rm = TRUE),
    TOTAL_THRU_CNT = case_when(
      NUM_LEGS - KNOWN_LEGS == 0 ~ TOTAL_THRU_CNT,
      NUM_LEGS - KNOWN_LEGS == 1 ~ TOTAL_THRU_CNT + 1,
      NUM_LEGS - KNOWN_LEGS == 2 ~ TOTAL_THRU_CNT + 2,
      NUM_LEGS - KNOWN_LEGS == 3 ~ TOTAL_THRU_CNT + 3,
      NUM_LEGS - KNOWN_LEGS == 4 ~ TOTAL_THRU_CNT + 4,
      NUM_LEGS - KNOWN_LEGS == 5 ~ TOTAL_THRU_CNT + 5,)
    ) %>%
  select(-(THRU_CNT_0:THRU_CNT_4))

# pivot by num legs for statistical purposes (setting all_legs to false allows
# us to do this even though there is only one column for NUM_LEGS)
IC <- expand_int_att(IC, "NUM_LEGS", all_legs = FALSE)

# Create a list of columns that have missing data to be filled.
missing <- IC %>% 
  ungroup() %>% 
  select(MAX_SPEED_LIMIT,
         MIN_SPEED_LIMIT,
         AVG_SPEED_LIMIT,
         URBAN_CODE, 
         PRINCIPAL_FUNCTIONAL_CLASS,
         COUNTY_CODE,
         UDOT_Region,
         MAX_THRU_CNT:AVG_THRU_WDTH) %>%
  colnames()
subsetting_var <- c("URBAN_CODE", "PRINCIPAL_FUNCTIONAL_CLASS")
subset_cols <- list(IC %>% 
    ungroup() %>% 
    select(contains("URBAN_CODE_")) %>% 
    colnames(),
  IC %>% 
    ungroup() %>% 
    select(contains("FUNCTIONAL_CLASS_")) %>% 
    colnames())

# Run the fill_all_missing function to interpolate remaining data
IC <- fill_all_missing(IC, missing, "INT_RT_0", subsetting_var, subset_cols)

# Get rid of remaining unnecessary columns (we can do this more cleanly in the future.)
IC <- IC %>%
  select(-(URBAN_CODE_99999:URBAN_CODE_78499))
         # -`FUNCTIONAL_CLASS_Other Principal Arterial`,
         # -(`FUNCTIONAL_CLASS_Minor Arterial`:`FUNCTIONAL_CLASS_Minor Collector`),
         # -MIN_THRU_CNT, -MIN_THRU_WDTH, -AVG_THRU_CNT, -AVG_THRU_WDTH)

# Manually fill in remaining missing data
IC <- IC %>%
  mutate(
    IntVol2020 = ifelse(Int_ID == 903 & IntVol2020 == 0, 72445, IntVol2020),
    IntVol2019 = ifelse(Int_ID == 903 & IntVol2019 == 0, 71169, IntVol2019),
    IntVol2018 = ifelse(Int_ID == 903 & IntVol2018 == 0, 65980, IntVol2018),
    IntVol2017 = ifelse(Int_ID == 903 & IntVol2017 == 0, 69160, IntVol2017),
    IntVol2016 = ifelse(Int_ID == 903 & IntVol2016 == 0, 65768, IntVol2016),
    MEV_2020 = ifelse(Int_ID == 903 & MEV2020 == 0, 26.44, MEV_2020),
    MEV_2019 = ifelse(Int_ID == 903 & MEV2019 == 0, 25.98, MEV_2019),
    MEV_2018 = ifelse(Int_ID == 903 & MEV2018 == 0, 24.08, MEV_2018),
    MEV_2017 = ifelse(Int_ID == 903 & MEV2017 == 0, 25.24, MEV_2017),
    MEV_2016 = ifelse(Int_ID == 903 & MEV2016 == 0, 24.01, MEV_2016),
    
    IntVol2020 = ifelse(Int_ID == 4017 & IntVol2020 == 0, 11305, IntVol2020),
    IntVol2019 = ifelse(Int_ID == 4017 & IntVol2019 == 0, 12994, IntVol2019),
    IntVol2018 = ifelse(Int_ID == 4017 & IntVol2018 == 0, 12865, IntVol2018),
    IntVol2017 = ifelse(Int_ID == 4017 & IntVol2017 == 0, 11759, IntVol2017),
    IntVol2016 = ifelse(Int_ID == 4017 & IntVol2016 == 0, 10809, IntVol2016),
    MEV_2020 = ifelse(Int_ID == 4017 & MEV2020 == 0, 4.13, MEV_2020),
    MEV_2019 = ifelse(Int_ID == 4017 & MEV2019 == 0, 4.74, MEV_2019),
    MEV_2018 = ifelse(Int_ID == 4017 & MEV2018 == 0, 4.70, MEV_2018),
    MEV_2017 = ifelse(Int_ID == 4017 & MEV2017 == 0, 4.29, MEV_2017),
    MEV_2016 = ifelse(Int_ID == 4017 & MEV2016 == 0, 3.95, MEV_2016),
    
    IntVol2020 = ifelse(Int_ID == 4022 & IntVol2020 == 0, 15593, IntVol2020),
    IntVol2019 = ifelse(Int_ID == 4022 & IntVol2019 == 0, 17774, IntVol2019),
    IntVol2018 = ifelse(Int_ID == 4022 & IntVol2018 == 0, 17608, IntVol2018),
    IntVol2017 = ifelse(Int_ID == 4022 & IntVol2017 == 0, 17395, IntVol2017),
    IntVol2016 = ifelse(Int_ID == 4022 & IntVol2016 == 0, 16964, IntVol2016),
    MEV_2020 = ifelse(Int_ID == 4022 & MEV2020 == 0, 5.69, MEV_2020),
    MEV_2019 = ifelse(Int_ID == 4022 & MEV2019 == 0, 6.49, MEV_2019),
    MEV_2018 = ifelse(Int_ID == 4022 & MEV2018 == 0, 6.43, MEV_2018),
    MEV_2017 = ifelse(Int_ID == 4022 & MEV2017 == 0, 6.35, MEV_2017),
    MEV_2016 = ifelse(Int_ID == 4022 & MEV2016 == 0, 6.19, MEV_2016),
    
    IntVol2020 = ifelse(Int_ID == 4030 & IntVol2020 == 0, 29839, IntVol2020),
    IntVol2019 = ifelse(Int_ID == 4030 & IntVol2019 == 0, 34148, IntVol2019),
    IntVol2018 = ifelse(Int_ID == 4030 & IntVol2018 == 0, 33820, IntVol2018),
    IntVol2017 = ifelse(Int_ID == 4030 & IntVol2017 == 0, 14473, IntVol2017),
    IntVol2016 = ifelse(Int_ID == 4030 & IntVol2016 == 0, 14119, IntVol2016),
    MEV_2020 = ifelse(Int_ID == 4030 & MEV2020 == 0, 10.89, MEV_2020),
    MEV_2019 = ifelse(Int_ID == 4030 & MEV2019 == 0, 12.46, MEV_2019),
    MEV_2018 = ifelse(Int_ID == 4030 & MEV2018 == 0, 12.34, MEV_2018),
    MEV_2017 = ifelse(Int_ID == 4030 & MEV2017 == 0, 5.28, MEV_2017),
    MEV_2016 = ifelse(Int_ID == 4030 & MEV2016 == 0, 5.15, MEV_2016),
    
    IntVol2020 = ifelse(Int_ID == 4031 & IntVol2020 == 0, 29839, IntVol2020),
    IntVol2019 = ifelse(Int_ID == 4031 & IntVol2019 == 0, 34148, IntVol2019),
    IntVol2018 = ifelse(Int_ID == 4031 & IntVol2018 == 0, 33820, IntVol2018),
    IntVol2017 = ifelse(Int_ID == 4031 & IntVol2017 == 0, 14473, IntVol2017),
    IntVol2016 = ifelse(Int_ID == 4031 & IntVol2016 == 0, 14119, IntVol2016),
    MEV_2020 = ifelse(Int_ID == 4031 & MEV2020 == 0, 10.89, MEV_2020),
    MEV_2019 = ifelse(Int_ID == 4031 & MEV2019 == 0, 12.46, MEV_2019),
    MEV_2018 = ifelse(Int_ID == 4031 & MEV2018 == 0, 12.34, MEV_2018),
    MEV_2017 = ifelse(Int_ID == 4031 & MEV2017 == 0, 5.28, MEV_2017),
    MEV_2016 = ifelse(Int_ID == 4031 & MEV2016 == 0, 5.15, MEV_2016),
    
    MAX_THRU_CNT = if_else(Int_ID == 4056 & MAX_THRU_CNT == 0, 2, MAX_THRU_CNT),
    MAX_THRU_WDTH = if_else(Int_ID == 4056 & MAX_THRU_WDTH == 0, 12, MAX_THRU_WDTH),
    
    MAX_THRU_CNT = if_else(Int_ID == 4058 & MAX_THRU_CNT == 0, 2, MAX_THRU_CNT),
    MAX_THRU_WDTH = if_else(Int_ID == 4058 & MAX_THRU_WDTH == 0, 12, MAX_THRU_WDTH),
    
    MAX_SPEED_LIMIT = if_else(Int_ID == 5537 & is.na(MAX_SPEED_LIMIT), 40, MAX_SPEED_LIMIT),
    MAX_THRU_CNT = if_else(Int_ID == 5537 & is.na(MAX_THRU_CNT), 2, MAX_THRU_CNT),
    MAX_THRU_WDTH = if_else(Int_ID == 5537 & is.na(MAX_THRU_WDTH), 12, MAX_THRU_WDTH),
    
    MAX_THRU_CNT = if_else(Int_ID == 5943 & is.na(MAX_THRU_CNT), 2, MAX_THRU_CNT),
    MAX_THRU_WDTH = if_else(Int_ID == 5943 & is.na(MAX_THRU_WDTH), 12, MAX_THRU_WDTH),
    
    IntVol2016 = ifelse(Int_ID == "7613" & IntVol2016 == 0, 17833.667, IntVol2016),
    MEV_2016 = ifelse(Int_ID == "7613" & MEV_2016 == 0, 6.509, MEV_2016),
    
    IntVol2016 = ifelse(Int_ID == "7614" & IntVol2016 == 0, 17833.667, IntVol2016),
    MEV_2016 = ifelse(Int_ID == "7614" & MEV_2016 == 0, 6.509, MEV_2016),
  
    COUNTY_CODE = if_else(Int_ID == 9916 & is.na(COUNTY_CODE), "Wasatch" , COUNTY_CODE),
    UDOT_Region = if_else(Int_ID == 9916 & is.na(UDOT_Region), 3 , UDOT_Region),
    PRINCIPAL_FUNCTIONAL_CLASS = if_else(Int_ID == 9916 & is.na(PRINCIPAL_FUNCTIONAL_CLASS), "Other Principal Arterial" , PRINCIPAL_FUNCTIONAL_CLASS)
    )

# Save a copy for future use
IC_byint <- IC

# Pivot AADT (add years)
IC <- pivot_aadt_int(IC) %>% 
  mutate(
    ENT_TRUCKS = (SUTRK + CUTRK),
    MAX_ENT_TRUCKS = (MAX_SUTRK + MAX_CUTRK) * MAX_AADT,
    MIN_ENT_TRUCKS = (MIN_SUTRK + MIN_CUTRK) * MIN_AADT,
    AVG_ENT_TRUCKS = (AVG_SUTRK + AVG_CUTRK) * AVG_AADT
  ) %>%
  rename(ESTIMATED_ENT_VEH = AADT) %>%
  select(-contains("SUTRK"),-contains("CUTRK"),-contains("AADT"),
         -MIN_ENT_TRUCKS,-AVG_ENT_TRUCKS) %>% 
  rename(ENT_VEH = IntVol,
         MEV = MEV_)

