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
print(paste("Time taken for code to run:", time.taken))

# Add segment id column
RC <- RC %>% rowid_to_column("SEG_ID")

# Save a copy for future use
RC_byseg <- RC

# Pivot AADT
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
