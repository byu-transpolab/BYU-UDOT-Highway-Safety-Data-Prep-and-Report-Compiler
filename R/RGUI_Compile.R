###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Compile Crash & Roadway Data
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

# Add Segment Crashes
RC$TotalCrashes <- 0
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  RCyear <- RC[["YEAR"]][i]
  crash_row <- which(crash_seg$route == RCroute & 
                       crash_seg$milepoint > RCbeg & 
                       crash_seg$milepoint < RCend &
                       crash_seg$crash_year == RCyear)
  RC[["TotalCrashes"]][i] <- length(crash_row)
}

# Write to output
output <- paste0("data/output/",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(RC, file = paste0("CAMS",output))