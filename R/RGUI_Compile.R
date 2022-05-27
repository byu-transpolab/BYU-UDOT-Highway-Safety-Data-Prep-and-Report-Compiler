###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Compile Crash & Roadway Data
###

# Add Crashes
RC$TotalCrashes <- 0
for (i in 1:nrow(RC)){
  RCroute <- RC[["ROUTE"]][i]
  RCbeg <- RC[["BEG_MP"]][i]
  RCend <- RC[["END_MP"]][i]
  RCyear <- RC[["YEAR"]][i]
  crash_row <- which(crash$route == RCroute & 
                       crash$milepoint > RCbeg & 
                       crash$milepoint < RCend &
                       crash$crash_year == RCyear)
  RC[["TotalCrashes"]][i] <- length(crash_row)
}

# Write to output
output <- paste0("data/output/",format(Sys.time(),"%d%b%y_%H_%M"),".csv")
write_csv(RC, file = output)