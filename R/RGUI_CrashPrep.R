###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Crash Data Prep
###

# Join Crash Files
crash <- left_join(location,rollups,by='crash_id')

# # Join Crash Vehicle File
# fullcrash <- left_join(crash,vehicle,by='crash_id')

# Filter out Ramps
crash <- crash %>% filter(is.na(ramp_id) | ramp_id == 0)

# Create Separate Date & Time Columns
crash$crash_date <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 1)
crash$crash_time <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 2)
crash$crash_year <- sapply(strsplit(as.character(crash$crash_date), "/"), "[", 3)
crash$crash_year <- as.integer(crash$crash_year)

# Format Routes to Match Roadway
crash$route <- paste(substr(crash$route, 1, 5), crash$route_direction, sep = "")
crash$route <- paste(substr(crash$route, 1, 6), "M", sep = "")
crash$route <- paste0("000", crash$route)
crash$route <- substr(crash$route, nchar(crash$route)-6+1, nchar(crash$route))
crash <- crash %>% filter(route %in% substr(main.routes, 1, 6))

# Add number of vehicles column
vehicle <- vehicle %>%
  group_by(crash_id) %>%
  mutate(num_veh = n())
crash <- left_join(crash, vehicle %>% select(crash_id,num_veh), by = "crash_id") %>%
  unique()

# Separate crashes into intersection and non-intersection related
crash$int_related <- FALSE       # we may just rely on int_id instead
crash$int_id <- NA
for(i in 1:nrow(crash)){
  rt <- crash$route[i]
  mp <- crash$milepoint[i]
  row <- which(FA$ROUTE == rt &
                 FA$BEG_MP < mp &
                 FA$END_MP > mp)
  if(length(row) > 0 | crash$intersection_related[i] == "Y"){
    crash$int_related[i] <- TRUE
    if(length(row) > 0){                           # identify intersection ids
      # print(paste("ID #",crash$crash_id[i]))
      closest <- 100
      closest_row <- NA
      for(k in 1:length(row)){                    # if which returns multiple intersections, return the closest
        # print(FA$Int_ID[row[k]])
        center <- FA$MP[row[k]]
        dist <- abs(center - mp)
        if(dist < closest){
          closest <- dist
          closest_row <- row[k]
        }
      }
      crash$int_id[i] <- FA$Int_ID[closest_row]     
    } else{                                      # figure out closest intersection if it doesn't fall under a functional area
      row <- which(FA$ROUTE == rt)      # redefine row to entire route
      if(length(row) > 0){
        # print(paste("ID #",crash$crash_id[i]))
        closest <- 100
        closest_row <- NA
        for(k in 1:length(row)){                    # if which returns multiple intersections, return the closest
          center <- FA$MP[row[k]]
          dist <- abs(center - mp)
          if(dist < closest){
            closest <- dist
            closest_row <- row[k]
          }
        }
        # print(paste(FA$Int_ID[closest_row],FA$MP[closest_row]))
        if(!is.na(closest_row)){
          crash$int_id[i] <- FA$Int_ID[closest_row]
        }
      }
    }
  }
}

# Filter "Intersection Related" crashes
crash_seg <- crash %>% filter(is.na(int_id)) %>% select(-int_related,-int_id)
crash_int <- crash %>% filter(!is.na(int_id)) %>% select(-int_related)

