###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Crash Data Prep
###

# Join Crash Files
crash <- left_join(severity, location, by='crash_id') %>%
  left_join(., rollups, by='crash_id') 

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
crash <- crash %>% filter(route %in% substr(state_routes, 1, 6))

# Add Number of Vehicles Column
vehicle <- vehicle %>%
  group_by(crash_id) %>%
  mutate(num_veh = n())
crash <- left_join(crash, vehicle %>% select(crash_id,num_veh), by = "crash_id") %>%
  unique()

# Separate Crashes into Intersection and Non-Intersection Related
crash$int_related <- FALSE  
crash$int_id <- NA
for(i in 1:nrow(crash)){
  rt <- crash$route[i]
  mp <- crash$milepoint[i]
  row <- which(FA$ROUTE == rt &
                 FA$BEG_MP < mp &
                 FA$END_MP > mp)
  if(length(row) > 0 | crash$intersection_related[i] == "Y"){
    crash$int_related[i] <- TRUE
    if(length(row) > 0){
      # Identify Intersection IDs
      # print(paste("ID #",crash$crash_id[i]))
      closest <- 100
      closest_row <- NA
      for(k in 1:length(row)){                    
        # If Which Returns Multiple Intersections, Returns the Closest
        # print(FA$Int_ID[row[k]])
        center <- FA$MP[row[k]]
        dist <- abs(center - mp)
        if(dist < closest){
          closest <- dist
          closest_row <- row[k]
        }
      }
      crash$int_id[i] <- FA$Int_ID[closest_row]     
    } else{                                      
      # Figure Out Closest Intersection if Crash is not within a Functional Area
      row <- which(FA$ROUTE == rt)      
      # Redefine Row to Entire Route
      if(length(row) > 0){
        # print(paste("ID #",crash$crash_id[i]))
        closest <- 1000000
        closest_row <- NA
        for(k in 1:length(row)){                    
          # If Which Returns Multiple Intersections, Returns the Closest
          center <- FA$MP[row[k]]
          dist <- abs(center - mp)
          if(dist < closest){
            closest <- dist
            closest_row <- row[k]
          }
        }
        # find distance from FA of closest intersection
        d1 <- abs(mp - FA$BEG_MP[closest_row])
        d2 <- abs(mp - FA$END_MP[closest_row])
        offst <- min(d1,d2)
        # remove crash from analysis if unknown error occurred
        # remove crash from analysis if distance from FA is too great
        # assign intersection id to crash otherwise.
        if(is.na(closest_row)){
          crash$int_id[i] <- "UNKNOWN"
          warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. Unknown error."))
        } else if(offst > 0.01){
          crash$int_id[i] <- "DNE"
          warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. FA error."))
        } else{
          crash$int_id[i] <- FA$Int_ID[closest_row]
        }
      # remove crash from analysis if there are no intersections on the route
      } else{
        crash$int_id[i] <- "DNE_route"
        warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. Route error."))
      }
    }
  }
}

# Filter "Intersection Related" Crashes
crash_seg <- crash %>% filter(is.na(int_id)) %>% select(-int_related,-int_id)
crash_int <- crash %>% filter(!is.na(int_id)) %>% select(-int_related)

# Save a csv so you don't have to run everything every time
write_csv(crash, file = "data/temp/crash.csv")
write_csv(crash_seg, file = "data/temp/crash_seg.csv")
write_csv(crash_int, file = "data/temp/crash_int.csv")


