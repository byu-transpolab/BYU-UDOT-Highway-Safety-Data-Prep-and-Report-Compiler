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
crash_sf <- st_as_sf(crash, coords = c("long", "lat"), crs = 4326, remove = F) %>%
  st_transform(crash, crs = 26912) %>%
  select(crash_id, intersection_related, long, lat)
int_buff <- st_as_sf(intersection, coords = c("long", "lat"), crs = 4326, remove = F) %>%
  st_transform(intersection, crs = 26912) %>%
  st_buffer(dist = intersection$Leg_Distan * 0.3048) %>% #buffer by FA
  select(Int_ID, long_int=long, lat_int=lat)

# determine which crashes are within a functional area
crash_sf <- st_join(crash_sf, int_buff, join = st_within) %>%
  st_drop_geometry() %>%
  mutate(int_id = ifelse(intersection_related=="N",NA,Int_ID)) %>%
  select(-Int_ID) %>%
  arrange(crash_id)

# choose the closest intersection when there are overlaps
for(i in 1:nrow(crash_sf)){
  if(!is.na(crash_sf$int_id[[i]])){
    # get list of possible intersections
    id <- crash_sf$crash_id[[i]]
    ints <- which(crash_sf$crash_id==id)
    # check for duplicates and find closest intersection
    if(length(ints)>1){
      # assign crash lat long
      lat_crash <- crash_sf$lat[[i]]
      long_crash <- crash_sf$long[[i]]
      # find closest intersection
      closest_id <- 0
      closest_dist <- 100000
      start <- ints[1]
      end <- ints[length(ints)]
      for(j in start:end){
        # assign intersection lat long
        lat_int <- crash_sf$lat_int[[j]]
        long_int <- crash_sf$long_int[[j]]
        # calculate euclidean distance
        dist <- sqrt((lat_crash-lat_int)^2+(long_crash-long_int)^2)
        # check if closest
        if(dist < closest_dist){
          closest_dist <- dist
          closest_id <- crash_sf$int_id[[j]]
        }
      }
      # assign closest intersection id to crash 
      for(k in start:end){
        crash_sf$int_id[k] <- closest_id
      }
      # skip ahead to avoid redundancy
      i <- ints[-1]
    }
  }
}

# delete duplicates
crash_sf <- crash_sf %>%
  select(crash_id, int_id) %>%
  unique()

# join back to crash file
crash <- left_join(crash, crash_sf, by = "crash_id")


### *Below is the old milepoint method. It worked well except that the milepoints
### we were given were not on the correct LRS. We are using a spatial method
### instead now

# crash$int_related <- FALSE  
# crash$int_id <- NA
# for(i in 1:nrow(crash)){
#   rt <- crash$route[i]
#   mp <- crash$milepoint[i]
#   row <- which(FA$ROUTE == rt &
#                  FA$BEG_MP < mp &
#                  FA$END_MP > mp)
#   if(length(row) > 0 & crash$intersection_related[i] == "Y"){
#     crash$int_related[i] <- TRUE
#     if(length(row) > 0){
#       # Identify Intersection IDs
#       # print(paste("ID #",crash$crash_id[i]))
#       closest <- 100
#       closest_row <- NA
#       for(k in 1:length(row)){                    
#         # If Which Returns Multiple Intersections, Returns the Closest
#         # print(FA$Int_ID[row[k]])
#         center <- FA$MP[row[k]]
#         dist <- abs(center - mp)
#         if(dist < closest){
#           closest <- dist
#           closest_row <- row[k]
#         }
#       }
#       crash$int_id[i] <- FA$Int_ID[closest_row]   
#       
#       
#       
#       
#     # NOTICE: SINCE WE NOW ARE USING AND & INSTEAD OF AN | THE CODE UNDER THIS 
#     # ELSE STATEMENT WILL NEVER RUN. IF WE EVER WENT BACK TO AN | STATEMENT THIS
#     # CODE WILL BE USEFUL. RIGHT NOW, CRASHES ASSIGNED AS "INTERSECTION_RELATED"
#     # BUT OUTSIDE AN FA WILL BE ASSIGNED TO SEGMENTS.
#     } else{                                      
#       # Figure Out Closest Intersection if Crash is not within a Functional Area
#       row <- which(FA$ROUTE == rt)      
#       # Redefine Row to Entire Route
#       if(length(row) > 0){
#         # print(paste("ID #",crash$crash_id[i]))
#         closest <- 1000000
#         closest_row <- NA
#         for(k in 1:length(row)){                    
#           # If Which Returns Multiple Intersections, Returns the Closest
#           center <- FA$MP[row[k]]
#           dist <- abs(center - mp)
#           if(dist < closest){
#             closest <- dist
#             closest_row <- row[k]
#           }
#         }
#         # find distance from FA of closest intersection
#         d1 <- abs(mp - FA$BEG_MP[closest_row])
#         d2 <- abs(mp - FA$END_MP[closest_row])
#         offst <- min(d1,d2)
#         # remove crash from analysis if unknown error occurred
#         # remove crash from analysis if distance from FA is too great
#         # assign intersection id to crash otherwise.
#         if(is.na(closest_row)){
#           crash$int_id[i] <- "UNKNOWN"
#           warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. Unknown error."))
#         # buffer the FA to include "intersection_related" crashes close to FA (buffer currently set to zero.)
#         } else if(offst > 0){ 
#           crash$int_id[i] <- "DNE"
#           warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. FA error."))
#         } else{
#           crash$int_id[i] <- FA$Int_ID[closest_row]
#         }
#       # remove crash from analysis if there are no intersections on the route
#       } else{
#         crash$int_id[i] <- "DNE_route"
#         warning(paste("Crash", crash$crash_id[i], "could not be assigned to a segment or intersection. Route error."))
#       }
#     }
#   }
# }

# Filter "Intersection Related" Crashes
crash_seg <- crash %>% filter(is.na(int_id)) %>% select(-int_id)
crash_int <- crash %>% filter(!is.na(int_id))

# Save a csv so you don't have to run everything every time
write_csv(crash, file = "data/temp/crash.csv")
write_csv(crash_seg, file = "data/temp/crash_seg.csv")
write_csv(crash_int, file = "data/temp/crash_int.csv")


