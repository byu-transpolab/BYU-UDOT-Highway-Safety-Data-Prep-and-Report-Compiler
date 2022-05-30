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
fullcrash <- left_join(crash,vehicle,by='crash_id')

# Filter out Ramps
crash <- crash %>% filter(is.na(ramp_id) | ramp_id == 0)

# Create Separate Date & Time Columns
crash$crash_date <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 1)
crash$crash_time <- sapply(strsplit(as.character(crash$crash_datetime), " "), "[", 2)
crash$crash_year <- sapply(strsplit(as.character(crash$crash_date), "/"), "[", 3)

# Format Routes to Match Roadway
crash$route <- paste(substr(crash$route, 1, 5), crash$route_direction, sep = "")
crash$route <- paste(substr(crash$route, 1, 6), "M", sep = "")
crash$route <- paste0("000", crash$route)
crash$route <- substr(crash$route, nchar(crash$route)-6+1, nchar(crash$route))
crash <- crash %>% filter(route %in% substr(main.routes, 1, 6))

# Create functional area reference table
FA_ref <- tibble(
  speed = c(5,10,15,20,25,30,35,40,45,50,55,60,65,70,75,80), 
  d1 = c(75,75,75,75,90,110,130,145,165,185,200,220,240,255,275,275), 
  d2 = c(70,70,70,70,105,150,225,290,360,440,525,655,755,875,995,995)
) %>%
  mutate(
    d3 = 50,
    total = d1 + d2 + d3
  )

# Create fed routes list
fed <- read_csv_file(fc_fp, fc_col)
names(fed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
fed <- fed %>% filter(grepl("M", ROUTE))
fed <- fed %>% filter(grepl("Fed Aid", RouteType))
fed.routes <- as.character(fed %>% pull(ROUTE) %>% unique() %>% sort())

# Create full speed file for merging with intersections
speed_full <- read_csv_file(speed_fp, speed_col)
names(speed_full)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
speed_full <- speed_full %>% filter(grepl("M", ROUTE))
speed_full <- compress_seg(speed_full)

# Add speed limit for each intersection approach on intersections file
intersection$INT_RT_1_SL <- NA
intersection$INT_RT_2_SL <- NA
intersection$INT_RT_3_SL <- NA
intersection$INT_RT_4_SL <- NA
for (i in 1:nrow(intersection)){
  int_route1 <- intersection[["INT_RT_1"]][i]
  int_route2 <- intersection[["INT_RT_2"]][i]
  int_route3 <- intersection[["INT_RT_3"]][i]
  int_route4 <- intersection[["INT_RT_4"]][i]
  int_mp_1 <- intersection[["INT_RT_1_M"]][i]
  int_mp_2 <- intersection[["INT_RT_2_M"]][i]
  int_mp_3 <- intersection[["INT_RT_3_M"]][i]
  int_mp_4 <- intersection[["INT_RT_4_M"]][i]
  speed_row1 <- NA
  speed_row2 <- NA
  speed_row3 <- NA
  speed_row4 <- NA
  # get route direction
  suffix = substr(intersection$ROUTE[i],5,6)
  # find rows in speed limit file
  if(!is.na(int_route1) & tolower(int_route1) != "local"){
    int_route1 <- paste0(int_route1, suffix)
    speed_row1 <- which(speed_full$ROUTE == int_route1 & 
                          speed_full$BEG_MP <= int_mp_1 & 
                          speed_full$END_MP >= int_mp_1)
    if(length(speed_row1) == 0 ){
      speed_row1 <- NA
    }
    # print(paste(speed_row1, "at", int_route1, int_mp_1))
  }
  if(!is.na(int_route2) & tolower(int_route2) != "local"){
    int_route2 <- paste0(int_route2, suffix)
    speed_row2 <- which(speed_full$ROUTE == int_route2 & 
                          speed_full$BEG_MP <= int_mp_2 & 
                          speed_full$END_MP >= int_mp_2)
    if(length(speed_row2) == 0 ){
      speed_row2 <- NA
    }
  }
  if(!is.na(int_route3) & tolower(int_route3) != "local"){
    int_route3 <- paste0(int_route3, suffix)
    speed_row3 <- which(speed_full$ROUTE == int_route3 & 
                          speed_full$BEG_MP <= int_mp_3 & 
                          speed_full$END_MP >= int_mp_3)
    if(length(speed_row3) == 0 ){
      speed_row3 <- NA
    }
  }
  if(!is.na(int_route4) & tolower(int_route4) != "local"){
    int_route4 <- paste0(int_route4, suffix)
    speed_row4 <- which(speed_full$ROUTE == int_route4 & 
                          speed_full$BEG_MP <= int_mp_4 & 
                          speed_full$END_MP >= int_mp_4)
    if(length(speed_row4) == 0 ){
      speed_row4 <- NA
    }
  }
  # fill values
  intersection[["INT_RT_1_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row1]
  intersection[["INT_RT_2_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row2]
  intersection[["INT_RT_3_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row3]
  intersection[["INT_RT_4_SL"]][i] <- speed_full$SPEED_LIMIT[speed_row4]
}

# Add Functional Area for each approach
intersection$INT_RT_1_FA <- NA
intersection$INT_RT_2_FA <- NA
intersection$INT_RT_3_FA <- NA
intersection$INT_RT_4_FA <- NA
stopbar <- 60
for(i in 1:nrow(intersection)){
  route1 <- intersection[["INT_RT_1"]][i]
  route2 <- intersection[["INT_RT_2"]][i]
  route3 <- intersection[["INT_RT_3"]][i]
  route4 <- intersection[["INT_RT_4"]][i]
  route1_sl <- intersection[["INT_RT_1_SL"]][i]
  route2_sl <- intersection[["INT_RT_2_SL"]][i]
  route3_sl <- intersection[["INT_RT_3_SL"]][i]
  route4_sl <- intersection[["INT_RT_4_SL"]][i]
  dist1 <- 0
  dist2 <- 0
  dist3 <- 0
  dist4 <- 0
  if(!is.na(route1)){
    if(!is.na(route1_sl)){
      FA_row1 <- which(FA_ref$speed == route1_sl)
      dist1 <- FA_ref$total[FA_row1]
    } else{
      dist1 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_1_FA"]][i] <- dist1 + stopbar
  }
  if(!is.na(route2)){
    if(!is.na(route2_sl)){
      FA_row2 <- which(FA_ref$speed == route2_sl)
      dist2 <- FA_ref$total[FA_row2]
    } else{
      dist2 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_2_FA"]][i] <- dist2 + stopbar
  }
  if(!is.na(route3)){
    if(!is.na(route3_sl)){
      FA_row3 <- which(FA_ref$speed == route3_sl)
      dist3 <- FA_ref$total[FA_row3]
    } else{
      dist3 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_3_FA"]][i] <- dist3 + stopbar
  }
  if(!is.na(route4)){
    if(!is.na(route4_sl)){
      FA_row4 <- which(FA_ref$speed == route4_sl)
      dist4 <- FA_ref$total[FA_row4]
    } else{
      dist4 <- 250   # this is the default if there is no speed limit
    }
    intersection[["INT_RT_4_FA"]][i] <- dist4 + stopbar
  }
}

# Create FA file which has FA's for each route
ROUTE <- NA
BEG_MP <- NA
END_MP <- NA
row <- 0
for(i in 1:nrow(intersection)){
  for(j in 1:4){
    rt <- intersection[[paste0("INT_RT_",j)]][i]
    mp <- intersection[[paste0("INT_RT_",j,"_M")]][i]
    fa <- intersection[[paste0("INT_RT_",j,"_FA")]][i]
    if(rt %in% substr(main.routes,1,4)){
      row <- row + 1
      ROUTE[row] <- rt
      BEG_MP[row] <- mp - (fa/5280)
      END_MP[row] <- mp + (fa/5280)
    }
  }
}
FA <- tibble(ROUTE, BEG_MP, END_MP)

# Separate crashes into intersection and non-intersection related
crash$int_related <- FALSE
for(i in 1:nrow(crash)){
  rt <- crash$route[i]
  mp <- crash$milepoint[i]
  row <- which(FA$ROUTE == rt &
                 FA$BEG_MP < mp &
                 FA$END_MP > mp)
  if(length(row) > 0 | crash$intersection_related == "Y"){
    crash$int_related[i] <- TRUE
  }
}

# Filter "Intersection Related" crashes
crash_seg <- crash %>% filter(int_related == FALSE)
crash_int <- crash %>% filter(int_related == TRUE)
