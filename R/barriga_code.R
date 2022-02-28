library(tidyverse)
library(sf)

# READ IN DATA AND SE
aadt <- read_sf("data/shapefile/AADT_Unrounded.shp")
# plot(aadt)
aadt.filepath <- "data/shapefile/AADT_Unrounded.shp"
aadt.columns <- c("ROUTE_NAME",
                  "START_ACCU", 
              "END_ACCUM",
              "AADT2020", 
              "SUTRK2020", 
              "CUTRK2020", 
              "AADT2019", 
              "SUTRK2019", 
              "CUTRK2019",
              "AADT2018", 
              "SUTRK2018", 
              "CUTRK2018",
              "AADT2017", 
              "SUTRK2017", 
              "CUTRK2017",
              "AADT2016", 
              "SUTRK2016", 
              "CUTRK2016",  
              "AADT2015", 
              "SUTRK2015", 
              "CUTRK2015",
              "AADT2014",
              "SUTRK2014", 
              "CUTRK2014",
              "Shape_Leng",
              "geometry")

fc <- read_sf("data/shapefile/Functional_Class_ALRS.shp")
# plot(fc)
fc.filepath <- "data/shapefile/Functional_Class_ALRS.shp"
fc.columns <- c("ROUTE_ID",                          
            "FROM_MEASU",                             
            "TO_MEASURE",
            "FUNCTIONAL",                             
            "RouteDir",                             
            "RouteType",
            "geometry")

speed <- st_read("data/shapefile/UDOT_Speed_Limits_2019.shp")
# plot(speed)
speed.filepath <- "data/shapefile/UDOT_Speed_Limits_2019.shp"
speed.columns <- c("ROUTE_ID",                         
               "FROM_MEASU",                         
               "TO_MEASURE",                         
               "SPEED_LIMI",
               "geometry")

lane <- read_sf("data/shapefile/Lanes.shp")
# plot(lane)

urban_small <- read_sf("data/shapefile/Urban_Boundaries_Small.shp")
# plot(urban_small)
small_col <- c("ROUTE_ID",                         
               "FROM_MEASURE",                         
               "TO_MEASURE",                         
               "SPEED_LIMIT",
               "geometry")

urban_large <- read_sf("data/shapefile/Urban_Boundaries_Large.shp")
# plot(urban_large)
large_col <- c("OBJECTID",                         
               "TYPE_",
               "geometry")

intersection <- read_sf("data/shapefile/Intersections.shp")
# plot(intersection)

pm <- read_sf("data/shapefile/Pavement_Messages.shp")
# plot(pm)

driveway <- read_sf("data/shapefile/Driveway.shp")
# plot(driveway)

median <- read_sf("data/shapefile/Medians.shp")
# plot(median)

shoulder <- read_sf("data/shapefile/Shoulders.shp")
# plot(shoulder)

## Read in files
read_filez <- function(filepath, columns) {
  if (str_detect(filepath, ".shp")) {
    print("reading shapefile")
    read_sf(filepath) %>% select(all_of(columns))
  }
  else {
    print("Error reading in:")
    print(filepath)
  }
}

###
## Functional Class Data Prep
###

# Read in fc File
fc <- read_filez(fc.filepath, fc.columns)

# Standardizing Column Names
names(fc)[c(1:3)] <- c("ROUTE", "START_ACCUM", "END_ACCUM")

# Only Selecting Main Routes
fc <- fc %>% filter(grepl("M", ROUTE))

#WTHeck is this doing???? Commented out until further analysis
#fc %>%
#  group_by(ROUTE, START_ACCUM) %>%
#  summarize(should_be_one = n()) %>%
#  filter(should_be_one > 1)
#  fc[which(fc$START_ACCUM == 0 & fc$ROUTE == "0201"),]



# Finding Number of Main Routes in the fc File
num.routes <- fc %>% pull(ROUTE) %>% unique() %>% length()
main.routes <- as.character(fc %>% pull(ROUTE) %>% unique() %>% sort())

# Take First Four Numbers of Route Column
fc$ROUTE <- substr(fc$ROUTE, 1, 4)

####
## AADT Data Prep
####

# Read in aadt File
aadt <- read_filez(aadt.filepath, aadt.columns)

# Standardizing Column Names
names(aadt)[c(1:3)] <- c("ROUTE", "START_ACCUM", "END_ACCUM")

# Making duplicate dataset for positive and negative sides of road
# aadt_pos <- aadt_neg <- aadt
# aadt_pos$ROUTE <- as.character(paste(aadt_pos$ROUTE_NAME, "+", sep = ""))
# aadt_neg$ROUTE <- as.character(paste(aadt_neg$ROUTE_NAME, "-", sep = ""))
# aadt <- rbind(aadt_pos, aadt_neg)

#WTHeck is this doing???? Commented out until further analysis
#aadt %>%
 # group_by(ROUTE, START_ACCU) %>%
 # summarize(should_be_one = n()) %>%
 # filter(should_be_one > 1)

# Getting Only Main Routes
aadt <- aadt %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(START_ACCUM < END_ACCUM)

# Take First Four Numbers of Route Column
aadt$ROUTE <- substr(aadt$ROUTE, 1, 4)

###
## Speed Limits Data Prep
###

# Read in speed File
speed <- read_filez(speed.filepath, speed.columns)

# Standardizing Column Names
names(speed)[c(1:3)] <- c("ROUTE", "START_ACCUM", "END_ACCUM")

# Making duplicate dataset for positive and negative sides of road
# sl_pos <- sl_neg <- speed
# sl_pos$ROUTE <- paste(sl_pos$ROUTE, "+", sep = "")
# sl_neg$ROUTE <- paste(sl_neg$ROUTE, "-", sep = "")
# speed <- rbind(sl_pos, sl_neg)

# Getting Only Main Routes
speed <- speed %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(START_ACCUM < END_ACCUM)

#WTHeck is this doing???? Commented out until further analysi
#speed %>%
#  group_by(ROUTE, START_ACCUM) %>%
#  mutate(should_be_one = n()) %>%
#  filter(should_be_one > 1)
#speed[which(speed$ROUTE == "0006" & speed$START_ACCUM == 0), ]

# Take First Four Numbers of Route Column
speed$ROUTE <- substr(speed$ROUTE, 1, 4)

###
##Lanes
###
lanes <- read_filez(lanes.filepath, lanes.columns)

#Taking off ramps
lanes <- lanes %>% filter(nchar(ROUTE) == 5)

#Adding direction onto route variable, taking route direction out of its own column
lanes$ROUTE <- ifelse(lanes$TRAVEL_DIR == "+",
                      paste(substr(lanes$ROUTE, 1, 4), "+", sep = ""),
                      paste(substr(lanes$ROUTE, 1, 4), "-", sep = ""))
lanes <- lanes %>% select(-TRAVEL_DIR)

#Getting only state routes
lanes <- lanes %>% filter(ROUTE %in% state_routes)

#Trying to define each row as separate segments, no overlap, the directions have different information
lanes %>%
  group_by(ROUTE, START_ACCUM) %>%
  summarize(should_be_one = n()) %>%
  filter(should_be_one > 1)
#No duplicates, not all data complete on negative side of the road though


###
##Intersections
###
intersections <- read_filez(intersections.filepath, intersections.columns)

#Finding all intersections with at least one state route
intersections <- intersections %>%
  filter(INT_RT_1 %in% substr(state_routes,1,4) |
           INT_RT_2 %in% substr(state_routes,1,4) | INT_RT_3 %in% substr(state_routes,1,4) |
           INT_RT_4 %in% substr(state_routes,1,4))
#Intersections are labeled at a point, not a segment

#Adding direction onto route variable, taking route direction out of its own column
intersections$ROUTE <- paste(substr(intersections$ROUTE,1,4), intersections$TRAVEL_DIR, sep = "")
intersections <- intersections %>% select(-TRAVEL_DIR)

intersections %>%
  group_by(ROUTE, START_ACCUM) %>%
  summarize(should_be_one = n()) %>%
  filter(should_be_one > 1)

###
##Pavement Messages
###
pavement <- read_filez(pavement.filepath, pavement.columns)

#Getting rid of ramps
pavement <- pavement %>% filter(nchar(ROUTE) == 5)

#Adding direction onto route variable, taking route direction out of its own column
pavement$ROUTE <- paste(substr(pavement$ROUTE,1,4), pavement$TRAVEL_DIR, sep = "")
pavement <- pavement %>% select(-TRAVEL_DIR)

#Getting only state routes
pavement <- pavement %>% filter(ROUTE %in% state_routes)

#Setting all crosswalk types to CROSSWALK
crosswalk_types <- c("CROSSWALK", "HIGH VISIBILITY CROSSWALK - SIDE STREET",
                     "SCHOOL CROSSWALK", "CROSSWALK - SIDE STREET", "HIGH VISIBILITY CROSSWALK",
                     "SCHOOL CROSSWALK - SIDE STREET")
pavement$TYPE <- ifelse(pavement$TYPE %in% crosswalk_types, "CROSSWALK", pavement$TYPE)

p_disc <- pavement %>%
  distinct() %>%
  group_by(ROUTE, START_ACCUM) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, START_ACCUM)


###
##Shoulders
###
shoulders <- read_filez(shoulders.filepath, shoulders.columns)

#Getting rid of ramps
shoulders <- shoulders %>% filter(nchar(ROUTE) == 5)

#Adding direction onto route variable, taking route direction out of its own column
shoulders$ROUTE <- paste(substr(shoulders$ROUTE,1,4), shoulders$TRAVEL_DIR, sep = "")
shoulders <- shoulders %>% select(-TRAVEL_DIR)

#Getting only state routes
shoulders <- shoulders %>% filter(END_ACCUM > START_ACCUM, BEG_LAT < 90, END_LAT < 90)

shd_disc <- shoulders %>%
  group_by(ROUTE, START_ACCUM, UTPOSITION) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, START_ACCUM)

###
##Medians
###
medians <- read_filez(medians.filepath, medians.columns)
names(medians)[1] <- "ROUTE"

#Getting rid of ramps
medians <- medians %>% filter(nchar(ROUTE) == 5, END_ACCUM > START_ACCUM)

#Adding direction onto route variable, taking route direction out of its own column
medians$ROUTE <- paste(substr(medians$ROUTE,1,4), medians$TRAVEL_DIR, sep = "")
medians <- medians %>% select(-TRAVEL_DIR)

#Getting only state routes
medians <- medians %>% filter(ROUTE %in% state_routes)

medians %>%
  group_by(ROUTE, START_ACCUM) %>%
  summarize(should_be_one = n()) %>%
  filter(should_be_one > 1)


###
##Driveways
###

driveway <- read_filez(driveway.filepath, driveway.columns)

#Getting rid of ramps
driveway <- driveway %>% filter(nchar(ROUTE) == 5, END_ACCUM > START_ACCUM)

#Adding direction onto route variable, taking route direction out of its own column
driveway$ROUTE <- paste(substr(driveway$ROUTE,1,4), driveway$DIRECTION, sep = "")
driveway <- driveway %>% select(-DIRECTION)

#Getting only state routes
driveway <- driveway %>% filter(ROUTE %in% state_routes)

driveway <- driveway %>%
  group_by(ROUTE, START_ACCUM) %>%
  mutate(should_be_one = n()) %>%
  arrange(ROUTE, START_ACCUM) %>%
  select(-contains("LONG")) %>%
  select(-contains("LAT")) %>%
  mutate(WIDTH = mean(WIDTH), END_ACCUM = max(END_ACCUM), TYPE = "Minor Residential Driveway") %>%
  distinct()

driveway %>% 
  group_by(ROUTE, START_ACCUM) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1)

########################################################################################
###################################### Aggregation #####################################
########################################################################################

##################################### Data read-in #####################################
# Datasets that we must have
#  sdtm1: AADT Unrounded
#  sdtm2: Functional Class
#  sdtm3: UDOT Speed Limits (2019)
#  sdtm4: Lanes
#  sdtm5: Intersections
#  sdtm6: Pavement Messages
#  sdtm7: Shoulders
#  sdtm8: Medians
#  sdtm9: Driveway

# Maybe put everything in a list? Then we can just lapply everything.
# Yes, we'll do that.
#sdtms <- list()

# For illustrative purposes; to be removed before code is finished.
#sdtm1 <- read.csv("/home/bdahl00/udot-data/original/AADT_Unrounded.csv", header = TRUE) %>% 
#  mutate(ROUTE = ROUTE_NAME,
#         START_ACCUM = START_ACCU)
#sdtm2 <- read.csv("/home/bdahl00/udot-data/original/UDOT_Speed_Limits_(2019).csv", header = TRUE) %>% 
#  mutate(ROUTE = ROUTE_ID,
#         START_ACCUM = FROM_MEASURE,
#         END_ACCUM = TO_MEASURE)
sdtms <- list(aadt, fc, speed)
sdtms <- lapply(sdtms, as_tibble)

####################### Generating time x distance merge shell #########################
# Intersection data has to be handled separately. But talk to the PhDs about how
# intersection lookup is supposed to happen anyway.

# Get all measurements where a road segment stops or starts
segment_breaks <- function(sdtm) {
  breaks_df <- sdtm %>% 
    group_by(ROUTE) %>% 
    summarize(startpoints = unique(c(START_ACCUM, END_ACCUM))) %>% 
    arrange(startpoints)
  return(breaks_df)
}

# This function is robust to data missing for segments of road presumed to exist (such as within
# the bounds of other declared road segments) where there is no data.
# This mostly seems to happen when a road changes hands (state road to interstate, or state to
# city road, or something like that).
shell_creator <- function(sdtms) {
  # This part might have to be adapted to exclude intersections.
  all_breaks <- bind_rows(lapply(sdtms, segment_breaks)) %>% 
    group_by(ROUTE) %>% 
    summarize(startpoints = unique(startpoints)) %>% 
    arrange(ROUTE, startpoints)
  
  endpoints <- c(all_breaks$startpoints[-1],
                 all_breaks$startpoints[length(all_breaks$startpoints)])
  
  all_segments <- bind_cols(all_breaks, endpoints = endpoints) %>% 
    group_by(ROUTE) %>% 
    mutate(lastrecord = (startpoints == max(startpoints))) # %>% 
  #  filter(!lastrecord) %>% 
  #  select(-lastrecord)
  # NOTE: Function produces, then discards one record per ROUTE whose starting and
  # ending point are both the biggest ending END_ACCUM value observed for that ROUTE.
  # This record could be useful depending on how certain characteristics are cataloged,
  # but it is removed here. If you want to un-remove it, just take out the last two
  # filter and select statements (and the pipe above them).
  
  shell <- full_join(all_segments, by = c("ROUTE")) %>% 
    group_by(ROUTE) %>% 
    select(ROUTE, startpoints, endpoints) %>% 
    arrange(ROUTE, startpoints)
  
  return(shell)
}

shell <- shell_creator(sdtms)


# This segment of the program is essentially reconstructing a SAS DATA step with a RETAIN
# statement to populate through rows.

# We start by left joining the shell to each sdtm dataset and populating through all the
# records whose union is the space interval x time interval we have data on. We then get rid
# of the original START_ACCUM, END_ACCUM, FROM_DATE, and TO_DATE variables, since
# startpoints, endpoints, starttime, and endtime have all the information we need.

shell_join <- function(sdtm) {
  joinable_sdtm <- sdtm %>% 
    mutate(startpoints = START_ACCUM)
  with_shell <- left_join(shell, joinable_sdtm, by = c("ROUTE", "startpoints")) %>% 
    select(ROUTE, startpoints, endpoints,
           START_ACCUM, END_ACCUM, everything()) %>% 
    arrange(ROUTE, startpoints)
  
  # pdv contains everyting besides ROUTE, startpoints, endpoints
  # The idea is to reset the pdv everytime we get a new sdtm record, and then copy the 
  # pdv into every row where the window referred to is completely contained in the
  # START_ACCUM to END_ACCUM window 
  pdv <- with_shell[1, 6:ncol(with_shell)]
  for (i in 1:nrow(with_shell)) {
    if (!is.na(with_shell$START_ACCUM[i]) &
        with_shell$START_ACCUM[i] == with_shell$startpoints[i]){
      pdv <- with_shell[i, 6:ncol(with_shell)]
    }
    
    if (!is.na(pdv$START_ACCUM) & !is.na(pdv$END_ACCUM) &
        pdv$START_ACCUM <= with_shell$startpoints[i] &
        with_shell$endpoints[i] <= pdv$END_ACCUM) {
      with_shell[i, 6:ncol(with_shell)] <- pdv
    }
  }
  
  with_shell <- with_shell %>% 
    select(-START_ACCUM, -END_ACCUM,
           # This part is done to avoid merge issues when everything gets
           # joined later
           -endpoints)
  
  return(with_shell)
}

joined_populated <- lapply(sdtms, shell_join)

# To be removed tomorrow, probably (just for accounting for duplicates, which is now
# done in preprocessing)
#joined_populated[[2]] <- joined_populated[[2]] %>% 
#  filter(!(OBJECTID %in% c(15764, 14553, 14990, 19615, 19926)))

################# Merging joined and populated sdtms together; formatting ##################
RC <- c(list(shell), joined_populated) %>% 
  reduce(left_join, by = c("ROUTE", "startpoints")) %>% 
  mutate(START_ACCUM = startpoints,
         END_ACCUM = endpoints) %>% 
  select(-startpoints, -endpoints) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, everything()) %>% 
  arrange(ROUTE, START_ACCUM)

######### Compressing by combining rows identical except for identifying information ########
# First across time
RC <- RC %>% 
  mutate(dropflag = FALSE) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, everything()) %>% 
  arrange(ROUTE, START_ACCUM)


# Adjust time boundaries and flag duplicate rows for deletion.
for (i in 1:(nrow(RC) - 1)) {
  if (identical(RC[i, -c(4, 5)], RC[i + 1, -c(4, 5)])) {
    RC$dropflag[i] <- TRUE
    RC[i + 1, 4] <- RC[i, 4]
  }
}

# Now compressing across road segment
RC <- RC %>% 
  filter(dropflag == FALSE) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, everything()) %>% 
  arrange(ROUTE, START_ACCUM)

# Adjust road segment boundaries and flag duplicate rows for deletion.
for (i in 1:(nrow(RC) - 1)) {
  if (identical(RC[i, -c(4, 5)], RC[i + 1, -c(4, 5)])) {
    RC$dropflag[i] <- TRUE
    RC[i + 1, 4] <- RC[i, 4]
  }
}

RC <- RC %>% 
  filter(dropflag == FALSE) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, everything()) %>% 
  select(-dropflag) %>% 
  arrange(ROUTE, START_ACCUM)

# Write to output
write.csv(RC, file = "data/testoutput.csv")

