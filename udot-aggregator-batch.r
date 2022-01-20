# I know next to nothing about how roadway datasets work, although I have learned a
# little over the course of this project. If this code doesn't work at some point
# in the future, it's probably because I either made an error or I didn't fully
# understand the structure of the datasets operated on. For questions, please email
# benjamin@ddahl.org.

# Library loading
library(tidyverse)
library(lubridate)

############################## Preprocessing (due to Caleb) ############################
## AADT Unrounded (No RT_Type)
aadt.filepath <- "/home/bdahl00/udot-data/original/AADT_Unrounded.csv"
aadt.columns <- c("START_ACCU", 
                  "END_ACCUM", 
                  "ROUTE_NAME", 
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
                  "Shape_Length")

## AADT Rounded
aadtrounded.filepath <- "/home/bdahl00/udot-data/original/AADT_Rounded.csv"
aadtrounded.columns <- c("RT_Type",
                         "ROUTE_NAME",
                         "START_ACCU",
                         "END_ACCUM",
                         "Shape_Length")


## Functional Class
functionalclass.filepath <- "/home/bdahl00/udot-data/original/Functional_Class_(ALRS).csv"
functionalclass.columns <- c("ROUTE_ID",
                             "FROM_MEASURE",
                             "TO_MEASURE",
                             "FROM_DATE", # Insertion by me
                             "TO_DATE", # Insertion by me
                             "FUNCTIONAL_CLASS",
                             "RouteDir",
                             "RouteType")

## UDOT Speed Limits (2019)
speedlimits.filepath <- "/home/bdahl00/udot-data/original/UDOT_Speed_Limits_(2019).csv"
speedlimits.columns <- c("ROUTE_ID",
                         "FROM_MEASURE",
                         "TO_MEASURE",
                         "FROM_DATE",
                         "TO_DATE",
                         "SPEED_LIMIT")

## Lanes
lanes.filepath <- "/home/bdahl00/udot-data/original/Lanes.csv"
lanes.columns <- c("ROUTE",
                   "START_ACCUM",
                   "END_ACCUM",
                   "L_TURN_CNT",
                   "R_TURN_CNT",
                   "THRU_CNT",
                   "THRU_WDTH",
                   "BEG_LONG",
                   "BEG_LAT",
                   "END_LONG",
                   "END_LAT",
                   "COLL_DATE",
                   "TRAVEL_DIR")

## Intersections
intersections.filepath <- "/home/bdahl00/udot-data/original/Intersections.csv"
intersections.columns <- c("ROUTE",
                           "TRAVEL_DIR",
                           "START_ACCUM",
                           "END_ACCUM",
                           "INT_TYPE",
                           "TRAFFIC_CO",
                           "INT_RT_1",
                           "INT_RT_2",
                           "INT_RT_3",
                           "INT_RT_4",
                           "BEG_LONG",
                           "BEG_LAT")

## Pavement Messages
pavement.filepath <- "/home/bdahl00/udot-data/original/Pavement_Messages.csv"
pavement.columns <- c("ROUTE",
                      "TRAVEL_DIR",
                      "START_ACCUM",
                      "END_ACCUM",
                      "TYPE",
                      "LEGEND")

## Shoulders
shoulders.filepath <- "/home/bdahl00/udot-data/original/Shoulders.csv"
shoulders.columns <- c("ROUTE",
                       "TRAVEL_DIR",
                       "START_ACCUM",
                       "END_ACCUM",
                       "UTPOSITION",
                       "SHLDR_WDTH",
                       "BEG_LONG",
                       "BEG_LAT",
                       "END_LONG",
                       "END_LAT")

## Medians
medians.filepath <- "/home/bdahl00/udot-data/original/Medians.csv"
medians.columns <- c("ROUTE_NAME",
                     "TRAVEL_DIR",
                     "START_ACCUM",
                     "END_ACCUM",
                     "MEDIAN_TYP",
                     "TRFISL_TYP",
                     "MDN_PRTCTN",
                     "BEG_LONG",
                     "BEG_LAT",
                     "END_LONG",
                     "END_LAT")

## Driveway
driveway.filepath <- "/home/bdahl00/udot-data/original/Asset_Data_2014.csv"
driveway.columns <- c("ROUTE",
                      "DIRECTION",
                      "START_ACCUM",
                      "END_ACCUM",
                      "TYPE",
                      "WIDTH",
                      "BEGIN_LONG",
                      "BEGIN_LAT",
                      "END_LONG",
                      "END_LAT")

## Read in files
read_filez <- function(filepath, columns) {
  if (str_detect(filepath, ".csv")) {
    print("reading csv")
    read.csv(filepath) %>% select(all_of(columns))
  }
  else {
    print("Error reading in:")
    print(filepath)
  }
}

####
##AADT Dataset Prep
####

#AADT Unrounded
aadt <- read_filez(aadt.filepath, aadt.columns)
names(aadt)[c(1,3)] <- c("START_ACCUM", "ROUTE")

#AADT Rounded (For route type)
aadt_rounded <- read_filez(aadtrounded.filepath, aadtrounded.columns)

#Sorting by shape length(each is unique) so we can get the route type from the aadt_rounded
#Only a four routes were not identical between the two data sets and those were because the
#starting or ending point were .001 off
aadt <- aadt[order(aadt$Shape_Length), ]
aadt_rounded <- aadt_rounded[order(aadt_rounded$Shape_Length), ]

#Final AADT set, adding on route type then filtering the state routes, removing shape length
#as it is no longer needed
aadt <- cbind(aadt, aadt_rounded$RT_Type)
names(aadt)[length(aadt)] <- "RT_Type" #Fixing variable name from cbind
aadt <- aadt %>% filter(RT_Type == "State Route") %>% select(-c(Shape_Length, RT_Type))
aadt$ROUTE <- substr(aadt$ROUTE, 1, 4)

#Making duplicate dataset for positive and negative sides of road
aadt_pos <- aadt_neg <- aadt
aadt_pos$ROUTE <- as.character(paste(aadt_pos$ROUTE, "+", sep = ""))
aadt_neg$ROUTE <- as.character(paste(aadt_neg$ROUTE, "-", sep = ""))
aadt <- rbind(aadt_pos, aadt_neg)

#Finding number of state routes in the aadt file
num.routes <- aadt %>% pull(ROUTE) %>% unique() %>% length()
state_routes <- as.character(aadt %>% pull(ROUTE) %>% unique() %>% sort())
#What to do with NAs for AADT Unrounded dataset

aadt %>%
  group_by(ROUTE, START_ACCUM) %>%
  summarize(should_be_one = n()) %>%
  filter(should_be_one > 1)

###
##Functional Class
###
functionalclass <- read_filez(functionalclass.filepath, functionalclass.columns)
#Standardizing names
names(functionalclass)[c(1:3)] <- c("ROUTE", "START_ACCUM", "END_ACCUM")
#Taking off ramps
functionalclass <- functionalclass %>% filter(nchar(ROUTE) == 6)

#Adding direction onto route variable, taking route direction out of its own column
functionalclass$ROUTE <- ifelse(functionalclass$RouteDir == "P",
                                paste(substr(functionalclass$ROUTE, 1, 4), "+", sep = ""),
                                paste(substr(functionalclass$ROUTE, 1, 4), "-", sep = ""))
functionalclass <- functionalclass %>% select(-RouteDir)

#Filtering state routes
functionalclass <- functionalclass %>% filter(ROUTE %in% state_routes)

functionalclass %>%
  group_by(ROUTE, START_ACCUM) %>%
  summarize(should_be_one = n()) %>%
  filter(should_be_one > 1)
# functionalclass[which(functionalclass$START_ACCUM == 0 & functionalclass$ROUTE == "0201"),]

###
##Speed Limits
###
speedlimits <- read_filez(speedlimits.filepath, speedlimits.columns)
#Standardizing names
names(speedlimits)[c(1:3)] <- c("ROUTE", "START_ACCUM", "END_ACCUM")

#Taking off ramps
speedlimits <- speedlimits %>% filter(nchar(ROUTE) == 6)
speedlimits$ROUTE <- substr(speedlimits$ROUTE, 1, 4)
# sl_pos <- sl_neg <- speedlimits
# sl_pos$ROUTE <- paste(sl_pos$ROUTE, "+", sep = "")
# sl_neg$ROUTE <- paste(sl_neg$ROUTE, "-", sep = "")
# speedlimits <- rbind(sl_pos, sl_neg)

#Getting only state routes
speedlimits <- speedlimits %>% filter(ROUTE %in% substr(state_routes, 1, 4)) %>%
  filter(START_ACCUM < END_ACCUM)

speedlimits %>%
  group_by(ROUTE, START_ACCUM, FROM_DATE) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1)

speedlimits[which(speedlimits$ROUTE == "0006" & speedlimits$START_ACCUM == 0), ]

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
sdtms <- list(aadt, functionalclass, lanes, medians, driveway)
sdtms <- lapply(sdtms, as_tibble)

################################ Making time numeric ###################################
current_time <- Sys.time()

has_date_col <- function(sdtm) {
  return("TO_DATE" %in% colnames(sdtm))
}

sdtms_with_date_column <- unlist(lapply(sdtms, has_date_col))

date_converter <- function(sdtm) {
  sdtm_num_time <- sdtm
  sdtm_num_time$FROM_DATE <- ymd_hms(sdtm$FROM_DATE)
  end_num <- ymd_hms(sdtm$TO_DATE)
  end_num[is.na(end_num)] <- current_time
  sdtm_num_time$TO_DATE <- end_num
  return(sdtm_num_time)
}

sdtms[sdtms_with_date_column] <- lapply(sdtms[sdtms_with_date_column], date_converter)

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

# Get all times where road characteristics change
time_breaks <- function(sdtm) {
  times_df <- sdtm %>% 
    group_by(ROUTE) %>% 
    summarize(starttimes = unique(c(FROM_DATE, TO_DATE))) %>% 
    arrange(starttimes)
  return(times_df)
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
    mutate(lastrecord = (startpoints == max(startpoints))) %>% 
    filter(!lastrecord) %>% 
    select(-lastrecord)
  # NOTE: Function produces, then discards one record per ROUTE whose starting and
  # ending point are both the biggest ending END_ACCUM value observed for that ROUTE.
  # This record could be useful depending on how certain characteristics are cataloged,
  # but it is removed here. If you want to un-remove it, just take out the last two
  # filter and select statements (and the pipe above them).
  
  all_starttimes <- bind_rows(lapply(sdtms[sdtms_with_date_column], time_breaks)) %>% 
    group_by(ROUTE) %>% 
    summarize(starttimes = unique(starttimes)) %>% 
    arrange(ROUTE, starttimes)
  
  endtimes <- c(all_starttimes$starttimes[-1],
                all_starttimes$starttimes[length(all_starttimes$starttimes)])
  
  all_times <- bind_cols(all_starttimes, endtimes = endtimes) %>% 
    group_by(ROUTE) %>% 
    mutate(lastrecord = (starttimes == max(starttimes))) %>% 
    filter(!lastrecord) %>% 
    select(-lastrecord)
  
  shell <- full_join(all_segments, all_times, by = c("ROUTE")) %>% 
    group_by(ROUTE) %>% 
    select(ROUTE, startpoints, endpoints, starttimes, endtimes) %>% 
    arrange(ROUTE, startpoints, starttimes)
  
  # This is a fudge to account for roads that occur only in datasets that don't
  # have date information. There might be a better way to account for this.
  default_start_date <- min(shell$starttimes, na.rm = TRUE)
  
  shell$starttimes[is.na(shell$starttimes)] <- default_start_date
  shell$endtimes[is.na(shell$endtimes)] <- current_time
  
  return(shell)
}

shell <- shell_creator(sdtms)

########### Replicating records through granular road segments and time windows ############
# We first need to add FROM_DATE and TO_DATE information to sdtms that don't already have it.
time_adder <- function(sdtm) {
  # Alfa et Omega
  # Extracts from the shell the earliest date on a ROUTE x startpoints combination and puts
  # the current time as the ending point
  aeto <- shell %>% 
    mutate(START_ACCUM = startpoints) %>% 
    select(-startpoints) %>% 
    group_by(ROUTE, START_ACCUM) %>% 
    summarize(FROM_DATE = min(starttimes),
              TO_DATE = current_time)
  
  # Tacking the earliest and latest date point onto the original sdtm
  with_aeto <- right_join(aeto, sdtm, by = c("ROUTE", "START_ACCUM")) %>% 
    select(ROUTE, START_ACCUM, END_ACCUM, FROM_DATE, TO_DATE, everything()) %>% 
    arrange(ROUTE, START_ACCUM)
  
  return(with_aeto)
}

sdtms[!sdtms_with_date_column] <- lapply(sdtms[!sdtms_with_date_column], time_adder)

# This segment of the program is essentially reconstructing a SAS DATA step with a RETAIN
# statement to populate through rows.

# We start by left joining the shell to each sdtm dataset and populating through all the
# records whose union is the space interval x time interval we have data on. We then get rid
# of the original START_ACCUM, END_ACCUM, FROM_DATE, and TO_DATE variables, since
# startpoints, endpoints, starttime, and endtime have all the information we need.

shell_join <- function(sdtm) {
  joinable_sdtm <- sdtm %>% 
    mutate(startpoints = START_ACCUM,
           starttimes = FROM_DATE)
  with_shell <- left_join(shell, joinable_sdtm, by = c("ROUTE", "startpoints", "starttimes")) %>% 
    select(ROUTE, startpoints, endpoints, starttimes, endtimes,
           START_ACCUM, END_ACCUM, FROM_DATE, TO_DATE, everything()) %>% 
    arrange(ROUTE, startpoints, starttimes)
  
  # pdv contains everyting besides ROUTE, startpoints, endpoints, starttimes, and endtimes
  # The idea is to reset the pdv everytime we get a new sdtm record, and then copy the 
  # pdv into every row where the window referred to is completely contained in the
  # START_ACCUM to END_ACCUM window and the FROM_DATE to TO_DATE window.
  pdv <- with_shell[1, 6:ncol(with_shell)]
  for (i in 1:nrow(with_shell)) {
    if (!is.na(with_shell$START_ACCUM[i]) & !is.na(with_shell$FROM_DATE[i]) &
        with_shell$START_ACCUM[i] == with_shell$startpoints[i] &
        with_shell$FROM_DATE[i] == with_shell$starttimes[i]) {
      pdv <- with_shell[i, 6:ncol(with_shell)]
    }
    
    if (!is.na(pdv$START_ACCUM) & !is.na(pdv$END_ACCUM) &
        !is.na(pdv$FROM_DATE) & !is.na(pdv$TO_DATE) &
        pdv$START_ACCUM <= with_shell$startpoints[i] &
        with_shell$endpoints[i] <= pdv$END_ACCUM &
        pdv$FROM_DATE <= with_shell$starttimes[i] &
        with_shell$endtimes[i] <= pdv$TO_DATE) {
      with_shell[i, 6:ncol(with_shell)] <- pdv
    }
  }
  
  with_shell <- with_shell %>% 
    select(-START_ACCUM, -END_ACCUM, -FROM_DATE, -TO_DATE,
           # This part is done to avoid merge issues when everything gets
           # joined later
           -endpoints, -endtimes)
  
  return(with_shell)
}

joined_populated <- lapply(sdtms, shell_join)

# To be removed tomorrow, probably (just for accounting for duplicates, which is now
# done in preprocessing)
#joined_populated[[2]] <- joined_populated[[2]] %>% 
#  filter(!(OBJECTID %in% c(15764, 14553, 14990, 19615, 19926)))

################# Merging joined and populated sdtms together; formatting ##################
RC <- c(list(shell), joined_populated) %>% 
  reduce(left_join, by = c("ROUTE", "startpoints", "starttimes")) %>% 
  mutate(START_ACCUM = startpoints,
         END_ACCUM = endpoints) %>% 
  select(-startpoints, -endpoints) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, starttimes, endtimes, everything()) %>% 
  arrange(ROUTE, START_ACCUM, starttimes)

######### Compressing by combining rows identical except for identifying information ########
# First across time
RC <- RC %>% 
  mutate(dropflag = FALSE) %>% 
  select(ROUTE, START_ACCUM, END_ACCUM, starttimes, endtimes, everything()) %>% 
  arrange(ROUTE, START_ACCUM, starttimes)


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
  select(ROUTE, starttimes, endtimes, START_ACCUM, END_ACCUM, everything()) %>% 
  arrange(ROUTE, starttimes, START_ACCUM)

# Adjust road segment boundaries and flag duplicate rows for deletion.
for (i in 1:(nrow(RC) - 1)) {
  if (identical(RC[i, -c(4, 5)], RC[i + 1, -c(4, 5)])) {
    RC$dropflag[i] <- TRUE
    RC[i + 1, 4] <- RC[i, 4]
  }
}

RC <- RC %>% 
  filter(dropflag == FALSE) %>% 
  select(ROUTE, starttimes, endtimes, START_ACCUM, END_ACCUM, everything()) %>% 
  select(-dropflag) %>% 
  arrange(ROUTE, starttimes, START_ACCUM)

# Write to output
write.csv(RC, file = "/home/bdahl00/udot-data/road-characteristics.csv")