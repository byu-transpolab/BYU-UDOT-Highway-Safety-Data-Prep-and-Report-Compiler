library(tidyverse)
# library(sf)

# Set filepath and Column Names

routes.filepath <- "data/UDOT_Routes_ALRS.csv"
routes.columns <- c("ROUTE_ID",                          
                "BEG_MILEAGE",                             
                "END_MILEAGE")

# aadt.filepath <- "data/shapefile/AADT_Unrounded.shp"
aadt.filepath <- "data/AADT_Unrounded.csv"
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
                  "CUTRK2014")

# fc.filepath <- "data/shapefile/Functional_Class_ALRS.shp"
# fc.columns <- c("ROUTE_ID",                          
#                 "FROM_MEASU",                             
#                 "TO_MEASURE",
#                 "FUNCTIONAL",                             
#                 "RouteDir",                             
#                 "RouteType")
fc.filepath <- "data/Functional_Class_ALRS.csv"
fc.columns <- c("ROUTE_ID",                          
                "FROM_MEASURE",                             
                "TO_MEASURE",
                "FUNCTIONAL_CLASS",                             
                "RouteDir",                             
                "RouteType")

# speed.filepath <- "data/shapefile/UDOT_Speed_Limits_2019.shp"
# speed.columns <- c("ROUTE_ID",                         
#                    "FROM_MEASU",                         
#                    "TO_MEASURE",                         
#                    "SPEED_LIMI")
speed.filepath <- "data/UDOT_Speed_Limits_2019.csv"
speed.columns <- c("ROUTE_ID",
                   "FROM_MEASURE",
                   "TO_MEASURE",
                   "SPEED_LIMIT")

# lane.filepath <- "data/shapefile/Lanes.shp"
# lane.columns <- c("ROUTE",
#                   "START_ACCU",
#                   "END_ACCUM",
#                   "THRU_CNT",
#                   "THRU_WDTH")
lane.filepath <- "data/Lanes.csv"
lane.columns <- c("ROUTE",
                  "START_ACCUM",
                  "END_ACCUM",
                  "THRU_CNT",
                  "THRU_WDTH")

# small <- read_sf("data/shapefile/Urban_Boundaries_Small.shp")
# small.filepath <- "data/shapefile/Urban_Boundaries_Small.shp"
# small.columns <- c("NAME",                         
#                    "TYPE_",
#                    "geometry")

# large <- read_sf("data/shapefile/Urban_Boundaries_Large.shp")
# large.filepath <- "data/shapefile/Urban_Boundaries_Large.shp"
# large.columns <- c("NAME",                         
#                    "TYPE_",
#                    "geometry")

# intersection <- read_sf("data/shapefile/Intersections.shp")

# driveway <- read_sf("data/shapefile/Driveway.shp")
# driveway.filepath <- "data/shapefile/Driveway.shp"
# driveway.columns <- c("ROUTE",
#                       "START_ACCU",
#                       "END_ACCUM",
#                       "DIRECTION",
#                       "TYPE",
#                       "geometry",
#                       "BEG_LONG",
#                       "BEG_LAT",
#                       "END_LONG",
#                       "END_LAT")

# median <- read_sf("data/shapefile/Medians.shp")
# median.filepath <- "data/shapefile/Medians.shp"
# median.columns <- c("ROUTE_NAME",
#                     "TRAVEL_DIR",
#                     "START_ACCUM",
#                     "END_ACCUM",
#                     "MEDIAN_TYP",
#                     "TRFISL_TYP",
#                     "MDN_PRTCTN",
#                     "BEG_LONG",
#                     "BEG_LAT",
#                     "END_LONG",
#                     "END_LAT")

# shoulder <- read_sf("data/shapefile/Shoulders.shp")
# shoulder.filepath <-"data/shapefile/Shoulders.shp"
# shoulder.columns <- c("ROUTE",
#                       "TRAVEL_DIR",
#                       "START_ACCU",
#                       "END_ACCUM",
#                       "UTPOSITION",
#                       "SHLDR_WDTH",
#                       "BEG_LONG",
#                       "BEG_LAT",
#                       "END_LONG",
#                       "END_LAT")

## Read in file function
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

read_filez_csv <- function(filepath, columns) {
  if (str_detect(filepath, ".csv")) {
    print("reading csv")
    read_csv(filepath) %>% select(all_of(columns))
  }
  else {
    print("Error reading in:")
    print(filepath)
  }
}

# Compress Segments Function
compress_seg <- function(df, col, variables) {
  # drop geometry (it will become useless after this anyways)
  # df <- df %>% st_drop_geometry()
  # get columns to preserve (assumes the first three are "route", "beg_mp", "end_mp")
  col <- tail(col, -3)
  # sort by route and milepoints (assumes consistent naming convention for these)
  df <- df %>%
    arrange(ROUTE, BEG_MP)
  # loop through rows checking the variables to merge by
  iter <- 0
  count <- 0
  route <- -1
  value <- variables
  df$ID <- 0
  for(i in 1:nrow(df)){
    new_route <- df$ROUTE[i]
    test = 0
    for (j in 1:length(variables)){      # test each of the variables for if they are unique from the prev row
      varName <- variables[j]
      new_value <- df[[varName]][i]
      if(is.na(new_value)){              # treat NA as zero to avoid errors
        new_value = 0
      }
      if(new_value != value[j]){         # set test = 1 if any of the variables are unique from prev row
        value[j] <- new_value
        test = 1
      }
    }
    if((new_route != route) | (test == 1)){    # create new ID ("iter") if test=1 or there is a unique route
      iter <- iter + 1
      route <- new_route
    } else {
      count = count + 1
    }
    df$ID[i] <- iter
  }
  # use summarize to compress segments
  df <- df %>% 
    group_by(ID) %>%
    summarise(
      ROUTE = unique(ROUTE),
      BEG_MP = min(BEG_MP),
      END_MP = max(END_MP), 
      across(.cols = col)
    ) %>%
    unique() %>%
    ungroup() %>%
    select(-ID)
  # report the number of combined rows
  print(paste("combined", count, "rows"))
  # return df
  return(df)
} 

# Convert AADT to Numeric Function
aadt_numeric <- function(aadt, aadt.columns){
  # coerce aadt to numeric
  for(i in 1:length(aadt.columns)){
    string <- substr(aadt.columns[i],1,4)
    if(string == "AADT" | string == "SUTR" | string == "CUTR"){
      aadt[i] <- as.numeric(unlist(aadt[i]))
    } 
  }
  return(aadt)
}

# Pivot AADT longer Function
pivot_aadt <- function(aadt){
  aadt <- aadt %>%
    pivot_longer(
      cols = starts_with("AADT") | starts_with("SUTRK") | starts_with("CUTRK"),
      names_to = "count_type",
      values_to = "count"
    ) %>%
    mutate(
      YEAR = as.integer(gsub(".*?([0-9]+).*", "\\1", count_type)),   # extract numeric characters
      count_type = sub("^([[:alpha:]]*).*", "\\1", count_type)     # extract alpha-beta characters
    ) %>%
    pivot_wider(
      names_from = count_type,
      values_from = count,
    )
  return(aadt)
}

# Fix last ending milepoints
fix_endpoints <- function(df, routes){  
  error_count <- 0
  error_count2 <- 0
  count <- 0
  for(i in 1:nrow(df)){
    end <- df[["END_MP"]][i]
    nxt <- df[["END_MP"]][i+1]
    # check if the next segment endpoint is less than the current
    if(end > nxt | is.na(nxt)){
      rt <- df[["ROUTE"]][i]
      rt2 <- df[["ROUTE"]][i+1]
      # check if the segments in question are on the same route. If they are, skip the iteration and count error
      if(rt == rt2 & !is.na(rt2)){
        error_count <- error_count + 1
        next
      }
      # assign endpoint value from the endpoint of the routes data
      rt_row <- which(routes$ROUTE == rt)
      end_rt <- routes[["END_MP"]][rt_row]
      # if the endpoints are not the same, correct the endpoint
      if(end_rt < end){
        prev <- df[["END_MP"]][i-1]
        # check if the previous endpoint is greater that the current and assign error if it is and skip the iteration
        if(end_rt < prev & !is.na(prev)){
          error_count2 <- error_count2 + 1
          print(paste(rt, prev, end))
          df[["END_MP"]][i-1]
          next
        }
        # assign new endpoint value to the segment
        df[["END_MP"]][i] <- end_rt
        count <- count + 1
      }
    }
  }
  # print number of iterations and any errors
  print(paste("corrected", count, "endpoints"))
  if(error_count > 0){
    print(paste("WARNING, the segments overlap at", error_count, "places"))
  }
  if(error_count2 > 0){
    print(paste("WARNING,", error_count2, "segments exist outside of the LRS. Corrected data should be obtained from UDOT."))
  }
  return(df)
}

###
## Routes Data Prep
###

# Read in routes File
routes <- read_filez_csv(routes.filepath, routes.columns)

# Standardize Column Names
names(routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
routes <- routes %>% filter(grepl("M", ROUTE))





###
## Functional Class Data Prep
###

# Read in fc File
fc <- read_filez_csv(fc.filepath, fc.columns)

# Standardize Column Names
names(fc)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
fc <- fc %>% filter(grepl("M", ROUTE))

# Select only State Routes
fc <- fc %>% filter(grepl("State", RouteType))

# Find Number of Unique Routes in the fc File
num.fc.routes <- fc %>% pull(ROUTE) %>% unique() %>% length()
main.routes <- as.character(fc %>% pull(ROUTE) %>% unique() %>% sort())

# Compress fc
fc <- compress_seg(fc, fc.columns, c("FUNCTIONAL_CLASS"))

# # Create routes dataframe
# routes <- fc %>% select(ROUTE, BEG_MP, END_MP)
# routes <- compress_seg(fc, c("ROUTE", "BEG_MP", "END_MP"), c("ROUTE"))

# fix ending endpoints
fc <- fix_endpoints(fc, routes)

# Unused Code for Filtering fc Data

# fctest <- fc
# fctest %>%
#   group_by(ROUTE) %>%
#   mutate(
#     New_fc = case_when(
#       row_number() == 1L ~ fc,
#       is.na(fc) & sum(fc) != rollsum(fc, k=as.scalar(row_number()), align='left') ~ lead(fc, n=1L)
#     )
#   )

#if row(i) ROUTE, Functional, RouteDir, RouteType = row(i+1)
#  BEG_MP(i), END_MP(i+1) , ROUTE, Functional, RouteDir, RouteType
#else
#End if 

# Take First Four Numbers of Route Column
# fc$ROUTE <- substr(fc$ROUTE, 1, 4)

# Making duplicate dataset for positive and negative sides of road
# fc_pos <- fc_neg <- aadt
# fc_pos$ROUTE <- as.character(paste(fc_pos$ROUTE, "P", sep = ""))
# fc_neg$ROUTE <- as.character(paste(fc_neg$ROUTE, "N", sep = ""))
# fc <- rbind(fc_pos, fc_neg)

####
## AADT Data Prep
####

# Read in aadt File
aadt <- read_filez_csv(aadt.filepath, aadt.columns)

# Standardize Column Names
names(aadt)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Get Only Main Routes
aadt <- aadt %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in aadt file
num.aadt.routes <- aadt %>% pull(ROUTE) %>% unique() %>% length()

# Convert aadt to numeric
aadt <- aadt_numeric(aadt, aadt.columns)

# Compress aadt
aadt <- compress_seg(aadt, aadt.columns, tail(aadt.columns, -3))

# fix ending endpoints
aadt <- fix_endpoints(aadt, routes)

# Fix missing negative direction AADT values
routes <- fc %>% select(ROUTE, BEG_MP, END_MP)
routes <- compress_seg(fc, c("ROUTE", "BEG_MP", "END_MP"), c("ROUTE"))
routes_n <- routes %>% filter(grepl("N",ROUTE))

aadt_p <- aadt %>% filter(grepl("P",ROUTE))

aadt_n <- aadt %>% filter(grepl("N",ROUTE))

test <- full_join(aadt_n, routes_n, by = "ROUTE") %>%
  mutate(
    label = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE)),
    BEG_MP = BEG_MP.x,
    END_MP = END_MP.x
  ) %>%
  select(-contains(".y"), -contains(".x"))




colns <- tail(aadt.columns, -3)

aadt_check <- RC %>%
  filter(grepl("N",ROUTE))

for(i in 1:nrow(aadt_check)){
  fix <- 1
  for(j in 1:length(colns)){
    if(!is.na(aadt_check[[colns[j]]][i])){
      fix <- 0
    }
  }
}

# Unused Code for Filtering aadt Data

# Take First Four Numbers of Route Column
# aadt$ROUTE <- substr(aadt$ROUTE, 1, 4)

# Making duplicate dataset for positive and negative sides of road
# aadt_pos <- aadt_neg <- aadt
# aadt_pos$ROUTE <- as.character(paste(aadt_pos$ROUTE, "P", sep = ""))
# aadt_neg$ROUTE <- as.character(paste(aadt_neg$ROUTE, "N", sep = ""))
# aadt <- rbind(aadt_pos, aadt_neg)

###
## Speed Limits Data Prep
###

# Read in speed File
speed <- read_filez_csv(speed.filepath, speed.columns)
# speed <- read_filez_csv(speed.filepath, speed.columns)

# Standardizing Column Names
names(speed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Getting Only Main Routes
speed <- speed %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in speed file
num.speed.routes <- speed %>% pull(ROUTE) %>% unique() %>% length()

# Compress speed
speed <- compress_seg(speed, speed.columns, tail(speed.columns, -3))

# fix ending endpoints
speed <- fix_endpoints(speed, routes)

# Unused Code for Filtering speed Data

# Take First Four Numbers of Route Column
# speed$ROUTE <- substr(speed$ROUTE, 1, 4)

# Making duplicate dataset for positive and negative sides of road
# sl_pos <- sl_neg <- speed
# sl_pos$ROUTE <- paste(sl_pos$ROUTE, "+", sep = "")
# sl_neg$ROUTE <- paste(sl_neg$ROUTE, "-", sep = "")
# speed <- rbind(sl_pos, sl_neg)

###
## Lanes Data Prep
###

# Read in lane file
lane <- read_filez_csv(lane.filepath, lane.columns)

# Standardize Column Names
names(lane)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Add M to Route Column to Standardize Route Format
lane$ROUTE <- paste(substr(lane$ROUTE, 1, 6), "M", sep = "")

# Get Only Main Routes
lane <- lane %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in lane file
num.lane.routes <- lane %>% pull(ROUTE) %>% unique() %>% length()

# Compress lanes
lane <- compress_seg(lane, lane.columns, c("THRU_CNT", "THRU_WDTH"))

# fix ending endpoints
lane <- fix_endpoints(lane, routes)

# Unused Code for Filtering lane Data

#Adding direction onto route variable, taking route direction out of its own column
# lane$TRAVEL_DIR <- ifelse(lane$TRAVEL_DIR == "+",
#                       paste(substr(lane$TRAVEL_DIR, 1, 0), "P", sep = ""),
#                       paste(substr(lane$TRAVEL_DIR, 1, 0), "N", sep = ""))

# Trying to define each row as separate segments, no overlap, the directions have different information
# lane %>%
#  group_by(ROUTE, BEG_MP) %>%
#  summarize(should_be_one = n()) %>%
#  filter(should_be_one > 1)
# No duplicates, not all data complete on negative side of the road though

# ###
# ## Intersection Data Prep
# ###
# intersection <- read_filez(intersection.filepath, intersection.columns)
# 
# #Finding all intersections with at least one state route
# intersection <- intersection %>%
#   filter(INT_RT_1 %in% substr(main.routes,1,4) |
#            INT_RT_2 %in% substr(main.routes,1,4) | INT_RT_3 %in% substr(main.routes,1,4) |
#            INT_RT_4 %in% substr(main.routes,1,4))
# #Intersections are labeled at a point, not a segment
# 
# #Adding direction onto route variable, taking route direction out of its own column
# intersection$ROUTE <- paste(substr(intersection$ROUTE,1,4), intersection$TRAVEL_DIR, sep = "")
# intersection <- intersection %>% select(-TRAVEL_DIR)
# 
# #WTHeck
# #intersection %>%
# #  group_by(ROUTE, BEG_MP) %>%
# #  summarize(should_be_one = n()) %>%
# #  filter(should_be_one > 1)
# 
# ###
# ## Shoulder Data Prep
# ###
# shoulder <- read_filez(shoulder.filepath, shoulder.columns)
# 
# #Getting rid of ramps
# shoulder <- shoulder %>% filter(nchar(ROUTE) == 5)
# 
# #Adding direction onto route variable, taking route direction out of its own column
# shoulder$ROUTE <- paste(substr(shoulder$ROUTE,1,4), shoulder$TRAVEL_DIR, sep = "")
# shoulder <- shoulder %>% select(-TRAVEL_DIR)
# 
# #Getting only state routes
# shoulder <- shoulder %>% filter(BEG_MP < END_MP, BEG_LAT < 90, END_LAT < 90)
# 
# shd_disc <- shoulder %>%
#   group_by(ROUTE, BEG_MP, UTPOSITION) %>%
#   mutate(should_be_one = n()) %>%
#   filter(should_be_one > 1) %>%
#   arrange(ROUTE, BEG_MP)
# 
# ###
# ## Median Data Prep
# ###
# median <- read_filez(median.filepath, median.columns)
# names(median)[1] <- "ROUTE"
# 
# #Getting rid of ramps
# median <- median %>% filter(nchar(ROUTE) == 5, BEG_MP < END_MP)
# 
# #Adding direction onto route variable, taking route direction out of its own column
# median$ROUTE <- paste(substr(median$ROUTE,1,4), median$TRAVEL_DIR, sep = "")
# median <- median %>% select(-TRAVEL_DIR)
# 
# #Getting only state routes
# median <- median %>% filter(ROUTE %in% main.routes)
# 
# median %>%
#   group_by(ROUTE, BEG_MP) %>%
#   summarize(should_be_one = n()) %>%
#   filter(should_be_one > 1)
# 
# ###
# ## Driveway Data Prep
# ###
# 
# driveway <- read_filez(driveway.filepath, driveway.columns)
# 
# #Getting rid of ramps
# driveway <- driveway %>% filter(nchar(ROUTE) == 5, BEG_MP < END_MP)
# 
# #Adding direction onto route variable, taking route direction out of its own column
# driveway$ROUTE <- paste(substr(driveway$ROUTE,1,4), driveway$DIRECTION, sep = "")
# driveway <- driveway %>% select(-DIRECTION)
# 
# #Getting only state routes
# driveway <- driveway %>% filter(ROUTE %in% main.routes)
# 
# driveway <- driveway %>%
#   group_by(ROUTE, BEG_MP) %>%
#   mutate(should_be_one = n()) %>%
#   arrange(ROUTE, BEG_MP) %>%
#   select(-contains("LONG")) %>%
#   select(-contains("LAT")) %>%
#   mutate(WIDTH = mean(WIDTH), END_MP = max(END_MP), TYPE = "Minor Residential Driveway") %>%
#   distinct()
# 
# driveway %>%
#   group_by(ROUTE, BEG_MP) %>%
#   mutate(should_be_one = n()) %>%
#   filter(should_be_one > 1)

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

#sdtms <- list()

# For illustrative purposes; to be removed before code is finished.
#sdtm1 <- read.csv("/home/bdahl00/udot-data/original/AADT_Unrounded.csv", header = TRUE) %>% 
#  mutate(ROUTE = ROUTE_NAME,
#         START_ACCUM = START_ACCU)
#sdtm2 <- read.csv("/home/bdahl00/udot-data/original/UDOT_Speed_Limits_(2019).csv", header = TRUE) %>% 
#  mutate(ROUTE = ROUTE_ID,
#         START_ACCUM = FROM_MEASURE,
#         END_ACCUM = TO_MEASURE)
sdtms <- list(aadt, fc, speed, lane)
sdtms <- lapply(sdtms, as_tibble)

####################### Generating time x distance merge shell #########################
# Intersection data has to be handled separately. But talk to the PhDs about how
# intersection lookup is supposed to happen anyway.

# Get all measurements where a road segment stops or starts
segment_breaks <- function(sdtm) {
  breaks_df <- sdtm %>% 
    group_by(ROUTE) %>% 
    summarize(startpoints = unique(c(BEG_MP, END_MP))) %>% 
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
    mutate(lastrecord = (startpoints == max(startpoints))) # %>%    # This isn't actually getting used
  #  filter(!lastrecord) %>% 
  #  select(-lastrecord)
  # NOTE: Function produces, then discards one record per ROUTE whose starting and
  # ending point are both the biggest ending END_ACCUM value observed for that ROUTE.
  # This record could be useful depending on how certain characteristics are cataloged,
  # but it is removed here. If you want to un-remove it, just take out the last two
  # filter and select statements (and the pipe above them).
  
  shell <- all_segments %>% 
    group_by(ROUTE) %>% 
    select(ROUTE, startpoints, endpoints) %>%
    filter(endpoints != 0) %>%
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
    mutate(startpoints = BEG_MP)
  with_shell <- left_join(shell, joinable_sdtm, by = c("ROUTE", "startpoints")) %>% 
    select(ROUTE, startpoints, endpoints,
           BEG_MP, END_MP, everything()) %>% 
    arrange(ROUTE, startpoints)
  
  # pdv contains everyting besides ROUTE, startpoints, endpoints
  # The idea is to reset the pdv everytime we get a new sdtm record, and then copy the 
  # pdv into every row where the window referred to is completely contained in the
  # START_ACCUM to END_ACCUM window 
  pdv <- with_shell[1, 4:ncol(with_shell)]
  for (i in 1:nrow(with_shell)) {
    if (!is.na(with_shell$BEG_MP[i]) &
        with_shell$BEG_MP[i] == with_shell$startpoints[i]){
      pdv <- with_shell[i, 4:ncol(with_shell)]
    }
    
    if (!is.na(pdv$BEG_MP) & !is.na(pdv$END_MP) &
        pdv$BEG_MP <= with_shell$startpoints[i] &
        with_shell$endpoints[i] <= pdv$END_MP) {
      with_shell[i, 4:ncol(with_shell)] <- pdv
    }
  }
  
  with_shell <- with_shell %>% 
    select(-BEG_MP, -END_MP,
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
  mutate(BEG_MP = startpoints,
         END_MP = endpoints) %>% 
  select(-startpoints, -endpoints) %>% 
  select(ROUTE, BEG_MP, END_MP, everything()) %>% 
  arrange(ROUTE, BEG_MP)

######### Compressing by combining rows identical except for identifying information ########

# Compressing across road segment
RC <- RC %>% 
  filter(dropflag == FALSE) %>% 
  select(ROUTE, BEG_MP, END_MP, everything()) %>% 
  arrange(ROUTE, BEG_MP)

# Adjust road segment boundaries and flag duplicate rows for deletion.
for (i in 1:(nrow(RC) - 1)) {
  if (identical(RC[i, -c(4, 5)], RC[i + 1, -c(4, 5)])) {
    RC$dropflag[i] <- TRUE
    RC[i + 1, 4] <- RC[i, 4]
  }
}

RC <- RC %>% 
  filter(dropflag == FALSE) %>% 
  select(ROUTE, BEG_MP, END_MP, everything()) %>% 
  select(-dropflag) %>% 
  arrange(ROUTE, BEG_MP)

# Pivot AADT
test <- pivot_aadt(RC)

# Write to output
output <- paste0("data/output/",format(Sys.time(),"%d%b%y_%H.%M"),".csv")
write.csv(RC, file = output)
















# Potential code to replace the shell method

# assign variables
sdtms <- list(aadt, fc, speed, lane)
sdtms <- lapply(sdtms, as_tibble)
df <- rbind(fc, lane, aadt, speed)
# extract columns from data
colns <- colnames(df)
colns <- tail(colns, -4)    # remove ID and first 
# sort by route and milepoints (assumes consistent naming convention for these)
df <- df %>%
  arrange(ROUTE, BEG_MP, END_MP) %>%
  select(ROUTE, BEG_MP, END_MP)
# loop through variables and rows
iter <- 0
count <- 0
route <- -1
value <- variables
df$ID <- 0
# create new variables to replace the old ones



for(i in 1:nrow(df)){
  new_route <- df$ROUTE[i]
  test = 0
  for (j in 1:length(variables)){      # test each of the variables for if they are unique from the prev row
    varName <- variables[j]
    new_value <- df[[varName]][i]
    if(is.na(new_value)){              # treat NA as zero to avoid errors
      new_value = 0
    }
    if(new_value != value[j]){         # set test = 1 if any of the variables are unique from prev row
      value[j] <- new_value
      test = 1
    }
  }
  if((new_route != route) | (test == 1)){    # create new ID ("iter") if test=1 or there is a unique route
    iter <- iter + 1
    route <- new_route
  } else {
    count = count + 1
  }
  df$ID[i] <- iter
}
# use summarize to compress segments
df <- df %>% 
  group_by(ID) %>%
  summarise(
    ROUTE = unique(ROUTE),
    BEG_MP = min(BEG_MP),
    END_MP = max(END_MP), 
    across(.cols = col)
  ) %>%
  unique()
# report the number of combined rows
print(paste("combined", count, "rows"))