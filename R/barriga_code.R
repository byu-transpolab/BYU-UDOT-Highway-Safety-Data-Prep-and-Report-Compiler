###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Set Filepath and Column Names for each Dataset
###

routes.filepath <- "data/csv/UDOT_Routes_ALRS.csv"
routes.columns <- c("ROUTE_ID",                          
                    "BEG_MILEAGE",                             
                    "END_MILEAGE")

aadt.filepath <- "data/csv/AADT_Unrounded.csv"
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

fc.filepath <- "data/csv/Functional_Class_ALRS.csv"
fc.columns <- c("ROUTE_ID",                          
                "FROM_MEASURE",                             
                "TO_MEASURE",
                "FUNCTIONAL_CLASS",                             
                "RouteDir",                             
                "RouteType")

speed.filepath <- "data/csv/UDOT_Speed_Limits_2019.csv"
speed.columns <- c("ROUTE_ID",
                   "FROM_MEASURE",
                   "TO_MEASURE",
                   "SPEED_LIMIT")

lane.filepath <- "data/csv/Lanes.csv"
lane.columns <- c("ROUTE",
                  "START_ACCUM",
                  "END_ACCUM",
                  "THRU_CNT",
                  "THRU_WDTH")

urban.filepath <- "data/csv/Urban_Code.csv"
urban.columns <- c("ROUTE_ID",
                   "FROM_MEASURE",
                   "TO_MEASURE",
                   "URBAN_CODE")

intersection.filepath <- "data/csv/Intersections.csv"
intersection.columns <- c("ROUTE",
                          "START_ACCUM",
                          "END_ACCUM",
                          "ID",
                          "INT_TYPE",
                          "TRAFFIC_CO",
                          "SR_SR",
                          "INT_RT_1",
                          "INT_RT_2",
                          "INT_RT_3",
                          "INT_RT_4",
                          "STATION",
                          "REGION",
                          "BEG_LONG",
                          "BEG_LAT",
                          "BEG_ELEV")


driveway.filepath <- "data/csv/Driveway.csv"
driveway.columns <- c("ROUTE",
                      "START_ACCUM",
                      "END_ACCUM",
                      "DIRECTION",
                      "TYPE")

median.filepath <- "data/csv/Medians.csv"
median.columns <- c("ROUTE_NAME",
                    "START_ACCUM",
                    "END_ACCUM",
                    "MEDIAN_TYP",
                    "TRFISL_TYP",
                    "MDN_PRTCTN")

shoulder.filepath <-"data/csv/Shoulders.csv"
shoulder.columns <- c("ROUTE",
                      "START_ACCUM", # these columns are rounded quite a bit. Should we use Mandli_BMP and Mandli_EMP?
                      "END_ACCUM",
                      "UTPOSITION",
                      "SHLDR_WDTH")

###
## Functions
###

# Read File functions
read_filez_shp <- function(filepath, columns) {
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
  # start timer
  start.time <- Sys.time()
  # sort by route and milepoints (assumes consistent naming convention for these)
  df <- df %>% 
    filter(BEG_MP < END_MP) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>% 
    arrange(ROUTE, BEG_MP)
  # get columns to preserve
  if(missing(col)) {
    col <- colnames(df)
  }
  col <- tail(col, -3)
  # set optional argument
  if(missing(variables)) {
    variables <- col
  }
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
  # record time
  end.time <- Sys.time()
  time.taken <- end.time - start.time
  print(paste("Time taken for code to run:", time.taken))
  # return df
  return(df)
} 

# Alternate Compress Segments Function (created by the stats team)
compress_seg_alt <- function(df) {
  # start timer
  start.time <- Sys.time()
  # sort and filter
  df <- df %>% 
    filter(BEG_MP < END_MP) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>% 
    arrange(ROUTE, BEG_MP)
  # Adjust road segment boundaries and flag duplicate rows for deletion.
  df$dropflag <- FALSE
  count <- 0
  for (i in 1:(nrow(df) - 1)) {
    if (identical(df[i, -c(2, 3)], df[i + 1, -c(2, 3)])) {
      df$dropflag[i] <- TRUE
      df[i + 1, 2] <- df[i, 2]    # change beg_mp to match previous
      count <- count + 1
    }
  }
  # Remove flagged rows
  df <- df %>%
    filter(dropflag == FALSE) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>%
    select(-dropflag) %>%
    arrange(ROUTE, BEG_MP)
  # report the number of combined rows
  print(paste("combined", count, "rows"))
  # record time
  end.time <- Sys.time()
  time.taken <- end.time - start.time
  print(paste("Time taken for code to run:", time.taken))
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

# Function to fix missing negative direction AADT values (can be optimized more)
aadt_neg <- function(aadt, rtes, divd){
  # isolate negative routes
  rtes_n <- rtes %>% filter(grepl("N",ROUTE)) %>% select(-BEG_MP, -END_MP)
  # isolate positive aadt entries
  aadt_p <- aadt %>% filter(grepl("P",ROUTE)) %>%
    mutate(
      label = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE))
    ) %>%
    select(-ROUTE)
  # isolate negative aadt entries
  aadt_n <- aadt %>% filter(grepl("N",ROUTE))
  # join negative routes with negative aadt (there may be a more efficient way to do this)
  df <- full_join(aadt_n, rtes_n, by = "ROUTE") %>%
    mutate(
      label = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE))
    ) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>%
    filter(is.na(BEG_MP))
  # join missing negative routes with positive segments on the same route
  df <- left_join(df, aadt_p, by = "label") %>%
    select(-label, -contains(".x"))
  # remove suffixes 
  col_new <- gsub('.y','',colnames(df))
  colnames(df) <- col_new
  # filter out excess negative segments
  divd_n <- divd %>% filter(grepl("N", ROUTE))
  divd_p <- divd %>% filter(grepl("P", ROUTE))
  df$flag <- FALSE
  for(i in 1:nrow(df)){
    rt <- df[["ROUTE"]][i]
    rt <- substr(rt, start = 1, stop = 4)
    beg_aadt <- df[["BEG_MP"]][i]
    end_aadt <- df[["END_MP"]][i]
    rt_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt & beg_aadt < divd_p$END_MP & end_aadt > divd_p$BEG_MP)
    # print(rt_row)
    # this if statement helps to deal with routes that have multiple divided segments
    if(length(rt_row) == 0L){
      rt_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt)
    } 
    beg_rt <- divd_p[["BEG_MP"]][rt_row]
    end_rt <- divd_p[["END_MP"]][rt_row]
    # correct routes
    if(length(beg_rt) == 0L){        # check if the route is on divd_p
      df[["flag"]][i] <- TRUE
      next
    }
    if(beg_aadt > end_rt[1]){        # flag segments that fall outside route
      df[["flag"]][i] <- TRUE
    }
    if(end_aadt > end_rt[1]){        # correct end mp if it's greater than route
      df[["END_MP"]][i] <- end_rt[1]
    }
    if(end_aadt < beg_rt[1]){        # flag segments that fall outside route
      df[["flag"]][i] <- TRUE
    }
    if(beg_aadt < beg_rt[1]){        # correct beg mp if it's less than route
      df[["BEG_MP"]][i] <- beg_rt[1]
    }
  }
  # filter out flagged rows
  df <- df %>% filter(flag == FALSE) %>% select(-flag)
  # convert positive milepoints to negative milepoints
  for(i in 1:nrow(df)){
    rt <- df[["ROUTE"]][i]
    rt <- substr(rt, start = 1, stop = 4)
    beg_aadt <- df[["BEG_MP"]][i]
    end_aadt <- df[["END_MP"]][i]
    # find segments on divd tables
    prt_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt & beg_aadt < divd_p$END_MP & end_aadt > divd_p$BEG_MP)
    nrt_row <- which(substr(divd_n$ROUTE, start = 1, stop = 4) == rt & beg_aadt < divd_n$END_MP & end_aadt > divd_n$BEG_MP)
    prt_start_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt)[1]
    # find beginning and ending milepoints for each route
    beg_prt <- divd_p[["BEG_MP"]][prt_row]
    end_prt <- divd_p[["END_MP"]][prt_row]
    beg_nrt <- divd_n[["BEG_MP"]][nrt_row]
    end_nrt <- divd_n[["END_MP"]][nrt_row]
    # get start of first segment
    prt_start <- divd_p[["BEG_MP"]][prt_start_row]
    # calculate conversion factor 
    # (converts positive milepoints to negative milepoints)
    c <- (end_nrt - beg_nrt) / (end_prt - beg_prt)
    # convert milepoints (and offset to zero using prt_start)
    df[["BEG_MP"]][i] <- (beg_aadt - prt_start) * c 
    df[["END_MP"]][i] <- (end_aadt - prt_start) * c 
    # fix segment breaks (treat gaps as zero)
    # I had to specify route 89 because there are routes with non-zero gaps which
    # need to be maintained. (mainly route 84 jumps from 42 to 81) If we can come 
    # up with a more robust solution that would be great.
    if(i-1>0){
      if(df[["BEG_MP"]][i] != df[["END_MP"]][i-1] & df[["ROUTE"]][i] == df[["ROUTE"]][i-1] & rt == "0089"){
        offset <- df[["BEG_MP"]][i] - df[["END_MP"]][i-1]
        end_aadt <- df[["END_MP"]][i]
        df[["END_MP"]][i] <- end_aadt - offset
        df[["BEG_MP"]][i] <- df[["END_MP"]][i-1]
      }
    }
  }
  # rbind df to aadt
  df <- rbind(aadt, df) %>%
    arrange(ROUTE, BEG_MP)
  # return df
  return(df)
}

# Fix last ending milepoints
fix_endpoints <- function(df, routes){  
  error_count <- 0
  error_count2 <- 0
  error_count3 <- 0
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
        print(paste("segments overlaps with next:", rt, "at", round(end,3)))
        next
      }
      # assign endpoint value from the endpoint of the routes data
      rt_row <- which(routes$ROUTE == rt)
      end_rt <- routes[["END_MP"]][rt_row]
      # if the endpoints are not the same, correct the endpoint
      if(end_rt != end){
        prev <- df[["END_MP"]][i-1]
        prev_rt <- df[["ROUTE"]][i-1]
        # check if the previous endpoint is greater that the current and assign error if it is and skip the iteration
        if(i == 1){
          # nothing
        } else if(end_rt < prev & rt == prev_rt){
          error_count2 <- error_count2 + 1
          print(paste("segment not on LRS:", rt, "previous endpoint:", round(prev,3), "endpoint:", round(end,3), "route endpoint:", round(end_rt,3)))
          df[["END_MP"]][i-1]
          next
        }
        # check if the endpoint varies greatly from the LRS
        if(abs(end_rt-end) > 0.5){
          error_count3 <- error_count3 + 1
          print(paste("route varies greatly from LRS:", rt, "by", round(end-end_rt,2), "miles."))
          next
        }
        # assign new endpoint value to the segment
        df[["END_MP"]][i] <- end_rt
        count <- count + 1
      }
    }
  }
  # print number of iterations and any errors
  print("")
  print(paste("corrected", count, "endpoints"))
  if(error_count > 0){
    print(paste("WARNING, the segments overlap at", error_count, "places. Overlaps will be discarded in the segmentation code."))
  }
  if(error_count2 > 0){
    print(paste("WARNING, at least", error_count2, "segment(s) appear to exist outside of the LRS. This may also be due to overlapping segments or a great variance from the LRS along with small segments."))
  }
  if(error_count3 > 0){
    print(paste("WARNING,", error_count3, "route(s) vary from the LRS by 0.5 miles or more."))
  }
  return(df)
}

# Function to fill in missing aadt and truck data
fill_missing_aadt <- function(df, colns){
  # ignore first 3 columns
  colns <- tail(colns, -3)
  # loop through columns and convert zeros to NA
  # this assumes all zeros equate to missing data. We should confirm with UDOT.
  for(i in 1:length(colns)){
    col <- colns[i]
    df[[col]][df[[col]] == 0] <- NA
  }
  # Make three lists of columns for aadt, sutrk, and cutrk and sort
  aadt_colns <- grep("AADT", colns, value = TRUE) %>% str_sort()
  sutrk_colns <- grep("SUTRK", colns, value = TRUE) %>% str_sort()
  cutrk_colns <- grep("CUTRK", colns, value = TRUE) %>% str_sort()
  col_tbl <- tibble(aadt_colns, sutrk_colns, cutrk_colns)
  years <- length(aadt_colns)
  # loop through years (except the last year)
  for(y in 1:years){
    # loop through rows
    for(i in 1:nrow(df)){
      # check if first year is blank (we won't fill in missing data if it is)
      aadt_first <- df[[aadt_colns[1]]][i]
      sutrk_first <- df[[sutrk_colns[1]]][i]
      cutrk_first <- df[[cutrk_colns[1]]][i]
      # check first year for aadt, sutrk, and cutrk
      if(!is.na(aadt_first) | !is.na(sutrk_first) | !is.na(cutrk_first)){
        # loop through the three columns of col_tbl
        for(k in 1:ncol(col_tbl)){
          cur_colns <- col_tbl[[k]]
          # loop through columns
          for(j in 1:length(cur_colns)){
            # check if the value is NA
            value <- df[[cur_colns[j]]][i]
            if(is.na(value)){
              # create a list of previous and next segments 
              # (treat as NA if segment isn't on the same route)
              # these variables represent the grid layout with years on the x axis
              # and segments on the y axis
              # [[a1 a2 a3]
              #  [b1 b2 b3]
              #  [c1 c2 c3]]
              a1 <- NA
              a2 <- NA
              a3 <- NA
              b1 <- NA
              b2 <- NA
              b3 <- NA
              c1 <- NA
              c2 <- NA
              c3 <- NA
              # fill top row
              if(df[["ROUTE"]][i] == df[["ROUTE"]][i-1]){
                if(j-1 >= 1){
                  a1 <- df[[cur_colns[j-1]]][i-1]
                }
                a2 <- df[[cur_colns[j]]][i-1]
                if(j+1 <= length(cur_colns)){
                  a3 <- df[[cur_colns[j+1]]][i-1]
                }
              }
              # fill middle row
              if(j-1 >= 1){
                b1 <- df[[cur_colns[j-1]]][i]
              }
              a2 <- df[[cur_colns[j]]][i-1]
              if(j+1 <= length(cur_colns)){
                b3 <- df[[cur_colns[j+1]]][i]
              }
              # fill bottom row
              if(df[["ROUTE"]][i] == df[["ROUTE"]][i+1]){
                if(j-1 >= 1){
                  c1 <- df[[cur_colns[j-1]]][i+1]
                }
                c2 <- df[[cur_colns[j]]][i+1]
                if(j+1 <= length(cur_colns)){
                  c3 <- df[[cur_colns[j+1]]][i+1]
                }
              }
              # estimate b2 using the comprehensive method (best case)
              r1 <- b1/a1
              r2 <- b1/c1
              r3 <- b3/a3
              r4 <- b3/c3
              R1 <- mean(c(r1,r3), na.rm = TRUE)
              R2 <- mean(c(r2,r4), na.rm = TRUE)
              b2 <- mean(c(a2*R1, c2*R2), na.rm = TRUE)
              # take average of previous and next segments aadt (worst case)
              if(is.na(b2) | is.nan(b2)){
                surround <- c(a2, c2)
                b2 <- mean(surround, na.rm = TRUE)
              }
              # take average of previous and next years (really worst case)
              if(is.na(b2) | is.nan(b2)){
                surround <- c(b1, b3)
                b2 <- mean(surround, na.rm = TRUE)
              }
              # assign b2 to the current cell
              df[[cur_colns[j]]][i] <- b2
            }
          }
        }
      }
    }
    # remove first year of data and repeat loop
    # (this allows us to thoroughly fill in missing data because we don't fill in
    # missing data if the first year is blank)
    aadt_colns <- tail(aadt_colns, -1)
    sutrk_colns <- tail(sutrk_colns, -1)
    cutrk_colns <- tail(cutrk_colns, -1)
    col_tbl <- tibble(aadt_colns, sutrk_colns, cutrk_colns)
  }
  # loop through columns and convert NaN to NA
  for(i in 1:length(colns)){
    col <- colns[i]
    df[[col]][is.nan(df[[col]])] <- NA
  }
  # return df
  return(df)
}

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

# This segment of the program is essentially reconstructing a SAS DATA step with a RETAIN
# statement to populate through rows.

# We start by left joining the shell to each sdtm dataset and populating through all the
# records whose union is the space interval we have data on. We then get rid
# of the original START_ACCUM, END_ACCUM variables, since
# startpoints, endpoints have all the information we need.

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

###
## Routes Data Prep
###

# Read in routes File
routes <- read_filez_csv(routes.filepath, routes.columns)

# Standardize Column Names
names(routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
routes <- routes %>% filter(grepl("M", ROUTE))

# Round milepoints to 6 decimal places
routes <- routes %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Find Number of Unique Routes in the routes File
num.routes.routes <- routes %>% pull(ROUTE) %>% unique() %>% length()

# load divided routes data
div_routes <- read_csv("data/csv/DividedRoutesList_20220307_Adjusted.csv") 
names(div_routes)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")
div_routes <- div_routes %>%
  mutate(LENGTH = END_MP - BEG_MP) %>%
  arrange(ROUTE, BEG_MP)

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
fc <- compress_seg(fc)

# # Create routes dataframe
# routes <- fc %>% select(ROUTE, BEG_MP, END_MP)
# routes <- compress_seg(fc, c("ROUTE", "BEG_MP", "END_MP"), c("ROUTE"))

# fix ending endpoints
fc <- fix_endpoints(fc, routes)

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
aadt <- compress_seg(aadt)

# fix negative aadt values
aadt <- aadt_neg(aadt, routes, div_routes)

# fix ending endpoints
aadt <- fix_endpoints(aadt, routes)

# fill in missing data
aadt <- fill_missing_aadt(aadt, aadt.columns)

###
## Speed Limits Data Prep
###

# Read in speed File
speed <- read_filez_csv(speed.filepath, speed.columns)

# Standardizing Column Names
names(speed)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Getting Only Main Routes
speed <- speed %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in speed file
num.speed.routes <- speed %>% pull(ROUTE) %>% unique() %>% length()

# Compress speed
speed <- compress_seg(speed)
# speed <- compress_seg_alt(speed)

# fix ending endpoints
speed <- fix_endpoints(speed, routes)

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
lane <- compress_seg(lane)     # note: the alt function is much slower for lanes

# fix ending endpoints
lane <- fix_endpoints(lane, routes)

###
## Urban Code Data Prep
###

# Read in urban code file
urban <- read_filez_csv(urban.filepath, urban.columns)

# Standardize Column Names
names(urban)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Select only Main Routes
urban <- urban %>% filter(grepl("M", ROUTE))

# Round milepoints to 6 decimal places
urban <- urban %>% 
  mutate(BEG_MP = round(BEG_MP,6), END_MP = round(END_MP,6))

# Get Only Main Routes
urban <- urban %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in urban code file
num.urban.routes <- urban %>% pull(ROUTE) %>% unique() %>% length()

# Compress urban code
urban <- compress_seg(urban)     

# fix ending endpoints
urban <- fix_endpoints(urban, routes)

###
## Intersection Data Prep
###

# Read in intersection file
intersection <- read_filez_csv(intersection.filepath, intersection.columns)

# Standardize Column Names
names(intersection)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Add M to Route Column to Standardize Route Format
intersection$ROUTE <- paste(substr(intersection$ROUTE, 1, 5), "M", sep = "")

# Getting only state routes
intersection <- intersection %>% filter(ROUTE %in% substr(main.routes, 1, 6))

# Find Number of Unique Routes in shoulder file
num.intersection.routes <- intersection %>% pull(ROUTE) %>% unique() %>% length()

#Finding all intersections with at least one state route
intersection <- intersection %>%
  filter(INT_RT_1 %in% substr(main.routes,1,4) |
           INT_RT_2 %in% substr(main.routes,1,4) | INT_RT_3 %in% substr(main.routes,1,4) |
           INT_RT_4 %in% substr(main.routes,1,4))
#Intersections are labeled at a point, not a segment

###
## Shoulder Data Prep
###

# Read in shoulder file
shoulder <- read_filez_csv(shoulder.filepath, shoulder.columns)

# Standardize Column Names
names(shoulder)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Getting rid of ramps
shoulder <- shoulder %>% filter(nchar(ROUTE) == 5)
                                
# Add M to Route Column to Standardize Route Format
shoulder$ROUTE <- paste(substr(shoulder$ROUTE, 1, 6), "M", sep = "")

# Getting only state routes
shoulder <- shoulder %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in shoulder file
num.shoulder.routes <- shoulder %>% pull(ROUTE) %>% unique() %>% length()

# Create Point to Reference Shoulders
shoulder <- shoulder %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))

# filter out seemingly duplicated shoulders for observation
shd_disc <- shoulder %>%
  group_by(ROUTE, BEG_MP, UTPOSITION) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)

###
## Median Data Prep
###

# Read in median file
median <- read_filez_csv(median.filepath, median.columns)

# Standardize Column Names
names(median)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

#Getting rid of ramps
median <- median %>% filter(nchar(ROUTE) == 5)

# Add M to Route Column to Standardize Route Format
median$ROUTE <- paste(substr(median$ROUTE, 1, 6), "M", sep = "")

# Getting only state routes
median <- median %>% filter(ROUTE %in% substr(main.routes, 1, 6)) %>%
  filter(BEG_MP < END_MP)

# Find Number of Unique Routes in median file
num.median.routes <- median %>% pull(ROUTE) %>% unique() %>% length()

# Compress Medians
median <- compress_seg(median, variables = c("MEDIAN_TYP", "TRFISL_TYP", "MDN_PRTCTN"))

# Create Point to Reference Driveways
median <- median %>% 
  mutate(MP = (BEG_MP+END_MP)/2) %>% 
  mutate(Length = (END_MP-BEG_MP))

###
## Driveway Data Prep
###

# Read in driveway file
driveway <- read_filez_csv(driveway.filepath, driveway.columns)

# Standardize Column Names
names(driveway)[c(1:3)] <- c("ROUTE", "BEG_MP", "END_MP")

# Adding M to Route
driveway$ROUTE <- paste(substr(driveway$ROUTE, 1, 6), "M", sep = "")

# Create Point to Reference Driveways
driveway <- driveway %>% mutate(MP = (BEG_MP+END_MP)/2)

# compress points
driveway <- driveway %>% unique()

#filter out seemingly duplicated driveways for observation
drv_disc <- driveway %>%
  group_by(ROUTE, BEG_MP) %>%
  mutate(should_be_one = n()) %>%
  filter(should_be_one > 1) %>%
  arrange(ROUTE, BEG_MP)

########################################################################################
###################################### Aggregation #####################################
########################################################################################

##################################### Data read-in #####################################
# Datasets that we must have
#  sdtm1: AADT Unrounded
#  sdtm2: Functional Class
#  sdtm3: UDOT Speed Limits (2019)
#  sdtm4: Lanes
#  sdtm4: Urban Code

sdtms <- list(aadt, fc, speed, lane, urban)
sdtms <- lapply(sdtms, as_tibble)

####################### Generating merge shell #########################

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
RC$Right_Shoulder_Avg <- 0
RC$Left_Shoulder_Avg <- 0
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
  RC[["Right_Shoulder_Max"]][i] <- max(shoulder[["SHLDR_WDTH"]][r_sho_row])
  RC[["Right_Shoulder_Min"]][i] <- min(shoulder[["SHLDR_WDTH"]][r_sho_row])
  RC[["Left_Shoulder_Max"]][i] <- max(shoulder[["SHLDR_WDTH"]][l_sho_row])
  RC[["Left_Shoulder_Min"]][i] <- min(shoulder[["SHLDR_WDTH"]][l_sho_row])
  # RC[["Right_Shoulder_Avg"]][i] <- shoulder[["SHLDR_WDTH"]][r_sho_row]
  # RC[["Left_Shoulder_Avg"]][i] <- ((shoulder[["SHLDR_WDTH"]][l_sho_row]*shoulder[["Length"]][l_sho_row])/(RC[["END_MP"]][i]-RC[["BEG_MP"]][i]))
}
# record time
end.time <- Sys.time()
time.taken <- end.time - start.time
print(paste("Time taken for code to run:", time.taken))

# Pivot AADT
RC <- pivot_aadt(RC)

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
output <- paste0("data/output/",format(Sys.time(),"%d%b%y_%H.%M"),".csv")
write.csv(RC, file = output)

# # Potential code to replace the shell method
# 
# # assign variables
# sdtms <- list(aadt, fc, speed, lane)
# sdtms <- lapply(sdtms, as_tibble)
# df <- rbind(fc, lane, aadt, speed)
# # extract columns from data
# colns <- colnames(df)
# colns <- tail(colns, -4)    # remove ID and first 
# # sort by route and milepoints (assumes consistent naming convention for these)
# df <- df %>%
#   arrange(ROUTE, BEG_MP, END_MP) %>%
#   select(ROUTE, BEG_MP, END_MP)
# # loop through variables and rows
# iter <- 0
# count <- 0
# route <- -1
# value <- variables
# df$ID <- 0
# # create new variables to replace the old ones
# 
# 
# 
# for(i in 1:nrow(df)){
#   new_route <- df$ROUTE[i]
#   test = 0
#   for (j in 1:length(variables)){      # test each of the variables for if they are unique from the prev row
#     varName <- variables[j]
#     new_value <- df[[varName]][i]
#     if(is.na(new_value)){              # treat NA as zero to avoid errors
#       new_value = 0
#     }
#     if(new_value != value[j]){         # set test = 1 if any of the variables are unique from prev row
#       value[j] <- new_value
#       test = 1
#     }
#   }
#   if((new_route != route) | (test == 1)){    # create new ID ("iter") if test=1 or there is a unique route
#     iter <- iter + 1
#     route <- new_route
#   } else {
#     count = count + 1
#   }
#   df$ID[i] <- iter
# }
# # use summarize to compress segments
# df <- df %>% 
#   group_by(ID) %>%
#   summarise(
#     ROUTE = unique(ROUTE),
#     BEG_MP = min(BEG_MP),
#     END_MP = max(END_MP), 
#     across(.cols = col)
#   ) %>%
#   unique()
# # report the number of combined rows
# print(paste("combined", count, "rows"))