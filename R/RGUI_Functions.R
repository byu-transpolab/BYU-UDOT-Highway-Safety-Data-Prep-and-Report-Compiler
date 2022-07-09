###
## Load in Libraries
###

library(tidyverse)
library(dplyr)
library(sf)

###
## Functions
###

# Read in Function
read_csv_file <- function(filepath, columns) {
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
  df$ID <- 0
  for(i in 2:nrow(df)){
    new_route <- df$ROUTE[i]
    test = 0
    for (j in 1:length(variables)){      # test each of the variables for if they are unique from the prev row
      varName <- variables[j]
      new_value <- df[[varName]][i]
      prev_value <- df[[varName]][i-1]
      if(is.na(new_value)){              # treat NA as unique string to avoid errors
        new_value = "N/A"
      }
      if(is.na(prev_value)){           
        prev_value = "N/A"
      }
      if(new_value != prev_value){         # set test = 1 if any of the variables are unique from prev row
        test = 1
      }
    }
    if((new_route != route) |               # create new ID ("iter") if there is a unique route
       (test == 1) |                        # or test = 1 (unique row)
       (df$BEG_MP[i] != df$END_MP[i-1])){   # or there is a gap between segments
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
  print(paste("Time taken for compress segments to run:", time.taken))
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
    if (identical(df[i, -c(2, 3)], df[i + 1, -c(2, 3)]) &   # check for identical rows except for milepoints
        df$END_MP[i] >= df$BEG_MP[i+1]) {                   # and check that there is no gap between segments
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
  print(paste("Time taken for compress segments to run:", time.taken))
  # return df
  return(df)
}

# Function to compress overlapping segments
compress_seg_ovr <- function(df) {
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
    if (df$END_MP[i] > df$BEG_MP[i+1] &
        df$ROUTE[i] == df$ROUTE[i+1]) {        # check for overlaps
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
  print(paste("Time taken for compress segments to run:", time.taken))
  # return df
  return(df)
}

# Function to compress segments by segment length
compress_seg_len <- function(df, len) {
  # start timer
  start.time <- Sys.time()
  # check that length is reasonable
  if(len >= 1){stop("this length may be too long and may damage the data. ending function")}
  # sort and filter
  df <- df %>% 
    filter(BEG_MP < END_MP) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>% 
    arrange(ROUTE, BEG_MP)
  # Adjust road segment boundaries and flag duplicate rows for deletion.
  df$dropflag <- FALSE
  count <- 0
  for (i in 1:(nrow(df) - 1)) {
    # check the segment length
    if (df$END_MP[i] - df$BEG_MP[i] < len){     
      df$dropflag[i] <- TRUE
      # check that there is no gap after seg and the routes match
      if(df$END_MP[i] == df$BEG_MP[i+1] &
         df$ROUTE[i] == df$ROUTE[i+1]){   
        # count number of columns that match
        match <- 0
        for(j in 1:ncol(df)){
          # print(paste(df[i,j], df[i+1,j]))
          this <- df[i,j]
          that <- df[i+1,j]
          if(is.na(this)){this = "N/A"}     # treat NA as unique string to avoid errors
          if(is.na(that)){that = "N/A"}
          if(this == that){
            match <- match + 1
          }
        }
        # determine percent matching (excluding first three columns since we know routes match and milepoints don't)
        perc <- (match-1)/(ncol(df)-3) 
        # if percent matching is at least 50% or the previous route doesn't match or the previous route was also too small
        if(perc >= 0.5 | df$ROUTE[i] != df$ROUTE[i-1] | k > 1){                        
          df[i + 1, 2] <- df[i, 2]    # change next beg_mp to match this beg_mp
        } else{
          df[i - k, 3] <- df[i, 3]    # change previous end_mp to match this end_mp
        }
      # if there is no gap before and routes don't change before
      } else if(df$BEG_MP[i] == df$END_MP[i-k] &
                df$ROUTE[i] == df$ROUTE[i-k]){                                  
        df[i - k, 3] <- df[i, 3]      # change previous end_mp to match this end_mp
      } else{
        df$dropflag[i] <- FALSE
        count <- count - 1
        k <- 0
      }
      count <- count + 1              # count flagged rows
      k <- k + 1                      # we need to keep track of adjacent small segments
    } else{
      k <- 1        
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
  print(paste("Time taken for compress segments to run:", time.taken))
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
      values_fn = first
    ) %>%
    mutate(
      NUM_TRUCKS = (SUTRK + CUTRK) * AADT
    ) %>%
    select(-SUTRK,-CUTRK)
  return(aadt)
}

# Pivot AADT longer modified for intersections
pivot_aadt_int <- function(aadt){
  aadt <- aadt %>%
    pivot_longer(
      cols = contains("AADT") | contains("SUTRK") | contains("CUTRK"),
      names_to = "count_type",
      values_to = "count"
    ) %>%
    mutate(
      YEAR = as.integer(gsub(".*?([0-9]+).*", "\\1", count_type)),   # extract numeric characters
      count_type = gsub("[0-9]+", "", count_type)     # extract everything else
    ) %>%
    pivot_wider(
      names_from = count_type,
      values_from = count,
      values_fn = first
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

# Fix last ending milepoints (This is also a backchecker for the input files)
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


# Add Crash Attributes Function
add_crash_attribute <- function(att_name, segs, csh, return_max_min){
  # optional arguments
  if(missing(return_max_min)){return_max_min <- FALSE}
  # set up dataframe for joining
  csh <- csh %>% select(seg_id, crash_year, !!att_name)
  att_name <- sym(att_name)
  # This function will work differently if we want to just return max,min,avg
  if(return_max_min == FALSE){
    # pivot wider
    csh <- csh %>%
      group_by(seg_id, crash_year, !!att_name) %>% 
      mutate(n = n()) %>% 
      ungroup() %>%
      pivot_wider(id_cols = c(seg_id, crash_year),
                  names_from = !!att_name,
                  names_prefix = paste0(att_name,"_"),
                  values_from = n,
                  values_fill = 0,
                  values_fn = first)
  } else{
    # mutate to get max, min, avg
    csh <- csh %>%
      group_by(seg_id, crash_year) %>% 
      mutate(
        !!sym(paste0("max_",att_name)) := max(!!att_name, na.rm = TRUE),
        !!sym(paste0("min_",att_name)) := min(!!att_name, na.rm = TRUE),
        !!sym(paste0("avg_",att_name)) := mean(!!att_name, na.rm = TRUE)
      ) %>%
      ungroup() %>%
      mutate(
        !!sym(paste0("max_",att_name)) := ifelse(is.infinite(!!sym(paste0("max_",att_name))), NA, !!sym(paste0("max_",att_name))),
        !!sym(paste0("min_",att_name)) := ifelse(is.infinite(!!sym(paste0("min_",att_name))), NA, !!sym(paste0("min_",att_name))),
        !!sym(paste0("avg_",att_name)) := ifelse(is.infinite(!!sym(paste0("avg_",att_name))), NA, !!sym(paste0("avg_",att_name)))
      ) %>%
      select(-!!att_name) %>%
      unique()
  }
  # join to segments
  df <- left_join_fill(segs, csh, by = c("SEG_ID"="seg_id", "YEAR"="crash_year"), fill = 0L)
  # return segments
  return(df)
}

# Add Crash Attributes Function for Intersections
add_crash_attribute_int <- function(att_name, ints, csh, return_max_min){
  # optional arguments
  if(missing(return_max_min)){return_max_min <- FALSE}
  # set up dataframe for joining
  csh <- csh %>% select(int_id, crash_year, !!att_name)
  att_name <- sym(att_name)
  # This function will work differently if we want to just return max,min,avg
  if(return_max_min == FALSE){
    # pivot wider
    csh <- csh %>%
      group_by(int_id, crash_year, !!att_name) %>% 
      mutate(n = n()) %>% 
      ungroup() %>%
      pivot_wider(id_cols = c(int_id, crash_year),
                  names_from = !!att_name,
                  names_prefix = paste0(att_name,"_"),
                  values_from = n,
                  values_fill = 0,
                  values_fn = first)
  } else{
    # mutate to get max, min, avg
    csh <- csh %>%
      group_by(int_id, crash_year) %>%
      mutate(
        !!sym(paste0("max_",att_name)) := max(!!att_name, na.rm = TRUE),
        !!sym(paste0("min_",att_name)) := min(!!att_name, na.rm = TRUE),
        !!sym(paste0("avg_",att_name)) := mean(!!att_name, na.rm = TRUE)
      ) %>%
      ungroup() %>%
      mutate(
        !!sym(paste0("max_",att_name)) := ifelse(is.infinite(!!sym(paste0("max_",att_name))), NA, !!sym(paste0("max_",att_name))),
        !!sym(paste0("min_",att_name)) := ifelse(is.infinite(!!sym(paste0("min_",att_name))), NA, !!sym(paste0("min_",att_name))),
        !!sym(paste0("avg_",att_name)) := ifelse(is.infinite(!!sym(paste0("avg_",att_name))), NA, !!sym(paste0("avg_",att_name)))
      ) %>%
      select(-!!att_name) %>%
      unique()
  }
  # join to segments
  ints <- left_join_fill(ints, csh, by = c("Int_ID"="int_id", "YEAR"="crash_year"), fill = 0L)
  # return segments
  return(ints)
}


# Add roadway attributes to intersections function
add_int_att <- function(int, att, is_aadt){
  # optional arguments
  if(missing(is_aadt)){is_aadt <- FALSE} else if(is_aadt == TRUE){
    int <- int %>% select(-contains("AADT"), -contains("SUTRK"), -contains("CUTRK"))
  }
  # create row id column for attribute data
  att <- att %>% rownames_to_column("att_rw")
  att$att_rw <- as.integer(att$att_rw)
  # identify attribute segments on intersection file
  for(k in 1:4){               # since we can expect intersections to have overlapping attributes, we need to identify all of these
    int[[paste0("att_rw_",k)]] <- NA
  }
  for (i in 1:nrow(int)){
    rt1 <- int$INT_RT_1[i]
    mp1 <- int$INT_RT_1_M[i]
    rt2 <- int$INT_RT_2[i]
    mp2 <- int$INT_RT_2_M[i]
    rt3 <- int$INT_RT_3[i]
    mp3 <- int$INT_RT_3_M[i]
    rt4 <- int$INT_RT_4[i]
    mp4 <- int$INT_RT_4_M[i]
    rw1 <- which(att$ROUTE == rt1 & 
                   att$BEG_MP < mp1 & 
                   att$END_MP > mp1)
    rw2 <- which(att$ROUTE == rt2 & 
                   att$BEG_MP < mp2 & 
                   att$END_MP > mp2)
    rw3 <- which(att$ROUTE == rt3 & 
                   att$BEG_MP < mp3 & 
                   att$END_MP > mp3)
    rw4 <- which(att$ROUTE == rt4 & 
                   att$BEG_MP < mp4 & 
                   att$END_MP > mp4)
    if(length(rw1) > 0){
      int$att_rw_1[i] <- att$att_rw[rw1]
    }
    if(length(rw2) > 0){
      int$att_rw_2[i] <- att$att_rw[rw2]
    }
    if(length(rw3) > 0){
      int$att_rw_3[i] <- att$att_rw[rw3]
    }
    if(length(rw4) > 0){
      int$att_rw_4[i] <- att$att_rw[rw4]
    }
    # for(n in 1:4){
    #   rw <- sym(paste0("rw",n))
    #   if(length(rw) > 0){
    #     int[[paste0("att_rw_",n)]][i] <- rw
    #   }
    # }
  }
  # remove basic info from att to avoid redundancy
  att <- att %>% select(-ROUTE, -BEG_MP, -END_MP)
  # join to segments
  int <- left_join(int, att, by = c("att_rw_1"="att_rw"))
  int <- left_join(int, att, by = c("att_rw_2"="att_rw"), suffix = c(".a", ".b"))
  int <- left_join(int, att, by = c("att_rw_3"="att_rw"))
  int <- left_join(int, att, by = c("att_rw_4"="att_rw"), suffix = c(".c", ".d"))
  # determine max, min, etc...
  for(i in colnames(int)){
    if(substrRight(i,2) == ".d"){
      var <- substrMinusRight(i,2)
      int <- int %>%
        rowwise() %>%
        mutate(
          !!sym(paste0("MAX_",var)) := max(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d"))), 
                                           na.rm = TRUE),
          !!sym(paste0("MIN_",var)) := min(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d"))), 
                                           na.rm = TRUE),
          !!sym(paste0("AVG_",var)) := mean(c(!!sym(paste0(var,".a")),
                                              !!sym(paste0(var,".b")),
                                              !!sym(paste0(var,".c")),
                                              !!sym(paste0(var,".d"))), 
                                            na.rm = TRUE),
          !!sym(paste0(var,"_1")) := !!sym(paste0(var,".a")),
          !!sym(paste0(var,"_2")) := !!sym(paste0(var,".b")),
          !!sym(paste0(var,"_3")) := !!sym(paste0(var,".c")),
          !!sym(paste0(var,"_4")) := !!sym(paste0(var,".d"))
        ) %>%
        # convert infinities to NA
        mutate(
          !!sym(paste0("MAX_",var)) := ifelse(is.infinite(!!sym(paste0("MAX_",var))), NA, !!sym(paste0("MAX_",var))),
          !!sym(paste0("MIN_",var)) := ifelse(is.infinite(!!sym(paste0("MIN_",var))), NA, !!sym(paste0("MIN_",var))),
          !!sym(paste0("AVG_",var)) := ifelse(is.nan(!!sym(paste0("AVG_",var))), NA, !!sym(paste0("AVG_",var)))
        )
    }
  }
  # aadt specific
  if(is_aadt == TRUE){
    for(i in colnames(int)){
      if(substrRight(i,2) == ".d"){
        var <- substrMinusRight(i,2)
        if(substr(var,1,4) == "AADT"){
          int <- int %>%
            rowwise() %>%
            mutate(
              KNOWN_LEGS := 4L - sum(is.na(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d"))))),
              !!sym(var) := case_when(
                as.integer(NUM_LEGS) == as.integer(KNOWN_LEGS) ~ 
                  sum(c(!!sym(paste0(var,".a"))/2, 
                        !!sym(paste0(var,".b"))/2, 
                        !!sym(paste0(var,".c"))/2, 
                        !!sym(paste0(var,".d"))/2), na.rm = TRUE),
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 1L ~ 
                  sum(c(!!sym(paste0(var,".a"))/2, 
                        !!sym(paste0(var,".b"))/2, 
                        !!sym(paste0(var,".c"))/2, 
                        !!sym(paste0(var,".d"))/2,
                      mean(c(!!sym(paste0(var,".a"))/2, 
                             !!sym(paste0(var,".b"))/2,
                             !!sym(paste0(var,".c"))/2, 
                             !!sym(paste0(var,".d"))/2), na.rm = TRUE) * 1), na.rm = TRUE),
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 2L ~ 
                  sum(c(!!sym(paste0(var,".a"))/2, 
                        !!sym(paste0(var,".b"))/2, 
                        !!sym(paste0(var,".c"))/2, 
                        !!sym(paste0(var,".d"))/2,
                      mean(c(!!sym(paste0(var,".a"))/2, 
                             !!sym(paste0(var,".b"))/2,
                             !!sym(paste0(var,".c"))/2, 
                             !!sym(paste0(var,".d"))/2), na.rm = TRUE) * 2), na.rm = TRUE),
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 3L ~ 
                  sum(c(!!sym(paste0(var,".a"))/2, 
                        !!sym(paste0(var,".b"))/2, 
                        !!sym(paste0(var,".c"))/2, 
                        !!sym(paste0(var,".d"))/2,
                      mean(c(!!sym(paste0(var,".a"))/2, 
                             !!sym(paste0(var,".b"))/2,
                             !!sym(paste0(var,".c"))/2, 
                             !!sym(paste0(var,".d"))/2), na.rm = TRUE) * 3), na.rm = TRUE),
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 4L ~ 
                  sum(c(!!sym(paste0(var,".a"))/2, 
                        !!sym(paste0(var,".b"))/2, 
                        !!sym(paste0(var,".c"))/2, 
                        !!sym(paste0(var,".d"))/2,
                      mean(c(!!sym(paste0(var,".a"))/2, 
                             !!sym(paste0(var,".b"))/2,
                             !!sym(paste0(var,".c"))/2, 
                             !!sym(paste0(var,".d"))/2), na.rm = TRUE) * 4), na.rm = TRUE)
              )
            )
        } else{
          int <- int %>%
            rowwise() %>%
            mutate(
              KNOWN_LEGS := 4L - sum(is.na(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d"))))),
              !!sym(var) := case_when(
                as.integer(NUM_LEGS) == as.integer(KNOWN_LEGS) ~ 
                  mean(c(!!sym(paste0(var,".a")),
                          !!sym(paste0(var,".b")),
                          !!sym(paste0(var,".c")),
                          !!sym(paste0(var,".d"))), na.rm = TRUE),
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 1L ~ 
                  sum(c(!!sym(paste0(var,".a")), 
                        !!sym(paste0(var,".b")),
                        !!sym(paste0(var,".c")), 
                        !!sym(paste0(var,".d")),
                      mean(c(!!sym(paste0(var,".a")),
                             !!sym(paste0(var,".b")),
                             !!sym(paste0(var,".c")),
                             !!sym(paste0(var,".d"))), 
                           na.rm = TRUE) * 1), na.rm = TRUE) / NUM_LEGS,
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 2L ~ 
                  sum(c(!!sym(paste0(var,".a")), 
                        !!sym(paste0(var,".b")),
                        !!sym(paste0(var,".c")), 
                        !!sym(paste0(var,".d")),
                      mean(c(!!sym(paste0(var,".a")),
                             !!sym(paste0(var,".b")),
                             !!sym(paste0(var,".c")),
                             !!sym(paste0(var,".d"))), 
                           na.rm = TRUE) * 2), na.rm = TRUE) / NUM_LEGS,
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 3L ~ 
                  sum(c(!!sym(paste0(var,".a")), 
                        !!sym(paste0(var,".b")),
                        !!sym(paste0(var,".c")), 
                        !!sym(paste0(var,".d")),
                      mean(c(!!sym(paste0(var,".a")),
                             !!sym(paste0(var,".b")),
                             !!sym(paste0(var,".c")),
                             !!sym(paste0(var,".d"))), 
                           na.rm = TRUE) * 3), na.rm = TRUE) / NUM_LEGS,
                
                as.integer(NUM_LEGS)-as.integer(KNOWN_LEGS) == 4L ~ 
                  sum(c(!!sym(paste0(var,".a")), 
                        !!sym(paste0(var,".b")),
                        !!sym(paste0(var,".c")), 
                        !!sym(paste0(var,".d")),
                      mean(c(!!sym(paste0(var,".a")),
                             !!sym(paste0(var,".b")),
                             !!sym(paste0(var,".c")),
                             !!sym(paste0(var,".d"))), 
                           na.rm = TRUE) * 4), na.rm = TRUE) / NUM_LEGS
              )
            )
        }
      }
    }
  }
  # remove att_rw info, and return segments
  int <- int %>% ungroup %>% select(-contains("att_rw"), -contains("."))
  return(int)
}

# Simple computational functions
substrRight <- function(x, n){
  substr(x, nchar(x)-n+1, nchar(x))
}
substrMinusRight <- function(x, n){
  substr(x, 1, nchar(x)-n)
}

left_join_fill <- function(x, y, by, fill = 0L, ...){
  z <- left_join(x, y, by, ...)
  new_cols <- setdiff(names(z), names(x))
  z <- replace_na(z, setNames(as.list(rep(fill, length(new_cols))), new_cols))
  return(z)
}

# Modify intersections to include bus stops and schools near intersections
mod_intersections <- function(intersections,UTA_Stops,schools){
  # intersections <- st_join(intersections, schools, join = st_within) %>%
  #   group_by(Int_ID) %>%
  #   mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
  #   select(-SchoolID)
  # 
  # intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
  #   group_by(Int_ID) %>%
  #   mutate(NUM_UTA = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
  #   select(-UTA_StopID)
  # 
  # intersections <- unique(intersections)
  # st_drop_geometry(intersections)
  
  # Alternative Method (more variables)

  # Separate schools and UTA by type
  bus_stops <- UTA_stops %>%
    filter(Mode == "Bus") %>%
    select(-Mode)

  rail_stops <- UTA_stops %>%
    filter(Mode == "Rail") %>%
    select(-Mode)

  high_schools <- schools %>%
    filter(SchoolLeve == "HIGH")

  mid_schools <- schools %>%
    filter(SchoolLeve == "MID")

  elem_schools <- schools %>%
    filter(SchoolLeve == "ELEM")

  prek_schools <- schools %>%
    filter(SchoolLeve == "PREK")

  ktwelve_schools <- schools %>%
    filter(SchoolLeve == "K12")

  # determine which intersections are 1000 ft from a school
  ints_near_schools <- st_join(intersections, schools, join = st_within) %>%
    filter(!is.na(SchoolID))

  # determine which intersections are 1000 ft from a UTA stop
  ints_near_UTA <- st_join(intersections, UTA_stops, join = st_within) %>%
    filter(!is.na(UTA_StopID))

  # modify intersections object
  intersections <- st_join(intersections, schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, prek_schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_PREK_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, elem_schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_ELEM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, mid_schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_MID_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, high_schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_HIGH_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, ktwelve_schools, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_K12_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
    select(-SchoolID, -SchoolLeve) %>%
    unique()

  intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_UTA_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
    select(-UTA_StopID, -Mode) %>%
    unique()

  intersections <- st_join(intersections, bus_stops, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_BUS_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
    select(-UTA_StopID) %>%
    unique()

  intersections <- st_join(intersections, rail_stops, join = st_within) %>%
    group_by(Int_ID) %>%
    mutate(NUM_RAIL_STOPS = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
    select(-UTA_StopID) %>%
    unique()

  intersections <- unique(intersections)
  return(intersections)
}





###
## Unused Code
###

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

# ###
# ## Unused Intersection Code
# ###
#
# # Modify intersections to include bus stops and schools near intersections
# mod_intersections <- function(intersections,UTA_Stops,schools){
#   intersections <- st_join(intersections, schools, join = st_within) %>%
#     group_by(Int_ID) %>%
#     mutate(NUM_SCHOOLS = length(SchoolID[!is.na(SchoolID)])) %>%
#     select(-SchoolID)
#   
#   intersections <- st_join(intersections, UTA_stops, join = st_within) %>%
#     group_by(Int_ID) %>%
#     mutate(NUM_UTA = length(UTA_StopID[!is.na(UTA_StopID)])) %>%
#     select(-UTA_StopID)
#   
#   intersections <- unique(intersections)
#   st_drop_geometry(intersections)
# }