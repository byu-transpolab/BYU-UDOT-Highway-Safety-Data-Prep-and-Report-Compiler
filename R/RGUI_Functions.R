###############################################################################

###
## Load in Libraries
###

library(tidyverse)
library(dplyr)
library(sf)

###############################################################################

###
## Initial Data Prep Functions
###

# Read in Function
## Refers to the filepath and columns list for each and every data set that we 
## have and ensures that only the selected columns are read in
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

###############################################################################

###
## Segment Cleaning Functions
###

# Compress Segments Function
## If adjacent segments in a data set have the same attributes from selected 
## columns they are combined into one segment.
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
## This does the exact same thing as the previous function, but uses a more
## streamlined process. However, this function is generally more computationally
## expensive so we typically use the other one.
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

# Compress Overlapping Segments Function
## Ensures no milepoints of unique segments overlap. This needs to be done in 
## order for the compress segments by length function and other functions to
## work.
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

# Segment Length Compress Function
## Ensuring that segment lengths are not less than 0.1 of a mile, (or other user
## specified length). If segment is less than 0.1 then combine to the adjacent 
## segment that is most similar.
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

###############################################################################

###
## AADT Functions
###

# Convert AADT to Numeric Function
## Ensuring AADT data type is a number not a string.
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
## Create a different row for every year the segment exists. This function 
## makes good use of the tidyverse pivot functions to simplify the process.
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
## The major change from pivot_aadt is that we use "contains" instead of
## "starts_with" This allows us to include the multitude of new aadt related 
## columns created by the intersection data combining process. This function 
## could easily replace the previous function if we want to just use one.
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

# Missing Negative AADT
## Fixes missing negative direction AADT values. For some reason, the AADT data
## does not contain negative route (divided route) data. It seems to assume
## negative and positive routes have the same AADT. However, we need to add in
## entries for negative routes so we can create segments along negative routes.
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

# Fill Missing AADT
## Function to fill in missing aadt and truck data. This function uses a 
## comprehensive method to estimate missing data. First, it determines whether
## the missing data should be filled or not by looking for data in previous 
## years. If there is no data for all previous years, we assume the segment did 
## not exist at that time and do not fill in the missing data. Otherwise, we 
## estimate the missing data by creating a grid from data in adjacent years 
## and adjacent segments. The following is the grid with years on the x axis 
## and segments on the y axis....
##
## [[a1 a2 a3]
##  [b1 b2 b3]
##  [c1 c2 c3]]
##
## We solve for b2 with the following equation....
##
## b2 <- mean(a2*mean(b1/a1,b3/a3),c2*mean(b1/c1,b3/c3))
##
## If this does not work we use one of the following....
##
## b2 <- mean(a2, c2)
## b2 <- mean(b1, b3)
##
## This function is repeated several times to ensure all missing data is filled
## in.
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
              if(df[["ROUTE"]][i] == df[["ROUTE"]][i-1] & i-1>0){
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
              if(df[["ROUTE"]][i] == df[["ROUTE"]][i+1] & i+1<nrow(df)){
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

###############################################################################

###
## LRS Correction Functions
###

# Negative to Positive
## Merge divided negative routes into positive routes for the full data set only.
## This is still a work in progress. Right now it's just a copy of the aadt_neg
## function because it will need to do the same thing in reverse essentially.
## This may need to be done because it seems like we are only given positive 
## routes in the intersection file. We are still waiting for more clarification
## from UDOT.
neg_to_pos_routes <- function(df, rtes, divd){
  # isolate negative routes
  rtes_n <- rtes %>% filter(grepl("N",ROUTE)) %>% select(-BEG_MP, -END_MP)
  # isolate positive att entries
  att_p <- att %>% filter(grepl("P",ROUTE)) %>%
    mutate(
      label = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE))
    ) %>%
    select(-ROUTE)
  # isolate negative att entries
  att_n <- att %>% filter(grepl("N",ROUTE))
  # join negative routes with negative att (there may be a more efficient way to do this)
  df <- full_join(att_n, rtes_n, by = "ROUTE") %>%
    mutate(
      label = as.integer(gsub(".*?([0-9]+).*", "\\1", ROUTE))
    ) %>%
    select(ROUTE, BEG_MP, END_MP, everything()) %>%
    filter(is.na(BEG_MP))
  # join missing negative routes with positive segments on the same route
  df <- left_join(df, att_p, by = "label") %>%
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
    beg_att <- df[["BEG_MP"]][i]
    end_att <- df[["END_MP"]][i]
    rt_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt & beg_att < divd_p$END_MP & end_att > divd_p$BEG_MP)
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
    if(beg_att > end_rt[1]){        # flag segments that fall outside route
      df[["flag"]][i] <- TRUE
    }
    if(end_att > end_rt[1]){        # correct end mp if it's greater than route
      df[["END_MP"]][i] <- end_rt[1]
    }
    if(end_att < beg_rt[1]){        # flag segments that fall outside route
      df[["flag"]][i] <- TRUE
    }
    if(beg_att < beg_rt[1]){        # correct beg mp if it's less than route
      df[["BEG_MP"]][i] <- beg_rt[1]
    }
  }
  # filter out flagged rows
  df <- df %>% filter(flag == FALSE) %>% select(-flag)
  # convert positive milepoints to negative milepoints
  for(i in 1:nrow(df)){
    rt <- df[["ROUTE"]][i]
    rt <- substr(rt, start = 1, stop = 4)
    beg_att <- df[["BEG_MP"]][i]
    end_att <- df[["END_MP"]][i]
    # find segments on divd tables
    prt_row <- which(substr(divd_p$ROUTE, start = 1, stop = 4) == rt & beg_att < divd_p$END_MP & end_att > divd_p$BEG_MP)
    nrt_row <- which(substr(divd_n$ROUTE, start = 1, stop = 4) == rt & beg_att < divd_n$END_MP & end_att > divd_n$BEG_MP)
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
    df[["BEG_MP"]][i] <- (beg_att - prt_start) * c 
    df[["END_MP"]][i] <- (end_att - prt_start) * c 
    # fix segment breaks (treat gaps as zero)
    # I had to specify route 89 because there are routes with non-zero gaps which
    # need to be maintained. (mainly route 84 jumps from 42 to 81) If we can come 
    # up with a more robust solution that would be great.
    if(i-1>0){
      if(df[["BEG_MP"]][i] != df[["END_MP"]][i-1] & df[["ROUTE"]][i] == df[["ROUTE"]][i-1] & rt == "0089"){
        offset <- df[["BEG_MP"]][i] - df[["END_MP"]][i-1]
        end_att <- df[["END_MP"]][i]
        df[["END_MP"]][i] <- end_att - offset
        df[["BEG_MP"]][i] <- df[["END_MP"]][i-1]
      }
    }
  }
  # rbind df to att
  df <- rbind(att, df) %>%
    arrange(ROUTE, BEG_MP)
  # return df
  return(df)
}

# Fix Endpoints
## Fix last ending milepoints (This is also a backchecker for the input files).
## Since our input datasets are apparently on different LRS's, the milepoints on
## one file don't match up with another file for the same route. This function
## does not convert the LRS, because that is not feasible to do without ArcGIS, 
## but it does make sure at least the last milepoint of the route matches 
## across all datasets to prevent the creation of nonexistent segments. 
## Unfortunately, THIS DOES NOT FIX THE UNDERLYING PROBLEM, but it will still be 
## useful when all input datasets are on the same LRS, because it will correct 
## any slight differences that exist, perhaps due to rounding. In the meantime,
## it is also useful for identifying routes where the LRS is significantly 
## different so we can pay special attention to those routes.
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

###############################################################################

###
## Segment Creation Functions (created by stats team)
###

# Segment Breaks
## Get all measurements where a road segment stops or starts 
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

###############################################################################

###
## Combine Functions
###

# Add Crash Attributes Function
## This function combines crash data to roadway segment data. There needs to be
## a column for segment id in the crash file which we take care of in the 
## "RGUI_Compile.R" script. Then, we pivot wider the point crash data so that it 
## makes more sense in the context of segments.
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
## This is essentially the same as "add_crash_attribute" but uses intersection
## variables instead. We could easily consolidate these into one function by 
## adding additional parameters or conditions if we wanted to.
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

# Add Medians Data to Segments Function
## Since medians are not a segmenting variable, but still need to be added to 
## segment data, we needed a different process to add them in. This process 
## could be adapted to work for shoulders and other non segmenting segment data
## if we want to, but we already coded most of that into the 
## "RGUI_RoadwayPrep.R" script. This function works by determining the most 
## prevalent median type along the segment. It also creates a new column for 
## each median type on the segments dataset and marks whether that median is 
## present on the segment with a 1 or a 0.
add_medians <- function(RC, median, min_portion){
  # optional arguments
  if(missing(min_portion)){min_portion <- 0}
  # set up dataframe for joining
  med <- median %>% select(ROUTE, BEG_MP, END_MP, MEDIAN_TYP)
  # identify segments for each median
  med$SEG_ID <- NA
  med$LENGTH <- NA
  med$flagged <- NA
  delete <- 0
  add <- 0
  for (i in 1:nrow(med)){
    rt <- med$ROUTE[i]
    med_beg <- med$BEG_MP[i]
    med_end <- med$END_MP[i]
    seg_row <- which(RC$ROUTE == rt & 
                       RC$BEG_MP < med_end & 
                       RC$END_MP > med_beg)
    seg_freq <- length(seg_row)
    seg_beg <- RC$BEG_MP[seg_row]
    seg_end <- RC$END_MP[seg_row]
    seg_id <- RC$SEG_ID[seg_row]
    # attach segment ids and interior lengths to medians
    if(seg_freq > 0){
      # create new rows for split medians
      if(seg_freq > 1){
        for(j in 2:seg_freq){
          next_row <- nrow(med) + 1
          med[next_row, 4] <- med$MEDIAN_TYP[i]
          med[next_row, 5] <- seg_id[j]
          if(med_beg < seg_beg[j]){beg <- seg_beg[j]}else{beg <- med_beg}
          if(med_end > seg_end[j]){end <- seg_end[j]}else{end <- med_end}
          med[next_row, 6] <- end - beg
          seg_len <- seg_end[j] - seg_beg[j]
          # delete row if the median a small portion of segment
          if(med[next_row, 6]/seg_len < min_portion){
            med$flagged[next_row] <- 1
            delete <- delete + 1
          } else{add <- add + 1}
        }
      }
      med$SEG_ID[i] <- seg_id[1]
      if(med_beg < seg_beg[1]){beg <- seg_beg[1]}else{beg <- med_beg}
      if(med_end > seg_end[1]){end <- seg_end[1]}else{end <- med_end}
      med$LENGTH[i] <- end - beg
      seg_len <- seg_end[1] - seg_beg[1]
      # delete row if the median a small portion of segment
      if(med$LENGTH[i]/seg_len < min_portion & seg_freq > 1){
        med$flagged[i] <- 1
        delete <- delete + 1
      }
    }
  }
  # organize median data
  med <- med %>% 
    filter(is.na(flagged), !is.na(SEG_ID)) %>%
    select(SEG_ID,MEDIAN_TYP) %>%
    arrange(SEG_ID)
  # check for deleted segments
  # print(paste("deleted",delete,"rows"))
  # print(paste("added",add,"new rows"))
  missing <- 0
  for(i in 2:nrow(med)){
    id <- med$SEG_ID[i]
    prev_id <- med$SEG_ID[i-1]
    if((id-prev_id) > 1){
      missing <- missing + 1
    }
  }
  missing <- missing + (RC$SEG_ID[nrow(RC)] - med$SEG_ID[nrow(med)])
  print(paste("missing",missing,"segments"))
  # pivot wider
  med <- med %>%
    group_by(SEG_ID, MEDIAN_TYP) %>% 
    mutate(n = n(),
           n = ifelse(n>1,1,n)) %>% 
    ungroup() %>%
    pivot_wider(id_cols = c(SEG_ID),
                names_from = MEDIAN_TYP,
                names_prefix = "MEDIAN_TYP_",
                values_from = n,
                values_fill = 0,
                values_fn = first)
  # join to segments
  df <- left_join_fill(RC, med, by = c("SEG_ID"="SEG_ID"), fill = 0L)
  return(df)
}

# Add roadway attributes to intersections function
## Since there is no segmenting process for intersections, all roadway variables
## need to be added in afterwards from different datasets. We created different
## datasets with the suffix "_full" to indicate that these contain all available 
## data from various routes, including non-state routes. We need this because 
## our intersections dataset includes state routes that intersect with non-state
## routes. This is probably our longest function because it needs to account 
## for a wide variety of different conditions and parameters. However, it uses
## a general process that works well for most input datasets. It creates columns
## for route 1, 2, 3, 4, and 5 as well as columns for min, max, and average. 
## it is up to the function user to decide which of these columns they will keep 
## for the specific application, or if they will use the resulting columns for 
## additional analysis.
##
## However, this process falls short with the functional class and AADT inputs. 
## For functional class, there are some routes listed on the intersection as 
## "local". Although we don't know what route these are, we do know the 
## functional class is "local", so there is a condition that checks for 
## functional class and makes sure "local" is included. For AADT, we use an 
## entirely different process because we need to estimate missing AADT where 
## there is no route number given, and we need to combine the AADT and truck
## in such a way that we get total entering vehicles instead. There is a
## condition in this function that checks for AADT and runs additional analysis
## if the input dataset is AADT.
add_int_att <- function(int, att, is_aadt, is_fc){
  # optional arguments
  if(missing(is_aadt)){is_aadt <- FALSE} else if(is_aadt == TRUE){
    int <- int %>% select(-contains("AADT"), -contains("SUTRK"), -contains("CUTRK"))
  }
  if(missing(is_fc)){is_fc <- FALSE}
  # create row id column for attribute data
  att <- att %>% rownames_to_column("att_rw")
  att$att_rw <- as.integer(att$att_rw)
  # identify attribute segments on intersection file
  for(k in 0:4){               # since we can expect intersections to have overlapping attributes, we need to identify all of these
    int[[paste0("att_rw_",k)]] <- NA
  }
  for (i in 1:nrow(int)){
    rt0 <- int$INT_RT_0[i]
    mp0 <- int$INT_RT_0_M[i]
    rt1 <- int$INT_RT_1[i]
    mp1 <- int$INT_RT_1_M[i]
    rt2 <- int$INT_RT_2[i]
    mp2 <- int$INT_RT_2_M[i]
    rt3 <- int$INT_RT_3[i]
    mp3 <- int$INT_RT_3_M[i]
    rt4 <- int$INT_RT_4[i]
    mp4 <- int$INT_RT_4_M[i]
    rw0 <- which(att$ROUTE == rt0 & 
                   att$BEG_MP < mp0 & 
                   att$END_MP > mp0)
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
    if(length(rw0) > 0){
      int$att_rw_0[i] <- att$att_rw[rw0]
    }
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
  int <- left_join(int, att, by = c("att_rw_0"="att_rw"))
  int <- left_join(int, att, by = c("att_rw_1"="att_rw"), suffix = c(".a", ".b"))
  int <- left_join(int, att, by = c("att_rw_2"="att_rw"))
  int <- left_join(int, att, by = c("att_rw_3"="att_rw"), suffix = c(".c", ".d"))
  int <- left_join(int, att, by = c("att_rw_4"="att_rw"))
  int <- left_join(int, att, by = c("att_rw_4"="att_rw"), suffix = c(".e", ".remove"))
  int <- int %>% select(-contains(".remove"))
  # determine max, min, etc...
  for(i in colnames(int)){
    if(substrRight(i,2) == ".a"){
      var <- substrMinusRight(i,2)
      int <- int %>%
        rowwise() %>%
        mutate(
          !!sym(paste0("MAX_",var)) := max(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d")),
                                             !!sym(paste0(var,".e"))), 
                                           na.rm = TRUE),
          !!sym(paste0("MIN_",var)) := min(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d")),
                                             !!sym(paste0(var,".e"))), 
                                           na.rm = TRUE),
          !!sym(paste0("AVG_",var)) := mean(c(!!sym(paste0(var,".a")),
                                              !!sym(paste0(var,".b")),
                                              !!sym(paste0(var,".c")),
                                              !!sym(paste0(var,".d")),
                                              !!sym(paste0(var,".e"))), 
                                            na.rm = TRUE),
          !!sym(paste0(var,"_0")) := !!sym(paste0(var,".a")),
          !!sym(paste0(var,"_1")) := !!sym(paste0(var,".b")),
          !!sym(paste0(var,"_2")) := !!sym(paste0(var,".c")),
          !!sym(paste0(var,"_3")) := !!sym(paste0(var,".d")),
          !!sym(paste0(var,"_4")) := !!sym(paste0(var,".e"))
        ) %>%
        # convert infinities to NA
        mutate(
          !!sym(paste0("MAX_",var)) := ifelse(is.infinite(!!sym(paste0("MAX_",var))), NA, !!sym(paste0("MAX_",var))),
          !!sym(paste0("MIN_",var)) := ifelse(is.infinite(!!sym(paste0("MIN_",var))), NA, !!sym(paste0("MIN_",var))),
          !!sym(paste0("AVG_",var)) := ifelse(is.nan(!!sym(paste0("AVG_",var))), NA, !!sym(paste0("AVG_",var)))
        )
      # add local functional classes if functional class
      if(is_fc == TRUE){
        int <- int %>%
          rowwise() %>%
          mutate(
            !!sym(paste0(var,"_0")) := case_when(tolower(INT_RT_0) == "local" ~ "Local",
                                                 TRUE ~ !!sym(paste0(var,".a"))),
            !!sym(paste0(var,"_1")) := case_when(tolower(INT_RT_1) == "local" ~ "Local",
                                                 TRUE ~ !!sym(paste0(var,".b"))),
            !!sym(paste0(var,"_2")) := case_when(tolower(INT_RT_2) == "local" ~ "Local",
                                                 TRUE ~ !!sym(paste0(var,".c"))),
            !!sym(paste0(var,"_3")) := case_when(tolower(INT_RT_3) == "local" ~ "Local",
                                                 TRUE ~ !!sym(paste0(var,".d"))),
            !!sym(paste0(var,"_4")) := case_when(tolower(INT_RT_4) == "local" ~ "Local",
                                                 TRUE ~ !!sym(paste0(var,".e")))
          )
      }
    }
  }
  # aadt specific
  if(is_aadt == TRUE){
    for(i in colnames(int)){
      if(substrRight(i,2) == ".a"){
        var <- substrMinusRight(i,2)
        if(substr(var,1,4) == "AADT"){
          int <- int %>%
            rowwise() %>%
            mutate(
              NUM_LEGS := 5L - sum(is.na(c(INT_RT_0,
                                           INT_RT_1,
                                           INT_RT_2,
                                           INT_RT_3,
                                           INT_RT_4))),
              KNOWN_LEGS := 5L - sum(is.na(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d")),
                                             !!sym(paste0(var,".e"))))),
              num_local = sum(tolower(c(INT_RT_0,
                                        INT_RT_1,
                                        INT_RT_2,
                                        INT_RT_3,
                                        INT_RT_4)) == "local", na.rm = TRUE),
              local_add = case_when(TRAFFIC_CO == "SIGNAL" ~ num_local * 1000,
                                    TRUE ~ num_local * 500),
              num_missing = (NUM_LEGS - KNOWN_LEGS) - num_local,
              # We could take average for routes still missing, but just adding 1000 or 500 is probably better.
              # missing_add = mean(c(!!sym(paste0(var,".a"))/2, 
              #                      !!sym(paste0(var,".b"))/2,
              #                      !!sym(paste0(var,".c"))/2, 
              #                      !!sym(paste0(var,".d"))/2,
              #                      !!sym(paste0(var,".e"))/2), 
              #                    na.rm = TRUE) * num_missing,
              missing_add = case_when(TRAFFIC_CO == "SIGNAL" ~ num_missing * 1000,
                                      TRUE ~ num_missing * 500),
              !!sym(var) := as.integer(sum(c(!!sym(paste0(var,".a"))/2, 
                                            !!sym(paste0(var,".b"))/2, 
                                            !!sym(paste0(var,".c"))/2, 
                                            !!sym(paste0(var,".d"))/2,
                                            !!sym(paste0(var,".e"))/2,
                                            local_add,
                                            missing_add), na.rm = TRUE))
            )
        } else{
          aadt_var <- paste0("AADT",substrRight(var,4))
          int <- int %>%
            rowwise() %>%
            mutate(
              NUM_LEGS := 5L - sum(is.na(c(INT_RT_0,
                                           INT_RT_1,
                                           INT_RT_2,
                                           INT_RT_3,
                                           INT_RT_4))),
              KNOWN_LEGS := 5L - sum(is.na(c(!!sym(paste0(var,".a")),
                                             !!sym(paste0(var,".b")),
                                             !!sym(paste0(var,".c")),
                                             !!sym(paste0(var,".d")),
                                             !!sym(paste0(var,".e"))))),
              num_local = sum(tolower(c(INT_RT_0,
                                        INT_RT_1,
                                        INT_RT_2,
                                        INT_RT_3,
                                        INT_RT_4)) == "local", na.rm = TRUE),
              # local_add = case_when(TRAFFIC_CO == "SIGNAL" ~ (num_local*0)/num_local,   # will cause infinities. fix this
              #                       TRUE ~ (num_local*0)/num_local),
              local_add = 0 * num_local,
              num_missing = (NUM_LEGS - KNOWN_LEGS) - num_local,
              # missing_add = mean(c(!!sym(paste0(var,".a"))*!!sym(paste0(aadt_var,".a"))/2, 
              #                      !!sym(paste0(var,".b"))*!!sym(paste0(aadt_var,".b"))/2,
              #                      !!sym(paste0(var,".c"))*!!sym(paste0(aadt_var,".c"))/2, 
              #                      !!sym(paste0(var,".d"))*!!sym(paste0(aadt_var,".d"))/2,
              #                      !!sym(paste0(var,".e"))*!!sym(paste0(aadt_var,".e"))/2), 
              #                    na.rm = TRUE) * num_missing,
              missing_add = 0 * num_missing,
              !!sym(var) := sum(c(!!sym(paste0(var,".a"))*!!sym(paste0(aadt_var,".a"))/2, 
                                  !!sym(paste0(var,".b"))*!!sym(paste0(aadt_var,".b"))/2,
                                  !!sym(paste0(var,".c"))*!!sym(paste0(aadt_var,".c"))/2, 
                                  !!sym(paste0(var,".d"))*!!sym(paste0(aadt_var,".d"))/2,
                                  !!sym(paste0(var,".e"))*!!sym(paste0(aadt_var,".e"))/2,
                                  local_add,
                                  missing_add), na.rm = TRUE)
            )
        }
      }
    }
    int <- int %>% select(-num_local,-local_add,-num_missing,-missing_add)
  }
  # remove att_rw info, and return segments
  int <- int %>% ungroup %>% select(-contains("att_rw"), -contains("."))
  return(int)
}

# Expand Intersection Attribute Function
## This function may be used in conjunction with add_int_att to perform
## additional analysis. It uses the pivot functions to create new columns for 
## each roadway attribute type (e.g. the types on functional class are "local",
## "interstate", "arterial", etc...). This is often a preferable format for 
## statistical analysis. It also removes the need to have different columns for
## each approach at the intersection.
expand_int_att <- function(IC, att){
  # extract attribute data
  df <- IC %>% 
    select(Int_ID, contains(att))
  # pivot longer
  df <- df %>%
    pivot_longer(cols = c(!!sym(paste0(att,"_0")),
                          !!sym(paste0(att,"_1")),
                          !!sym(paste0(att,"_2")),
                          !!sym(paste0(att,"_3")),
                          !!sym(paste0(att,"_4"))),
                 names_to = "delete",
                 values_to = att) %>%
    select(-delete) %>%
    filter(!is.na(!!sym(att)))
  # pivot wider
  df <- df %>%
    group_by(Int_ID, !!sym(att)) %>% 
    mutate(n = n(),
           n = ifelse(n>1,1,n)) %>% 
    ungroup() %>%
    pivot_wider(id_cols = c(Int_ID),
                names_from = !!sym(att),
                names_prefix = paste0(att,"_"),
                values_from = n,
                values_fill = 0,
                values_fn = max)
  # join to segments
  df <- left_join_fill(IC, df, by = c("Int_ID"="Int_ID"), fill = 0L) %>%
    select(-(!!sym(paste0(att,"_0")):!!sym(paste0(att,"_4"))))
  return(df)
}

# Modify intersections to include bus stops and schools near intersections
## This function uses R's geospatial capabilities within the "sf" package to 
## combine transit and school data to intersections. It uses a spatial buffer 
## and spatial intersect to determine how many and which type of transit stops 
## and schools are within 1000 feet from the intersection.
mod_intersections <- function(intersections,UTA_Stops,schools){
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

###############################################################################

###
## Additional Interpolation/Data Cleaning Functions
###

# Fill in Missing Roadway Data
## This function simply looks through a list of columns and fills in missing 
## data with data from adjacent segments
fill_all_missing <- function(df, missing, rt_col, subsetting_var, subset_vars){
  n <- 1   # set the subsetting_var index
  subset_chk <- 0   # set the variable to check for subsetting
  for(var in missing){
    filled <- 0     # reset filled count
    not_filled <- 0     # reset not filled count
    # check if the previous variable was subsetting then increment subsetting index
    if(subset_chk == 1 & n < length(subsetting_var)){
      n <- n + 1
      subset_chk <- 0
    }
    # display progress
    print(paste("Filling Missing Values for",var))
    # loop forwards then backwards in case data is missing at beginning of route
    for(iter in 1:2){
      if(iter == 1){b <- 0}else{b <- nrow(df)+1}
      # loop through rows
      for(i in 1:nrow(df)){
        a <- abs(b - i)
        # check if the value is NA
        value <- df[[var]][a]
        if(is.na(value)){
          # record previous and next segments 
          prev <- NA
          nxt <- NA
          prev_match <- 0
          nxt_match <- 0
          # fill previous segment
          if(df[[rt_col]][a] == df[[rt_col]][a-1] & a-1>0){
            prev <- df[[var]][a-1]
            # count variables in common with previous segment
            for(j in 1:ncol(df)){
              if(df[a,j] == df[a-1,j] & !is.na(df[a,j]) & !is.na(df[a-1,j])){
                prev_match <- prev_match + 1
              }
            }
          }
          # fill next segment
          if(df[[rt_col]][a] == df[[rt_col]][a+1] & a+1<nrow(df)){
            nxt <- df[[var]][a+1]
            # count variables in common with next segment
            for(j in 1:ncol(df)){
              if(df[a,j] == df[a+1,j] & !is.na(df[a,j]) & !is.na(df[a+1,j])){
                nxt_match <- nxt_match + 1
              }
            }
          }
          # check whether previous or next segment has more in common and fill
          if(nxt_match > prev_match & !is.na(nxt) | is.na(prev)){
            df[[var]][a] <- nxt
            # fill subsetted variables if applicable
            if(var == subsetting_var[n]){
              subset_chk <- 1
              for(k in subset_vars[[n]]){
                df[[k]][a] <- df[[k]][a+1]
              }
            }
          } else {
            df[[var]][a] <- prev
            # fill subsetted variables if applicable
            if(var == subsetting_var[n]){
              subset_chk <- 1
              for(k in subset_vars[[n]]){
                df[[k]][a] <- df[[k]][a-1]
              }
            }
          }
          # count the variables filled and not filled
          if(!is.na(df[[var]][a])){
            filled <- filled + 1
          } else if(is.na(df[[var]][a]) & iter == 2){
            not_filled <- not_filled + 1
          }
        }
      }
    }
    # announce variables filled and not filled
    print(paste("Filled",filled,"missing entries. However",(not_filled),"entries could not be filled."))
    print("-------------------------------------")
  }
  # return df
  return(df)
}


###############################################################################

###
## Simple Computational Functions
###

## Modifies the substr function to take n characters from the right
substrRight <- function(x, n){
  substr(x, nchar(x)-n+1, nchar(x))
}


## Modifies the substr function to remove n characters from the right
substrMinusRight <- function(x, n){
  substr(x, 1, nchar(x)-n)
}


## Adds an additional parameter to the left_join function which specifies how to
## fill in missing data rather than defaulting to "NA"
left_join_fill <- function(x, y, by, fill = 0L, ...){
  z <- left_join(x, y, by, ...)
  new_cols <- setdiff(names(z), names(x))
  z <- replace_na(z, setNames(as.list(rep(fill, length(new_cols))), new_cols))
  return(z)
}

