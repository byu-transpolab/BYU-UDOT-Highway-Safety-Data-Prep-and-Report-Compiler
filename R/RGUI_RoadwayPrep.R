# The following is the functions that prepare the roadway data to be joined

# Keeps the following columns from the aadt file
check_aadt <- function(table){
  table %>%
    select(rt_num, start_accu,end_accum, station, aadrt2020 ,aadt2019, aadt2018, aadt2017, aadt2016,sutrk2020, cutrk2020)
}

# Keeps the following columns from the functional class file
check_functionalclass <- function(table){
  table %>%
    select(from_measure, to_measure, functional_class)
}

# Keeps the following columns from the intersections file
check_intersections <- function(table){
  table %>%
    select()
}

# Keeps the following columns from the lanes file
check_lanes <- function(table){
  table %>%
    select()
}

# Keeps the following columns from the pavement messages file
check_pavementmessages <- function(table){
  table %>%
    select()
}

# Keeps the following columns from the speed limits file
check_speedlimits <- function(table){
  table %>%
    select()
}