library(targets)
tar_option_set(packages = c("tidyverse", "dpylr"))
source("R/Targets_CrashPrep.R")
source("R/Targets_RoadwayPrep.R")
list(
  # Read in Crash Files
  tar_target(location, read_csv_file(location_fp, location_col)),
  tar_target(rollups, read_csv_file(rollups, rollups_fp, rollups_col)),
  tar_target(vehicle, read_csv_file(vehicle_fp, vehicle_col)),
  # Combine Crash Files
  tar_target(crash, left_join(location,rollups,by='crash_id')),
  tar_target(fullcrash,left_join(crash,vehicle,by='crash_id')),
  #
  tar_target(crashdata, crashfilter(crash))
)
