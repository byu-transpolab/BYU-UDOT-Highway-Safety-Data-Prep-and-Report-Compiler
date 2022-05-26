library(targets)
tar_option_set(packages = c("tidyverse", "dpylr"))
source("R/Targets_CrashPrep.R")
source("R/Targets_RoadwayPrep.R")
list(
  # Read in Crash Files
  tar_target(location, read_csv_file(location_fp, location_col)),
  tar_target(rollups, read_csv_file(rollups_fp, rollups_col)),
  tar_target(vehicle, read_csv_file(vehicle_fp, vehicle_col)),
  # Combine Crash Files
  tar_target(crash, left_join(location, rollups, by='crash_id')),
  tar_target(fullcrash,left_join(crash, vehicle, by='crash_id')),
  # Filter Crashes
  tar_target(crashdata, crashfilter(crash)),
  # Read in Roadway Files
  tar_target(routes, read_csv_file(routes_fp, routes_col)),
  tar_target(fc, read_csv_file(fc_fp, fc_col)),
  tar_target(aadt, read_csv_file(aadt_fp, aadt_col)),
  tar_target(speed, read_csv_file(speed_fp, speed_col)),
  tar_target(lane, read_csv_file(lane_fp, lane_col)),
  tar_target(urban, read_csv_file(urban_fp, urban_col)),
  tar_target(driveway, read_csv_file(driveway_fp, location_col)),
  tar_target(median, read_csv_file(median_fp, median_col)),
  tar_target(shoulder, read_csv_file(shoulder_fp, shoulder_col))
)

