###
## Load in Libraries
###

library(tidyverse)
library(dplyr)

###
## Set Filepath and Column Names for each Dataset
###

routes_fp <- "data/csv/UDOT_Routes_ALRS.csv"
routes_col <- c("ROUTE_ID",                          
                    "BEG_MILEAGE",                             
                    "END_MILEAGE")

aadt_fp <- "data/csv/AADT_Unrounded.csv"
aadt_col <- c("ROUTE_NAME",
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

fc_fp <- "data/csv/Functional_Class_ALRS.csv"
fc_col <- c("ROUTE_ID",                          
                "FROM_MEASURE",                             
                "TO_MEASURE",
                "FUNCTIONAL_CLASS",                             
                "RouteDir",                             
                "RouteType")

speed_fp <- "data/csv/UDOT_Speed_Limits_2019.csv"
speed_col <- c("ROUTE_ID",
                   "FROM_MEASURE",
                   "TO_MEASURE",
                   "SPEED_LIMIT")

lane_fp <- "data/csv/Lanes.csv"
lane_col <- c("ROUTE",
                  "START_ACCUM",
                  "END_ACCUM",
                  "THRU_CNT",
                  "THRU_WDTH")

urban_fp <- "data/csv/Urban_Code.csv"
urban_col <- c("ROUTE_ID",
                   "FROM_MEASURE",
                   "TO_MEASURE",
                   "URBAN_CODE")

intersection_fp <- "data/csv/Intersections.csv"
intersection_col <- c("ROUTE",
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


driveway_fp<- "data/csv/Driveway.csv"
driveway_col<- c("ROUTE",
                      "START_ACCUM",
                      "END_ACCUM",
                      "DIRECTION",
                      "TYPE")

median_fp <- "data/csv/Medians.csv"
median_col <- c("ROUTE_NAME",
                    "START_ACCUM",
                    "END_ACCUM",
                    "MEDIAN_TYP",
                    "TRFISL_TYP",
                    "MDN_PRTCTN")

shoulder_fp <-"data/csv/Shoulders.csv"
shoulder_col <- c("ROUTE",
                      "START_ACCUM", # these columns are rounded quite a bit. Should we use Mandli_BMP and Mandli_EMP?
                      "END_ACCUM",
                      "UTPOSITION",
                      "SHLDR_WDTH")

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