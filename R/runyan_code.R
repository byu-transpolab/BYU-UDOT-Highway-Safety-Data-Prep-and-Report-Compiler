# I'm writing my code here instead of directly on the .Rmd because I don't want
# to conflict with Tommy's code.

# INSTALL AND LOAD PACKAGES ################################

# Install pacman ("package manager") if needed
if (!require("pacman")) install.packages("pacman")

# pacman must already be installed; then load contributed
# packages (including pacman) with pacman
pacman::p_load(magrittr, pacman, tidyverse)


# FILTER CAMS CRASH DATA

# change route 089A to 0011... There seems to be more to this though
crash$ROUTE_ID[crash$ROUTE_ID == '089A'] <- 0011

# change route ID to 194 from 0085 for milepoint less than 2.977 
crash$ROUTE_ID[crash$MILEPOINT < 2.977 & crash$ROUTE_ID == 0085] <- 0194

# eliminate rows with ramps
crash %>%
  as_tibble() %>%
  filter(RAMP_ID == 0 & is.na(RAMP_ID))
  