# Add spatial data to CAMS. Temporary solution to Excel GUI problem.
add_spatial <- function(df, sf){
  sf <- sf %>%
    select(ROUTE_NAME, geometry)
  df <- left_join(df, sf, by = c("Route_Name" = "ROUTE_NAME")) %>%
    st_as_sf(wkt=geometry)
  df
}