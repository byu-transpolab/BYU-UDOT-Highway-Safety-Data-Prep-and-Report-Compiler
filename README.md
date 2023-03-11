# BYU/UDOT Highway Safety Data Prep and Report Compiler
This repository contains files related to data preparation and report compiling for UDOT highway safety network screening. These include VBA files which were used in the old Excel-based model. Instructions for how to use that model are contained within the accompanying excel file kept in a Box drive folder. The main folder contains a file "RGUI.Rmd" which is used as the graphical user interface (GUI) for the current project. The "R" folder contains all the R code referenced by the RGUI file. The "Python" folder contains some Python code which were used in tandem with ArcGIS Pro for visualization purposes but were not used within the model. The "data" folder contains all input and output data for the R model. The data folder is included in the ".gitignore" file so it needs to be populated with input data kept in the Box drive folder for this project.

## ReadIn
The code to read in and clean the raw data files are in the file called "RGUI_ReadIn.R". Currently, the user needs to specify which columns of data and the filenames to read in from this R script. In the future, it may be beneficial to include user inputs for this. New data cleaning code will need to be written if additional files are included.

## CrashPrep
The code to prepare the crash data are in the file called "RGUI_CrashPrep.R". This includes combining crash files into one, setting the formating for crash files, and separating crashes by intersection and segment crashes. Originally, milepoints were used to determine the functional area of the intersections for separating crashes. However, due to LRS inconsistencies in the intersection file, it was more reasonable to use spatial buffers to determine the functional areas. This required using a spatial join between the crashes and intersections, followed by code to determine the closest intersection using the euclidian distance.

## RoadwayPrep
The code to prepare the roadway data are in the file called "RGUI_RoadwayPrep.R". This includes creating a segments file using segmenting variables, adding roadway attributes to the segments and intersections files, and pivoting the data by year of available AADT data. This last step means the dataframes will contain multiple rows of data for each segment or intersection, so this must be accounted for in future scripts. This is also the step where missing data are filled in and segment boundaries are cleaned. It is important to clean or "compress" segments so there are not a large number of very small segments.

## Compile
The code to compile crash and roadway information into a single intersection and segments files are in the file called "RGUI_Compile.R". This is done with a modified version of the `left_join()` function in R. It is important that intersection and segment ID's are assigned to crashes prior to joining. This is also the step where "parameters" files are created to prepare for report compiling. These are essentially just a modified version of the crash dataframes which are designed specifically to work as inputs for the old excel report compiler.

## Report Compiler
The code to compile reports is contained in the file called "RGUI_Report_Compiler.R". This may eventually all be done in R, but for now, this just takes the statistical output files from the statistics team and reformats it to work as an input for the old excel report compiler.

## Functions
The functions used within the R code are contained in the file called "RGUI_Functions.R". Below is a summary what each of these functions do.

### Data Reading Functions

### `read_csv_file(filepath, columns)`
This uses the dplyr function `read_csv()` to read .csv files into tibble dataframes. It takes the filepath and desired columns to keep from the .csv file as inputs. The filepath is given as a string and the list of columns is given as a list or vector. This function forces the user to make sure input dataframes contain the correct columns, else an error message will occur. For many of the dataframes in the model, it is important to list route, beginning milepoint, and ending milepoint as the first three columns in the desired columns input.
- **filepath:** (string) filepath of the .csv file to input, including name and .csv extension
- **columns:** (list) desired columns to keep from the .csv file. The order of this list does not need to match the order of the columns in the input file.

### `read_filez_shp(filepath, columns)`
This uses the sf function `read_sf()` to read .shp files into tibble dataframes. It works much the same way as `read_csv_file()`. 
- **filepath:** (string) filepath of the .shp file to input, including name and .shp extension
- **columns:** (list) desired columns to keep from the .shp file. The order of this list does not need to match the order of the columns in the input file.


### Segment Cleaning Functions

### `compress_seg(df, col = colnames(df), variables = tail(col, -3))`
This function has a similar purpose as the `unique()` function except that it accounts for the spatial nature of milepoints. If adjacent segments in a data set have the same attributes from selected columns they are combined into one segment. It takes the lower beginning milepoint of the two segments and the larger ending milepoint to combine into one longer segment.
- **df:** (tibble) dataframe to be compressed. must include the columns ROUTE, BEG_MP, and END_MP.
- **col:** (list) list of column names in the dataframe to keep. leave blank to keep all columns.
- **variables:** (list) list of column names to consider for joining segments. by default, this is every column included in col except the first three. (this assumes the first three columns are ROUTE, BEG_MP, and END_MP.)

### `compress_seg_alt(df)`
This function works in the same exact way as the `compress_seg()` function, except without the optional arguments. This code was developed separately by the Statistics students who developed the segmentation functions. It has been generalized to work with any input data because it sometimes runs faster than the original `compress_seg()` function.
- **df:** (tibble) dataframe to be compressed. must include the columns ROUTE, BEG_MP, and END_MP.

### `compress_seg_ovr(df)`
This function ensures no milepoints of unique segments overlap. This needs to be done in order for the compress segments by length function and other functions to work.
- **df:** (tibble) dataframe to be compressed. must include the columns ROUTE, BEG_MP, and END_MP.

### `compress_seg_len(df, len)`
This function ensures that segment lengths are not less than 0.1 of a mile, (or other user specified length). If a segment is less than the specified length then it combines with the adjacent segment that is most similar.
- **df:** (tibble) dataframe to be compressed. must include the columns ROUTE, BEG_MP, and END_MP.
- **len:** (double) minimum segment length to be specified.


### AADT Functions

### `aadt_numeric(aadt, aadt.columns)`
This function ensures AADT data type is a number not a string.
- **aadt:** (tibble) dataframe containing AADT data.
- **aadt.columns:** (list) list of AADT columns to consider. (This is admittedly pointless since the code still works if all columns are included, but it is still a required input.)

### `pivot_aadt(aadt)`
Create a different row for every year the segment exists. This function makes good use of the tidyverse pivot functions to simplify the process.
- **aadt:** (tibble) dataframe containing unpivoted AADT data. This will most likely be the RC dataframe which contains compiled segment data.

### `pivot_aadt_int(aadt)`
A more robust version of `pivot_aadt()` which works with intersection data instead. With minor adjustments it could be generalized to work for both. (The major change from pivot_aadt is that we use "contains" instead of "starts_with" This allows us to include the multitude of new aadt related columns created by the intersection data combining process.)
- **aadt:** (tibble) dataframe containing unpivoted AADT data. This will most likely be the IC dataframe which contains compiled intersection data.

### `aadt_neg(aadt, rtes, divd)`
Fixes missing negative direction AADT values. For some reason, the raw AADT data does not contain negative route (divided route) data. It seems to assume negative and positive routes have the same AADT. However, we need to add in entries for negative routes so we can create segments along negative routes.
- **aadt:** (tibble) dataframe containing AADT data.
- **rtes:** (tibble) dataframe containing routes data.
- **divd:** (tibble) dataframe containing divided routes data.

### `fill_missing_aadt(df, colns)`
Function to fill in missing aadt and truck data. This function uses a comprehensive method to estimate missing data. First, it determines whether the missing data should be filled or not by looking for data in previous years. If there is no data for all previous years, we assume the segment did not exist at that time and do not fill in the missing data. Otherwise, we estimate the missing data by creating a grid from data in adjacent years and adjacent segments. The following is the grid with years on the x axis and segments on the y axis....

|    |    |    |
|----|----|----|
| a1 | a2 | a3 |
| b1 | b2 | b3 |
| c1 | c2 | c3 |

We solve for b2 with the following equation....

`b2 <- mean(a2*mean(b1/a1,b3/a3),c2*mean(b1/c1,b3/c3))`

If this does not work we use one of the following....

`b2 <- mean(a2, c2)`
`b2 <- mean(b1, b3)`

- **df:** (tibble) dataframe containing AADT data to be filled in.
- **colns:** (list) list of columns to consider.


### LRS Correction Functions

### `neg_to_pos_routes(df, rtes, divd)`
Merge divided negative routes into positive routes for the full data set only. This is still a work in progress. Right now it's just a copy of the aadt_neg function because it will need to do the same thing in reverse essentially. This may need to be done because it seems like we are only given positive routes in the intersection file. We are still waiting for more clarification from UDOT. (update: There isn't a significant need for this. This was going to be used for the intersection dataframe, IC, but our current methodology is more streamlined and doesn't require as much data.)

### `fix_endpoints(df, routes)`
Fixes the last ending milepoints (This is also a backchecker for the input files). Since our input datasets are apparently on different LRS's, the milepoints on one file don't match up with another file for the same route. This function does not convert the LRS, because that is not feasible to do without ArcGIS, but it does make sure at least the last milepoint of the route matches across all datasets to prevent the creation of nonexistent segments. Unfortunately, THIS DOES NOT FIX THE UNDERLYING PROBLEM, but it will still be useful when all input datasets are on the same LRS, because it will correct any slight differences that exist, perhaps due to rounding. In the meantime, it is also useful for identifying routes where the LRS is significantly different so we can pay special attention to those routes. (**Update:** the input files are on the same LRS now, but this is still a very useful function.)
- **df:** (tibble) dataframe to be fixed. must include the columns ROUTE, BEG_MP, and END_MP.
- **routes:** (tibble) dataframe containing routes data.


### Segmentation Functions

### `segment_breaks(sdtm)`
Created by statistics students for a separate project. Get all measurements where a road segment stops or starts

### `shell_creator(sdtms)`
Created by statistics students for a separate project. This function is robust to data missing for segments of road presumed to exist (such as within the bounds of other declared road segments) where there is no data. This mostly seems to happen when a road changes hands (state road to interstate, or state to city road, or something like that).

### `shell_join(sdtm)`
Created by statistics students for a separate project. This segment of the program is essentially reconstructing a SAS DATA step with a RETAIN statement to populate through rows. We start by left joining the shell to each sdtm dataset and populating through all the records whose union is the space interval we have data on. We then get rid of the original START_ACCUM, END_ACCUM variables, since startpoints, endpoints have all the information we need.


### Join Functions

### `add_crash_attribute(att_name, segs, csh, return_max_min = FALSE)`
This function combines crash data to roadway segment data. There needs to be a column for segment id in the crash file which we take care of in the  "RGUI_Compile.R" script. Then, we pivot wider the point crash data so that it makes more sense in the context of segments.
- **att_name:** (string) name of crash attribute to join.
- **segs:** (tibble) segment dataframe to join to.
- **csh:** (tibble) crash dataframe to pull crash attribute from.
- **return_max_min:** (boolean) Does the crash attribute need to be added as max/min/avg variables? The default is to create a column for each possible value and tally the number of crashes where that outcome occured at each segment. However, sometimes max/min/avg is better for more quantitative variables such as number of vehicles involved in a crash.

### `add_crash_attribute_int(att_name, ints, csh, return_max_min = FALSE)`
This is essentially the same as "add_crash_attribute" but uses intersection variables instead. We could easily consolidate these into one function by adding additional parameters or conditions if we wanted to.
- **att_name:** (string) name of crash attribute to join.
- **ints:** (tibble) intersection dataframe to join to.
- **csh:** (tibble) crash dataframe to pull crash attribute from.
- **return_max_min:** (boolean) Does the crash attribute need to be added as max/min/avg variables? The default is to create a column for each possible value and tally the number of crashes where that outcome occured at each intersection. However, sometimes max/min/avg is better for more quantitative variables such as number of vehicles involved in a crash.

### `add_medians(RC, median, min_portion = 0)`
Since medians are not a segmenting variable, but still need to be added to segment data, we needed a different process to add them in. This process could be adapted to work for shoulders and other non segmenting segment data if we want to, but we already coded most of that into the "RGUI_RoadwayPrep.R" script. This function works by determining the most prevalent median type along the segment. It also creates a new column for each median type on the segments dataset and marks whether that median is present on the segment with a 1 or a 0. (*Note:* This function is not very robust to changes in the input data. If a new medians file is used, this function may need to be improved.)
- **RC:** (tibble) segments dataframe to join to.
- **median:** (tibble) medians dataframe to join from.
- **min_portion:** (double) minimum allowable ratio of median length to segment length. This way, medians which take only a small portion of the segment aren't used to describe the segment.

### `add_int_att(int, att, is_aadt = FALSE, is_fc = FALSE)`
Since there is no segmenting process for intersections, all roadway variables need to be added in afterwards from different datasets. We created different datasets with the suffix "_full" to indicate that these contain all available data from various routes, including non-state routes. We need this because our intersections dataset includes state routes that intersect with non-state routes. This is probably our longest function because it needs to account for a wide variety of different conditions and parameters. However, it uses a general process that works well for most input datasets. It creates columns for route 1, 2, 3, 4, and 5 as well as columns for min, max, and average. it is up to the function user to decide which of these columns they will keep for the specific application, or if they will use the resulting columns for additional analysis. However, this process falls short with the functional class and AADT inputs. For functional class, there are some routes listed on the intersection as "local". Although we don't know what route these are, we do know the functional class is "local", so there is a condition that checks for functional class and makes sure "local" is included. For AADT, we use an entirely different process because we need to estimate missing AADT where there is no route number given, and we need to combine the AADT and truck data in such a way that we get total entering vehicles instead. There is a condition in this function that checks for AADT and runs additional analysis if the input dataset is AADT.
(**Update:** The latest intersection file has entering vehicles so we don't use the AADT data anymore. We still use trucks from the AADT data though.)
- **int:** (tibble) intersection dataframe to join to.
- **att:** (tibble) roadway attribute dataframe to join from.
- **is_aadt:** (boolean) is att the AADT dataframe?
- **is_fc:** (boolean) is att the functional class dataframe?

### `expand_int_att <- function(IC, att, all_legs = TRUE)`
This function may be used in conjunction with add_int_att to perform additional analysis. It uses the pivot functions to create new columns for each roadway attribute type (e.g. the types on functional class are "local", "interstate", "arterial", etc...). This is often a preferable format for statistical analysis. It also removes the need to have different columns for each approach at the intersection.
- **IC:** (tibble) intersection dataframe to expand.
- **att:** (string) name of the roadway attribute within the dataframe to expand.
- **all_legs:** (boolean) is there only one column or a column for each intersection leg?

### `mod_intersections <- function(intersections, UTA_Stops, schools)`
Modify intersections to include bus stops and schools near intersections This function uses R's geospatial capabilities within the "sf" package to combine transit and school data to intersections. It uses a spatial buffer and spatial intersect to determine how many and which type of transit stops and schools are within 1000 feet from the intersection.
(**Note:** The UTA stops and schools files need to be buffered before using this function.)
- **intersections:** (tibble) intersection dataframe to join to.
- **UTA_Stops:** (tibble) buffered UTA stops dataframe to join.
- **schools:** (tibble) buffered schools dataframe to join.


### Additional Interpolation / Data Cleaning Functions

### `fill_all_missing <- function(df, missing, rt_col, subsetting_var, subset_vars)`
This function simply looks through a list of columns and fills in missing data with data from adjacent segments or intersections
- **df:** (tibble) dataframe to fill missing values for.
- **missing:** (string list) list of columns which contain missing data.
- **rt_col:** (string) name of the column which contains the route name to compare adjecent segments or intersections on.
- **subsetting_var:** (string list) variables in missing which where expanded using the `expand_int_att()` function.
- **subset_vars:** (list of string lists) the variables which were expanded. the order of this list must correlate to the order of subsetting_var.


### Simple Comutational Functions

### `substrRight <- function(x, n)`
Modifies the substr function to take n characters from the right

### `substrMinusRight <- function(x, n)`
Modifies the substr function to remove n characters from the right

### `left_join_fill <- function(x, y, by, fill = 0L, ...)`
Adds an additional parameter to the left_join function which specifies how to fill in missing data rather than defaulting to "NA"

### `user.input <- function(prompt)`
Dynamic user input function pulled from https://stackoverflow.com/questions/27112370/make-readline-wait-for-input-in-r
