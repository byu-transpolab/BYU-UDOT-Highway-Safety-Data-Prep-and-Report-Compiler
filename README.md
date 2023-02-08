# BYU/UDOT Highway Safety Data Prep and Report Compiler
This repository contains files related to data preparation and report compiling for UDOT highway safety network screening. These include VBA files which were used in the old Excel-based model. Instructions for how to use that model are contained within the accompanying excel file kept in a Box drive folder. The main folder contains a file "RGUI.Rmd" which is used as the graphical user interface (GUI) for the current project. The "R" folder contains all the R code referenced by the RGUI file. The "Python" folder contains some Python code which were used in tandem with ArcGIS Pro for visualization purposes but were not used within the model. The "data" folder contains all input and output data for the R model. The data folder is included in the ".gitignore" file so it needs to be populated with input data kept in the Box drive folder for this project.

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
Fixes the last ending milepoints (This is also a backchecker for the input files). Since our input datasets are apparently on different LRS's, the milepoints on one file don't match up with another file for the same route. This function does not convert the LRS, because that is not feasible to do without ArcGIS, but it does make sure at least the last milepoint of the route matches across all datasets to prevent the creation of nonexistent segments. Unfortunately, THIS DOES NOT FIX THE UNDERLYING PROBLEM, but it will still be useful when all input datasets are on the same LRS, because it will correct any slight differences that exist, perhaps due to rounding. In the meantime, it is also useful for identifying routes where the LRS is significantly different so we can pay special attention to those routes. (update: the input files are on the same LRS now, but this is still a very useful function.)
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
- **return_max_min:** (boolean) Does the crash attribute need to be added as a max/min variable or a single value?



