# BYU-UDOT-Highway-Safety-Data-Prep-and-Report-Compiler
This repository contains files related to data preparation and report compiling for UDOT highway safety network screening

## Functions
### `read_csv_file(filepath, columns)`
This uses the dplyr function `read_csv()` to read .csv files into tibble dataframes. It takes the filepath and desired columns to keep from the .csv file as inputs. The filepath is given as a string and the list of columns is given as a list or vector. This function forces the user to make sure input dataframes contain the correct columns, else an error message will occur. For many of the dataframes in the model, it is important to list route, beginning milepoint, and ending milepoint as the first three columns in the desired columns input.
- **filepath:** (string) filepath of the .csv file to input, including name and .csv extension
- **columns:** (list) desired columns to keep from the .csv file. The order of this list does not need to match the order of the columns in the input file.

### `read_filez_shp(filepath, columns)`
This uses the sf function `read_sf()` to read .shp files into tibble dataframes. It works much the same way as `read_csv_file()`. 
- **filepath:** (string) filepath of the .shp file to input, including name and .shp extension
- **columns:** (list) desired columns to keep from the .shp file. The order of this list does not need to match the order of the columns in the input file.

### `compress_seg(df, col = colnames(df), variables = colnames(df))`
This function has a similar purpose as the `unique()` function except that it accounts for spatial nature of milepoints. If adjacent segments in a data set have the same attributes from selected columns they are combined into one segment. It takes the lower beginning milepoint of the two segments and the larger ending milepoint to combine into one longer segment.
- **df:** (tibble) dataframe to be compressed. must include the columns ROUTE, BEG_MP, and END_MP.
- **col:** (list) list of column names in the dataframe to keep. leave blank to keep all columns.
- **variables:** (list) list of column names to consider for joining segments. by default, this is every column included in col except the first three. (this assumes the first three columns are ROUTE, BEG_MP, and END_MP.)

### `compress_seg_alt(df)`
