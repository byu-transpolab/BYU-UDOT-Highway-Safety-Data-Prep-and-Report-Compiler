# BYU-UDOT-Highway-Safety-Data-Prep-and-Report-Compiler
This repository contains files related to data preparation and report compiling for UDOT highway safety network screening
1) Most of the code for copying datasets is in the CopyDataSets Macro on the "Home" module. This is conveniently used by all models.
2) If a line of code looks suspicious I put a comment next to it saying "FLAGGED". This way you can use Ctrl-F to search for any "FLAGGED" code. 
This is usually code which seems too rigid, and might cause issues if the data changes.
3) If there is a button which starts a long string of code, a lot of that code is likely stored in the button's code. Look through the forms in VBA to find it. 
This code is usually a hub which references several macros within various modules. Look here to get the big picture of what's happening.
