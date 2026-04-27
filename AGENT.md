# AGENT Instructions

## Overview

You are software engineering developing a light-weight excel workbook generator based on configuration files. The entire method should take two inputs: (1) a roster of names and date of births, (2) a configuration file that specifies evolutions and metrics collected for each evolution (we often use evo for shorthand).

After completing the workbook generator, we are adding a simple GUI application for the end user to use the application without needing to run any command line scripts. 
We are use a simple MVC architecture with PyQt. 

The program outputs a data collection ready excel notebook pre-populated with a one to one mapping for evolutions and sheets. Each evo (sheet) displays candidate names locked on the left side (y-axis) and metrics on the x-axis with predefined data types (and dropdowns is specified in configs).

The program should also contain an "inverse" operation that takes an excel workbook and converts to a master excel file based on the data contract for the master excel file.

We want to prioritize flexibility and correctness. To be more specific, the users should be able to change the configuration values at will and run the program to re-generate a new excel sheet. Even better, the user can have different config files for different sets on evos. It should be correct in that no data should be lost when aggregating the scores from the data collection workbook to the master excel file.

### Structure

- `config` contains all the configuration files for creating workbooks including the roster for input.
- `src` contains the source code for main and excel operations
- `workbooks` is the drop folder for generating workbooks. The program should drop newly generated excels here.

