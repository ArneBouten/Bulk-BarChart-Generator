# README for bulk-barchart-generator

## Overview
This R script is designed to analyze survey (Likert-scale) data from an SPSS file and generate a PowerPoint presentation with a bar chart on each slide. It creates a presentation for each group in the data (e.g., children, adolescents, adults, etc.) For each question, it creates a double bar chart comparing the responses of one group to all the other groups combined.

## Prerequisites
- R environment set up on your system.

## Setup Instructions
1. Prepare the SPSS data file:
     - Ensure that your SPSS data file has a structure similar to "example_data_file.sav".
     - The graphs will be automatically created based on the labels, values, etc. defined in SPSS, so properly setting up the SPSS file is crucial.
     - Create new titles for each question, add questions as labels, define values (e.g., for "Group": 1 = children, etc.), set missing values to -999.
2. Place your SPSS data file in the same directory as the R script.
3. Open the R script using your preferred R IDE, such as R-studio.
4. Customize the script to suit your specific data:
    - Values that need to be adapted are always shown in the code with # ***** ADAPT THIS VALUE *****
    - Replace `"your_data_file.sav"` with your SPSS data file's name.
    - Substitute `-999` with your data's actual placeholder for missing values, if it differs.
    - Alter `"your_group_variable"` to reflect the variable name representing groups within your data.
    - Update the `questions` vector with your survey's actual question variable names.
    - In the `switch` statement, adjust the `short_title` mapping to set short titles for each PowerPoint slide.
    - Modify section names and their corresponding question counts as necessary.
    - Set the `ppt_file_name` prefix to your desired output file name.

## Usage Guidelines
1. Executing the script will process each group within the data, generating a  PowerPoint presentation for each.
2. Generated PowerPoint files will be saved in the same directory as the R script, following the naming convention `"your_output_prefix_[group_label].pptx"`.
3. Monitor the console for any messages or warnings that might arise during the script's execution.

## Customization Options
- The appearance of the bar charts can be tailored by modifying the `create_bar_chart` function. This includes adjustments to colors, fonts, sizes, etc.
- For changes to the PowerPoint slides' layout or design, refer to the `officer` package documentation and modify the relevant sections of the script accordingly.

## For further help or questions, please don't hesitate to reach out.
