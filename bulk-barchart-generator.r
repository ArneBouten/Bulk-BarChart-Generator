# ===================
# PACKAGE INSTALLATION
# ===================

# Ensure all required packages are installed and loaded
required_packages <- c("haven", "ggplot2", "dplyr", "officer", "stringr")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)
lapply(required_packages, library, character.only = TRUE)

# =====================
# BAR CHART CREATION
# =====================

# Function to create bar charts
create_bar_chart <- function(data, question, group_var, group_value, current_group_label) {
 # Convert the question variable to a factor using SPSS value labels
 data[[question]] <- haven::as_factor(data[[question]])

 # Get the label for the question variable
 # If no label is found, use the question name as a fallback
 question_label <- attr(data[[question]], "label")
 if (is.null(question_label)) question_label <- question
 
 # Wrap the title if it's too long to fit on the slide
 # Adjust the width as needed
 wrapped_title <- str_wrap(question_label, width = 60)
 # Add a newline character to create space below the title
 wrapped_title <- paste(wrapped_title, "\n ")

 # Filter out NA values for the question and group variable
 data_filtered <- data %>% filter(!is.na(.data[[question]]), !is.na(.data[[group_var]]))

 # Check if all values are missing for the current question in the current group
 if (sum(!is.na(data_filtered[[question]][data_filtered[[group_var]] == group_value])) == 0) {
   cat("Skipping question", question, "for group", current_group_label, "due to only missing values.\n")
   return(NULL)
 }

 # Create the plot using ggplot2
 ggplot(data_filtered, aes(x = .data[[question]], fill = factor(.data[[group_var]] == group_value, labels = c("Alle andere groepen", current_group_label)))) +
   geom_bar(aes(y = after_stat(prop), group = factor(.data[[group_var]] == group_value)), position = position_dodge(preserve = "single"), alpha = 0.7) +
   labs(title = wrapped_title, x = NULL, y = NULL, fill = "Group") +
   scale_y_continuous(labels = scales::percent) +
   scale_fill_manual(values = c("#1172e0e1", "#39c964e5")) +
   scale_x_discrete(drop = FALSE) +
   theme_minimal() +
   theme(plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
         axis.text.x = element_text(size = 9.5),
         axis.text.y = element_text(size = 9.5),
         legend.text = element_text(size = 9.5))
}

# ========================
# DATA LOADING AND PREPARATION
# ========================

# ***** ADAPT THIS VALUE *****
# Replace "your_data_file.sav" with the actual file name of your SPSS data file
your_data <- read_sav("your_data_file.sav")

# ***** ADAPT THIS VALUE *****
# Replace -999 with the actual missing value placeholder used in your data (if different)
your_data[your_data == -999] <- NA

# ***** ADAPT THIS VALUE *****
# Replace "your_group_variable" with the actual variable name representing the groups in your data
your_data$your_group_variable <- as_factor(your_data$your_group_variable)

# ***** ADAPT THESE VALUES *****
# Replace the question variable names with your actual question variable names
questions <- c("question1", "question2", "question3", ...)

# ========================
# POWERPOINT GENERATION
# ========================

# Loop through each group in the data
for (group_value in levels(your_data$your_group_variable)) {
 
# Use the group_value directly as the group label
group_label <- group_value

# Check if the group has at least one person in it
 if (sum(!is.na(your_data$your_group_variable[your_data$your_group_variable == group_value])) == 0) {
   cat("Skipping group", group_label, "due to zero people in this group.\n")
   next  # Skip to the next group if there are no people in the current group
 }

# Print the group label before processing each group
cat("Processing group:", group_label, "\n")

# Initialize a new PowerPoint presentation
ppt <- read_pptx()

# Add a title slide to the presentation
 ppt <- add_slide(ppt, layout = "Title Slide", master = "Office Theme")
 
 # ***** ADAPT THIS VALUE *****
 # Replace "YOUR TITLE" with the desired title for your PowerPoint presentation
 group_label_upper <- toupper(group_value)  # Convert the group label to uppercase
 my_text <- ftext(paste0("YOUR TITLE ", group_label_upper), 
                  prop = fp_text(font.size = 54, bold = TRUE, font.family = "Calibri (Headings)"))
 
 # Create the paragraph formatting object for center alignment
 my_format <- fp_par(text.align = "center")
 
 # Add the formatted text to the slide, using the center alignment for the paragraph
 ppt <- ph_with(ppt, value = fpar(my_text, fp_p = my_format), location = ph_location_type(type = "ctrTitle"))

# ***** ADAPT THIS VALUE *****
# Replace "SECTION 1" with the desired name for the first section
ppt <- add_slide(ppt, layout = "Title Slide", master = "Office Theme")
ppt <- ph_with(ppt, value = "SECTION 1", location = ph_location_type(type = "ctrTitle"))

# Initialize a counter to keep track of the current question number
question_counter <- 1
 
# Loop through each question to add slides and charts to the PowerPoint
for (q in questions) {
 # Create a bar chart for the current question
 chart <- create_bar_chart(your_data, q, "your_group_variable", group_value, group_label)

  # Check if the chart is NULL (i.e., when there are only missing values for the current question in the current group)
  if (is.null(chart)) {
    question_counter <- question_counter + 1
    next  # Skip to the next question if the chart is NULL
  }
 # Add a new slide with the "Title and Content" layout
 ppt <- add_slide(ppt, layout = "Title and Content", master = "Office Theme")

 # ***** ADAPT THESE VALUES *****
 # Creating a custom title for each slide in PowerPoint
 # Update the switch statement with the desired short titles for each question
 # If no match is found, the original question label will be used as a fallback
 short_title <- switch(q,
                       "question1" = "Question 1 Title",
                       "question2" = "Question 2 Title",
                       "question3" = "Question 3 Title",
                       ...,
                       attr(your_data[[q]], "label"))
 
 # Add the short title to the slide
 ppt <- ph_with(ppt, value = short_title, location = ph_location_type(type = "title"))
 
 # Add the chart to the slide
 ppt <- ph_with(ppt, value = chart, location = ph_location_type(type = "body"))
 
 # Increment the question counter
 question_counter <- question_counter + 1

 # ***** ADAPT THESE VALUES *****
 # Adding section separator slides at specific question intervals
 # Modify the section names and their corresponding question counter values as needed
 if (question_counter == 5 || question_counter == 11 || question_counter == 24 || question_counter == 27) {
   section_name <- ifelse(question_counter == 5, "SECTION 2",
                          ifelse(question_counter == 11, "SECTION 3",
                                 ifelse(question_counter == 24, "SECTION 4", "SECTION 5")))
   
   # Add a section separator slide with the corresponding section name
   ppt <- add_slide(ppt, layout = "Title Slide", master = "Office Theme")
   ppt <- ph_with(ppt, value = section_name, location = ph_location_type(type = "ctrTitle"))
 }
}
 
# ***** ADAPT THIS VALUE *****
# Replace "your_output_prefix" with the desired prefix for the output PowerPoint file names
ppt_file_name <- paste0("your_output_prefix_", group_label, ".pptx")
print(ppt, target = ppt_file_name)
}
