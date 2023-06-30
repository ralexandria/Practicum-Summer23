# Import the necessary libraries
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)

# Read the Excel file
file_path <- "/Users/ralexandria/Downloads/survey.xlsx"
sheet1 <- read_excel(file_path, sheet = 1)
sheet2 <- read_excel(file_path, sheet = 2)

# Convert 'Timestamp' to datetime type in both sheets
sheet1$Timestamp <- as.POSIXct(sheet1$Timestamp, origin = "1970-01-01")
sheet2$Timestamp <- as.POSIXct(sheet2$Timestamp, origin = "1970-01-01")

# Combine the sheets into a single data frame
combined_data <- bind_rows(sheet1, sheet2)

# Clean the 'Age' variable
combined_data$Age <- as.character(combined_data$Age)
combined_data$Age <- case_when(
  combined_data$Age %in% as.character(1:130) ~ combined_data$Age,
  TRUE ~ NA_character_
)

# Create a new data frame 'age_cleaned_data' with cleaned variable
age_cleaned_data <- data.frame(Age = combined_data$Age)

# Assign the cleaned 'age' variable to 'combined_data'
combined_data$Age <- age_cleaned_data$Age

# Create a barchart of the 'Age' variable
ggplot(combined_data, aes(x = Age)) +
  geom_bar(stat = "count") +
  labs(title = "Distribution of Age", x = "Age", y = "Frequency")

# Clean the 'Country' variable
combined_data$Country <- tolower(combined_data$Country)
combined_data$Country <- case_when(
  combined_data$Country == "" ~ NA_character_,
  combined_data$Country == "us" ~ "United States",
  combined_data$Country == "uk" ~ "United Kingdom",
  TRUE ~ combined_data$Country
)

# Capitalize the first letter of each country name
combined_data$Country <- tools::toTitleCase(combined_data$Country)

# Create a new data frame 'country_cleaned_data' with cleaned variable
country_cleaned_data <- data.frame(Country = combined_data$Country)

# Assign the cleaned 'country' variable to 'combined_data'
combined_data$Country <- country_cleaned_data$Country

# Clean the 'state' variable
combined_data$state <- case_when(
  combined_data$state == "" ~ NA_character_,
  combined_data$state == "Alabama" ~ "AL",
  combined_data$state == "Alaska" ~ "AK",
  combined_data$state == "Arizona" ~ "AZ",
  combined_data$state == "Arkansas" ~ "AR",
  combined_data$state == "California" ~ "CA",
  combined_data$state == "Colorado" ~ "CO",
  combined_data$state == "Connecticut" ~ "CT",
  combined_data$state == "Delaware" ~ "DE",
  combined_data$state == "Florida" ~ "FL",
  combined_data$state == "Georgia" ~ "GA",
  combined_data$state == "Hawaii" ~ "HI",
  combined_data$state == "Idaho" ~ "ID",
  combined_data$state == "Illinois" ~ "IL",
  combined_data$state == "Indiana" ~ "IN",
  combined_data$state == "Iowa" ~ "IA",
  combined_data$state == "Kansas" ~ "KS",
  combined_data$state == "Kentucky" ~ "KY",
  combined_data$state == "Louisiana" ~ "LA",
  combined_data$state == "Maine" ~ "ME",
  combined_data$state == "Maryland" ~ "MD",
  combined_data$state == "Massachusetts" ~ "MA",
  combined_data$state == "Michigan" ~ "MI",
  combined_data$state == "Minnesota" ~ "MN",
  combined_data$state == "Mississippi" ~ "MS",
  combined_data$state == "Missouri" ~ "MO",
  combined_data$state == "Montana" ~ "MT",
  combined_data$state == "Nebraska" ~ "NE",
  combined_data$state == "Nevada" ~ "NV",
  combined_data$state == "New Hampshire" ~ "NH",
  combined_data$state == "New Jersey" ~ "NJ",
  combined_data$state == "New Mexico" ~ "NM",
  combined_data$state == "New York" ~ "NY",
  combined_data$state == "North Carolina" ~ "NC",
  combined_data$state == "North Dakota" ~ "ND",
  combined_data$state == "Ohio" ~ "OH",
  combined_data$state == "Oklahoma" ~ "OK",
  combined_data$state == "Oregon" ~ "OR",
  combined_data$state == "Pennsylvania" ~ "PA",
  combined_data$state == "Rhode Island" ~ "RI",
  combined_data$state == "South Carolina" ~ "SC",
  combined_data$state == "South Dakota" ~ "SD",
  combined_data$state == "Tennessee" ~ "TN",
  combined_data$state == "Texas" ~ "TX",
  combined_data$state == "Utah" ~ "UT",
  combined_data$state == "Vermont" ~ "VT",
  combined_data$state == "Virginia" ~ "VA",
  combined_data$state == "Washington" ~ "WA",
  combined_data$state == "West Virginia" ~ "WV",
  combined_data$state == "Wisconsin" ~ "WI",
  combined_data$state == "Wyoming" ~ "WY",
  TRUE ~ combined_data$state
)

# Create a new data frame 'state_cleaned_data' with cleaned variable
state_cleaned_data <- data.frame(State = combined_data$state)

# Assign the cleaned 'state' variable to 'combined_data'
combined_data$state <- state_cleaned_data$State

# Create a data frame with state counts
state_counts <- combined_data %>% 
  count(state)

# Create the pie chart
ggplot(state_counts, aes(x = "", y = n, fill = state)) +
  geom_bar(stat = "identity", width = 1) +
  coord_polar("y", start = 0) +
  labs(title = "Distribution of States", fill = "State") +
  theme_void() +
  theme(legend.position = "bottom")

# Create a new data frame for leave_people_cleaned_data
leave_people_cleaned_data <- combined_data %>%
  transmute(work_interfere = ifelse(work_interfere == 0, "Never", work_interfere),
            leave = ifelse(leave == 0, "Never", leave),
            coworkers,
            supervisor,
            comments) %>%
  mutate_all(~ ifelse(. %in% c("", "0"), NA, as.character(.)))

# Add a comment describing the options for each variable
comment(leave_people_cleaned_data$work_interfere) <- "Accepted values: Never, Sometimes, Often, Rarely"
comment(leave_people_cleaned_data$leave) <- "Accepted values: Very easy, Somewhat easy, Don't know, Somewhat difficult, Very difficult"
comment(leave_people_cleaned_data$coworkers) <- "Accepted values: Yes, No, Some of them"
comment(leave_people_cleaned_data$supervisor) <- "Accepted values: Yes, No, Some of them"
comment(leave_people_cleaned_data$comments) <- "Additional comments provided by the participants"

# Display the leave_people_cleaned_data dataframe in the global environment
leave_people_cleaned_data
# Create a new data frame for yes_no_cleaned_data
yes_no_cleaned_data <- combined_data[, c("self_employed", "family_history", "treatment", "remote_work", "tech_company",
                                         "benefits", "care_options", "wellness_program", "seek_help", "anonymity",
                                         "mental_health_consequence", "phys_health_consequence", "mental_health_interview",
                                         "phys_health_interview", "mental_vs_physical", "obs_consequence")]

# Clean the variables in yes_no_cleaned_data
yes_no_cleaned_data$self_employed <- ifelse(yes_no_cleaned_data$self_employed %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$self_employed, NA)
yes_no_cleaned_data$family_history <- ifelse(yes_no_cleaned_data$family_history %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$family_history, NA)
yes_no_cleaned_data$treatment <- ifelse(yes_no_cleaned_data$treatment %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$treatment, NA)
yes_no_cleaned_data$remote_work <- ifelse(yes_no_cleaned_data$remote_work %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$remote_work, NA)
yes_no_cleaned_data$tech_company <- ifelse(yes_no_cleaned_data$tech_company %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$tech_company, NA)
yes_no_cleaned_data$benefits <- ifelse(yes_no_cleaned_data$benefits %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$benefits, NA)
yes_no_cleaned_data$care_options <- ifelse(yes_no_cleaned_data$care_options %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$care_options, NA)
yes_no_cleaned_data$wellness_program <- ifelse(yes_no_cleaned_data$wellness_program %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$wellness_program, NA)
yes_no_cleaned_data$seek_help <- ifelse(yes_no_cleaned_data$seek_help %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$seek_help, NA)
yes_no_cleaned_data$anonymity <- ifelse(yes_no_cleaned_data$anonymity %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$anonymity, NA)
yes_no_cleaned_data$mental_health_consequence <- ifelse(yes_no_cleaned_data$mental_health_consequence %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$mental_health_consequence, NA)
yes_no_cleaned_data$phys_health_consequence <- ifelse(yes_no_cleaned_data$phys_health_consequence %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$phys_health_consequence, NA)
yes_no_cleaned_data$mental_health_interview <- ifelse(yes_no_cleaned_data$mental_health_interview %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$mental_health_interview, NA)
yes_no_cleaned_data$phys_health_interview <- ifelse(yes_no_cleaned_data$phys_health_interview %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$phys_health_interview, NA)
yes_no_cleaned_data$mental_vs_physical <- ifelse(yes_no_cleaned_data$mental_vs_physical %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$mental_vs_physical, NA)
yes_no_cleaned_data$obs_consequence <- ifelse(yes_no_cleaned_data$obs_consequence %in% c("Yes", "No", "Don't Know", "Maybe", "Not sure"), yes_no_cleaned_data$obs_consequence, NA)

# Describe the options for each variable
comment(yes_no_cleaned_data$self_employed) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$family_history) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$treatment) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$remote_work) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$tech_company) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$benefits) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$care_options) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$wellness_program) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$seek_help) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$anonymity) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$mental_health_consequence) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$phys_health_consequence) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$mental_health_interview) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$phys_health_interview) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$mental_vs_physical) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"
comment(yes_no_cleaned_data$obs_consequence) <- "Accepted values: Yes, No, Don't Know, Maybe, Not sure"

# Replace missing values and dates with NA in no_employees variable
employees_cleaned_data <- combined_data %>%
  select(no_employees) %>%
  mutate(no_employees = ifelse(is.na(as.character(no_employees)) | grepl("\\d{1,2}/\\d{1,2}/\\d{2,4}", no_employees), NA, no_employees))

# Check if the Gender column exists and is non-empty
if ("Gender" %in% names(combined_data) && length(combined_data$Gender) > 0) {
  # Trim whitespace in Gender column
  combined_data$Gender <- trimws(combined_data$Gender)
  
  # Replace missing values, unknown values, and invalid values with NA
  combined_data$Gender[combined_data$Gender == "" | combined_data$Gender == "unknown"] <- NA
  
  # Replace F/M or spelling mistakes with Male or Female
  combined_data$Gender <- gsub("(?i)^(f$|fem.*)", "Female", combined_data$Gender)
  combined_data$Gender <- gsub("(?i)^(m$|masc.*|man.*)", "Male", combined_data$Gender)
  
  # Create gender_cleaned_data data frame
  gender_cleaned_data <- data.frame(Gender = combined_data$Gender, stringsAsFactors = FALSE)
} else {
  # If Gender column is empty or not found, create an empty data frame
  gender_cleaned_data <- data.frame(Gender = character(), stringsAsFactors = FALSE)
}

# Check if the Timestamp column exists and is non-empty
if ("Timestamp" %in% names(combined_data) && length(combined_data$Timestamp) > 0) {
  # Trim whitespace in Timestamp column
  combined_data$Timestamp <- trimws(combined_data$Timestamp)
  
  # Define the desired format as a regular expression pattern
  format_pattern <- "^\\d{1,2}/\\d{1,2}/\\d{2} \\d{1,2}:\\d{2}$"
  
  # Replace missing values with NA
  combined_data$Timestamp[combined_data$Timestamp == ""] <- NA
  
  # Identify values that match the desired format
  format_matched <- grepl(format_pattern, combined_data$Timestamp, ignore.case = TRUE)
  
  # Create timestamp_cleaned_data data frame with the original values and NA for non-matching values
  timestamp_cleaned_data <- data.frame(Timestamp = ifelse(format_matched, combined_data$Timestamp, NA), stringsAsFactors = FALSE)
} else {
  # If Timestamp column is empty or not found, create an empty data frame
  timestamp_cleaned_data <- data.frame(Timestamp = character(), stringsAsFactors = FALSE)
}

# Export the cleaned dataset to an Excel file
write.xlsx(combined_data, file = "cleaned_data.xlsx", rowNames = FALSE)
