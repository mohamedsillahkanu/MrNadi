### Step 1: Install packages
Install packages that support data manipulation, reading and writing Excel files.

`dplyr`: Provide tools for data manipulation

`tidyr`: Offers functions for reshaping and tidying data

`writexl`: Enable writing data to Excel files

`readxl`: Enables reading Excel files

```r
install.packages("dplyr")
install.packages("tidyr")
install.packages("writexl")
install.packages("readxl")
```
### Step 2: Load packages
Load the installed libraries to make their functions available in the current session.
```r
library(dplyr)
library(tidyr)
library(writexl)
library(readxl)
```
### Step 3: Download the files and set working directory
- #### Step 3.1: Download the files here
  The file below is a zip file. Download and then extract. Copy that file path and paste in the next step.
  [routine data](https://github.com/ahadi-analytics/snt-data-files/raw/99c7f670f38e46d6e872c941e45e11dd68844d4b/Routine%20data.zip)
- #### Step 3.2: Set the working directory
Specify the folder location where the Excel files are stored so R can acess them easily. 
To adapt this code replace with actual file path
```r
setwd("C:/Users/User/Downloads/routine data") 
```
### Step 4: Get list of Excel files
Create a list of all Excel file paths in the working directory using a regular expression for `.xlsx`.
```r
file_paths <- list.files(pattern = "\\.xlsx$", full.names = TRUE)
```
### Step 5: Read the files to data frames
Use a loop to read each file into R and store all data frames in a list
```r
data_files <- lapply(file_paths, read_excel)
```
### Step 6: Combine data frames
Combine all data frames into one.
```r  
df <- bind_rows(data_files)
```
### Step 7: View combined data 
Check the structure of the combined dataset by printing the first few rows.
```r
print(head(df))
```
### Step 8: Rename columns using rename
Rename long column names to a short and standard naming convention. This varies by country. For this purpose of this work, Sierra Leone data dictionary is being used to do the renaming.
```r
df <- df %>%
  rename(
    adm0 = orgunitlevel1,
    adm1 = orgunitlevel2,
    adm2 = orgunitlevel3,
    adm3 = orgunitlevel4,
    hf   = organisationunitname,
    allout_u5   = OPD (New and follow-up curative) 0-59m_X,
    allout_ov5  = OPD (New and follow-up curative) 5+y_X
)
```
### Step 9: Split periodname into year and month
Separate `periodname` (e.g. , "April 2023") into two columns 
```r
df <- separate(df, col=periodname, into=c("month", "year"), sep = " ")
```
### Step 10: Recode month to numeric values
```r
df$month <- recode(df$month, 
  "January" = "01",
  "February" = "02",
  "March" = "03",
  "April" = "04",
  "May" = "05",
  "June" = "06",
  "July" = "07",
  "August" = "08",
  "September" = "09",
  "October" = "10",
  "November" = "11",
  "December" = "12"
)
```
### Step 11: Combine `year` and `month` into one column
```r
df <- df %>%
  unite("Date", Year, Month, sep = "-", remove = FALSE)
```
### Step 12: Create new variables 
```r
df$allout <- rowSums(df[c("allout_u5", "allout_ov5")], na.rm = TRUE)
df$susp <- rowSums(df[c(
  "susp_u5_hf", "susp_5_14_hf", "susp_ov15_hf",
  "susp_u5_com", "susp_5_14_com", "susp_ov15_com"
)], na.rm = TRUE)

# [See all variables in the full code]
```
### Step 13: Agregrate the data monthly at adm3
```r
monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year, month, date) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")
```
### Step 14: Save the agrregated monthly data at adm3 level
```r
write_excel(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
```
### Step 15: Aggregrate the data yearly at adm3 level
```r
monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")
```
### Step 16: Save the agreegated yearly data at adm3 level
```r
write_excel(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
```

## Full code
```r
install.packages("dplyr")
install.packages("tidyr")
install.packages("writexl")
install.packages("readxl")

library(dplyr)
library(tidyr)
library(writexl)
library(readxl)

setwd("path/to/your/directory") 

file_paths <- list.files(pattern = "\\.xlsx$", full.names = TRUE)

data_files <- lapply(file_paths, read_excel)

df <- bind_rows(data_files)

print(head(df))


df <- df %>%
  rename(
    adm0 = orgunitlevel1,
    adm1 = orgunitlevel2,
    adm2 = orgunitlevel3,
    adm3 = orgunitlevel4,
    hf   = organisationunitname,
    allout_u5   = OPD (New and follow-up curative) 0-59m_X,
    allout_ov5  = OPD (New and follow-up curative) 5+y_X,

    maladm_u5   = Admission - Child with malaria 0-59 months_X,
    maladm_5_14 = Admission - Child with malaria 5-14 years_X,
    maladm_ov15 = Admission - Malaria 15+ years_X,

    maldth_1_59m = Child death - Malaria 1-59m_X,
    maldth_10_14 = Child death - Malaria 10-14y_X,
    maldth_5_9   = Child death - Malaria 5-9y_X,
    maldth_fem_ov15 = Death malaria 15+ years Female,
    maldth_mal_ov15 = Death malaria 15+ years Male,
    maldth_u5   = Separation - Child with malaria 0-59 months_X Death,
    maldth_5_14 = Separation - Child with malaria 5-14 years_X Death,
    maldth_ov15 = Separation - Malaria 15+ years_X Death,

    susp_u5_hf     = Fever case - suspected Malaria 0-59m_X,
    susp_5_14_hf   = Fever case - suspected Malaria 5-14y_X`,
    susp_ov15_hf   = Fever case - suspected Malaria 15+y_X`,
    susp_u5_com    = Fever case in community (Suspected Malaria) 0-59m_X,
    susp_5_14_com  = Fever case in community (Suspected Malaria) 5-14y_X,
    susp_ov15_com  = Fever case in community (Suspected Malaria) 15+y_X,
  
    tes_neg_rdt_u5_com   = Fever case in community tested for Malaria (RDT) - Negative 0-59m_X,
    tes_pos_rdt_u5_com   = Fever case in community tested for Malaria (RDT) - Positive 0-59m_X,
    tes_neg_rdt_5_14_com = Fever case in community tested for Malaria (RDT) - Negative 5-14y_X,
    tes_pos_rdt_5_14_com = Fever case in community tested for Malaria (RDT) - Positive 5-14y_X,
    tes_neg_rdt_ov15_com = Fever case in community tested for Malaria (RDT) - Negative 15+y_X,
    test_neg_mic_u5_hf   = Fever case tested for Malaria (Microscopy) - Negative 0-59m_X,
    test_pos_mic_u5_hf   = Fever case tested for Malaria (Microscopy) - Positive 0-59m_X,
    test_neg_mic_5_14_hf = Fever case tested for Malaria (Microscopy) - Negative 5-14y_X,
    test_pos_mic_5_14_hf = Fever case tested for Malaria (Microscopy) - Positive 5-14y_X,
    test_neg_mic_ov15_hf = Fever case tested for Malaria (Microscopy) - Negative 15+y_X,
    test_pos_mic_ov15_hf = Fever case tested for Malaria (Microscopy) - Positive 15+y_X,

    maltreat_u24_u5_com   = Malaria treated in community with ACT <24 hours 0-59m_X,
    maltreat_ov24_u5_com  = Malaria treated in community with ACT >24 hours 0-59m_X,
    maltreat_u24_5_14_com = Malaria treated in community with ACT <24 hours 5-14y_X,
    maltreat_ov24_5_14_com = Malaria treated in community with ACT >24 hours 5-14y_X,
    maltreat_u24_ov15_com = Malaria treated in community with ACT <24 hours 15+y_X,
    maltreat_ov24_ov15_com = Malaria treated in community with ACT >24 hours 15+y_X,
  
    maltreat_u24_u5_hf    = Malaria treated with ACT <24 hours 0-59m_X,
    maltreat_ov24_u5_hf   = Malaria treated with ACT >24 hours 0-59m_X,
    maltreat_u24_5_14_hf  = Malaria treated with ACT <24 hours 5-14y_X,
    maltreat_ov24_5_14_hf = Malaria treated with ACT >24 hours 5-14y_X,
    maltreat_u24_ov15_hf  = Malaria treated with ACT <24 hours 15+y_X,
    maltreat_ov24_ov15_hf = Malaria treated with ACT >24 hours 15+y_X
)

df <- separate(df, col=periodname, into=c("month", "year"), sep = " ")

df$month <- recode(df$month, 
  "January" = "01",
  "February" = "02",
  "March" = "03",
  "April" = "04",
  "May" = "05",
  "June" = "06",
  "July" = "07",
  "August" = "08",
  "September" = "09",
  "October" = "10",
  "November" = "11",
  "December" = "12"
)

df <- df %>%
  unite("Date", Year, Month, sep = "-", remove = FALSE)

df$allout <- rowSums(df[c("allout_u5", "allout_ov5")], na.rm = TRUE)

df$susp <- rowSums(df[c(
  "susp_u5_hf", "susp_5_14_hf", "susp_ov15_hf",
  "susp_u5_com", "susp_5_14_com", "susp_ov15_com"
)], na.rm = TRUE)

df$test_hf <- rowSums(df[c(
  "test_neg_mic_u5_hf", "test_pos_mic_u5_hf",
  "test_neg_mic_5_14_hf", "test_pos_mic_5_14_hf",
  "test_neg_mic_ov15_hf", "test_pos_mic_ov15_hf",
  "tes_neg_rdt_u5_hf", "tes_pos_rdt_u5_hf",
  "tes_neg_rdt_5_14_hf", "tes_pos_rdt_5_14_hf",
  "tes_neg_rdt_ov15_hf", "tes_pos_rdt_ov15_hf"
)], na.rm = TRUE)

df$test_com <- rowSums(df[c(
  "tes_neg_rdt_u5_com", "tes_pos_rdt_u5_com",
  "tes_neg_rdt_5_14_com", "tes_pos_rdt_5_14_com",
  "tes_neg_rdt_ov15_com", "tes_pos_rdt_ov15_com"
)], na.rm = TRUE)

df$test <- rowSums(df[c("test_hf", "test_com")], na.rm = TRUE)

df$conf_hf <- rowSums(df[c(
  "test_pos_mic_u5_hf", "test_pos_mic_5_14_hf", "test_pos_mic_ov15_hf",
  "tes_pos_rdt_u5_hf", "tes_pos_rdt_5_14_hf", "tes_pos_rdt_ov15_hf"
)], na.rm = TRUE)

df$conf_com <- rowSums(df[c(
  "tes_pos_rdt_u5_com", "tes_pos_rdt_5_14_com", "tes_pos_rdt_ov15_com"
)], na.rm = TRUE)

df$conf <- rowSums(df[c("conf_hf", "conf_com")], na.rm = TRUE)

df$maltreat_com <- rowSums(df[c(
  "maltreat_u24_u5_com", "maltreat_ov24_u5_com",
  "maltreat_u24_5_14_com", "maltreat_ov24_5_14_com",
  "maltreat_u24_ov15_com", "maltreat_ov24_ov15_com"
)], na.rm = TRUE)

df$maltreat_hf <- rowSums(df[c(
  "maltreat_u24_u5_hf", "maltreat_ov24_u5_hf",
  "maltreat_u24_5_14_hf", "maltreat_ov24_5_14_hf",
  "maltreat_u24_ov15_hf", "maltreat_ov24_ov15_hf"
)], na.rm = TRUE)

df$maltreat <- rowSums(df[c("maltreat_hf", "maltreat_com")], na.rm = TRUE)

df$pres_com <- pmax(df$maltreat_com - df$conf_com, 0, na.rm = TRUE)

df$pres_hf <- pmax(df$maltreat_hf - df$conf_hf, 0, na.rm = TRUE)


df$maladm <- rowSums(df[c("maladm_u5", "maladm_5_14", "maladm_ov15")], na.rm = TRUE)

df$maldth <- rowSums(df[c(
  "maldth_u5", "maldth_1_59m", "maldth_10_14", "maldth_5_9",
  "maldth_5_14", "maldth_ov15", "maldth_fem_ov15", "maldth_mal_ov15"
)], na.rm = TRUE)

monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year, month, date) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")

write_excel(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")

monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")

write_excel(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
```

