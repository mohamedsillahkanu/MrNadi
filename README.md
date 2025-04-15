#===================================================================
#   Step 1: Install packages
# ==================================================================

install.packages("dplyr")
install.packages("tidyr")
install.packages("writexl")

# ==================================================================
#   Step 2: Load the libraries
# ==================================================================

library(dplyr)
library(writexl)

# ==================================================================
#   Step 3: Set working directory usinf setwd
# ==================================================================

setwd("path/to/your/directory")

# ==================================================================
#   Step 4: Get list of Excel files using list.files
# ==================================================================

file_paths <- list.files(pattern = "\\.xlsx$", full.names = TRUE)

# ==================================================================
#   Step 5: Read the files to a dataframe using lapply
# ==================================================================

data_files <- lapply(file_paths, read_excel)

# ==================================================================
#   Step 6: Combine into a single data frame using bind_rows
# ==================================================================
  
df <- bind_rows(data_files)

# ==================================================================
#   Step 7: View the combined dataframe
# ==================================================================
  
print(head(df))

# ==================================================================
#   Step 8: Rename columns using rename
# ==================================================================

df <- df %>%
  rename(
    adm0 = orgunitlevel1,
    adm1 = orgunitlevel2,
    adm2 = orgunitlevel3,
    adm3 = orgunitlevel4,
    hf   = organisationunitname,

    allout_u5   = `OPD (New and follow-up curative) 0-59m_X`,
    allout_ov5  = `OPD (New and follow-up curative) 5+y_X`,

    maladm_u5   = `Admission - Child with malaria 0-59 months_X`,
    maladm_5_14 = `Admission - Child with malaria 5-14 years_X`,
    maladm_ov15 = `Admission - Malaria 15+ years_X`,

    maldth_1_59m = `Child death - Malaria 1-59m_X`,
    maldth_10_14 = `Child death - Malaria 10-14y_X`,
    maldth_5_9   = `Child death - Malaria 5-9y_X`,

    maldth_fem_ov15 = `Death malaria 15+ years Female`,
    maldth_mal_ov15 = `Death malaria 15+ years Male`,

    maldth_u5   = `Separation - Child with malaria 0-59 months_X Death`,
    maldth_5_14 = `Separation - Child with malaria 5-14 years_X Death`,
    maldth_ov15 = `Separation - Malaria 15+ years_X Death`,

    susp_u5_hf     = `Fever case - suspected Malaria 0-59m_X`,
    susp_5_14_hf   = `Fever case - suspected Malaria 5-14y_X`,
    susp_ov15_hf   = `Fever case - suspected Malaria 15+y_X`,
    susp_u5_com    = `Fever case in community (Suspected Malaria) 0-59m_X`,
    susp_5_14_com  = `Fever case in community (Suspected Malaria) 5-14y_X`,
    susp_ov15_com  = `Fever case in community (Suspected Malaria) 15+y_X`,

    tes_neg_rdt_u5_com   = `Fever case in community tested for Malaria (RDT) - Negative 0-59m_X`,
    tes_pos_rdt_u5_com   = `Fever case in community tested for Malaria (RDT) - Positive 0-59m_X`,
    tes_neg_rdt_5_14_com = `Fever case in community tested for Malaria (RDT) - Negative 5-14y_X`,
    tes_pos_rdt_5_14_com = `Fever case in community tested for Malaria (RDT) - Positive 5-14y_X`,
    tes_neg_rdt_ov15_com = `Fever case in community tested for Malaria (RDT) - Negative 15+y_X`,
    tes_pos_rdt_ov15_com = `Fever case in community tested for Malaria (RDT) - Positive 15+y_X`,

    test_neg_mic_u5_hf   = `Fever case tested for Malaria (Microscopy) - Negative 0-59m_X`,
    test_pos_mic_u5_hf   = `Fever case tested for Malaria (Microscopy) - Positive 0-59m_X`,
    test_neg_mic_5_14_hf = `Fever case tested for Malaria (Microscopy) - Negative 5-14y_X`,
    test_pos_mic_5_14_hf = `Fever case tested for Malaria (Microscopy) - Positive 5-14y_X`,
    test_neg_mic_ov15_hf = `Fever case tested for Malaria (Microscopy) - Negative 15+y_X`,
    test_pos_mic_ov15_hf = `Fever case tested for Malaria (Microscopy) - Positive 15+y_X`,

    tes_neg_rdt_u5_hf    = `Fever case tested for Malaria (RDT) - Negative 0-59m_X`,
    tes_pos_rdt_u5_hf    = `Fever case tested for Malaria (RDT) - Positive 0-59m_X`,
    tes_neg_rdt_5_14_hf  = `Fever case tested for Malaria (RDT) - Negative 5-14y_X`,
    tes_pos_rdt_5_14_hf  = `Fever case tested for Malaria (RDT) - Positive 5-14y_X`,
    tes_neg_rdt_ov15_hf  = `Fever case tested for Malaria (RDT) - Negative 15+y_X`,
    tes_pos_rdt_ov15_hf  = `Fever case tested for Malaria (RDT) - Positive 15+y_X`,

    maltreat_u24_u5_com   = `Malaria treated in community with ACT <24 hours 0-59m_X`,
    maltreat_ov24_u5_com  = `Malaria treated in community with ACT >24 hours 0-59m_X`,
    maltreat_u24_5_14_com = `Malaria treated in community with ACT <24 hours 5-14y_X`,
    maltreat_ov24_5_14_com = `Malaria treated in community with ACT >24 hours 5-14y_X`,
    maltreat_u24_ov15_com = `Malaria treated in community with ACT <24 hours 15+y_X`,
    maltreat_ov24_ov15_com = `Malaria treated in community with ACT >24 hours 15+y_X`,

    maltreat_u24_u5_hf    = `Malaria treated with ACT <24 hours 0-59m_X`,
    maltreat_ov24_u5_hf   = `Malaria treated with ACT >24 hours 0-59m_X`,
    maltreat_u24_5_14_hf  = `Malaria treated with ACT <24 hours 5-14y_X`,
    maltreat_ov24_5_14_hf = `Malaria treated with ACT >24 hours 5-14y_X`,
    maltreat_u24_ov15_hf  = `Malaria treated with ACT <24 hours 15+y_X`,
    maltreat_ov24_ov15_hf = `Malaria treated with ACT >24 hours 15+y_X`
  )


# ==================================================================
#   Step 9: Split periodname into Year and Month usinf separate
# ==================================================================

df <- separate(df, col=periodname, into=c("month", "year"), sep = " ")

# ==================================================================
#   Step 10: Recode Month to numeric values using recode
# ==================================================================
  
df$Month <- recode(df$Month, 
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

# ==================================================================
#   Step 11: Combine 'Year' and 'Month' into one column using unite
# ==================================================================

df <- df %>%
  unite("Date", Year, Month, sep = "-", remove = FALSE)

# ==================================================================
#   Step 12: Create new variables using rowSums
# ==================================================================

# allout
df$allout <- rowSums(df[c("allout_u5", "allout_ov5")], na.rm = TRUE)

# susp
df$susp <- rowSums(df[c(
  "susp_u5_hf", "susp_5_14_hf", "susp_ov15_hf",
  "susp_u5_com", "susp_5_14_com", "susp_ov15_com"
)], na.rm = TRUE)

# test_hf
df$test_hf <- rowSums(df[c(
  "test_neg_mic_u5_hf", "test_pos_mic_u5_hf",
  "test_neg_mic_5_14_hf", "test_pos_mic_5_14_hf",
  "test_neg_mic_ov15_hf", "test_pos_mic_ov15_hf",
  "tes_neg_rdt_u5_hf", "tes_pos_rdt_u5_hf",
  "tes_neg_rdt_5_14_hf", "tes_pos_rdt_5_14_hf",
  "tes_neg_rdt_ov15_hf", "tes_pos_rdt_ov15_hf"
)], na.rm = TRUE)

# test_com
df$test_com <- rowSums(df[c(
  "tes_neg_rdt_u5_com", "tes_pos_rdt_u5_com",
  "tes_neg_rdt_5_14_com", "tes_pos_rdt_5_14_com",
  "tes_neg_rdt_ov15_com", "tes_pos_rdt_ov15_com"
)], na.rm = TRUE)

# test
df$test <- rowSums(df[c("test_hf", "test_com")], na.rm = TRUE)

# conf_hf
df$conf_hf <- rowSums(df[c(
  "test_pos_mic_u5_hf", "test_pos_mic_5_14_hf", "test_pos_mic_ov15_hf",
  "tes_pos_rdt_u5_hf", "tes_pos_rdt_5_14_hf", "tes_pos_rdt_ov15_hf"
)], na.rm = TRUE)

# conf_com
df$conf_com <- rowSums(df[c(
  "tes_pos_rdt_u5_com", "tes_pos_rdt_5_14_com", "tes_pos_rdt_ov15_com"
)], na.rm = TRUE)

# conf
df$conf <- rowSums(df[c("conf_hf", "conf_com")], na.rm = TRUE)

# maltreat_com
df$maltreat_com <- rowSums(df[c(
  "maltreat_u24_u5_com", "maltreat_ov24_u5_com",
  "maltreat_u24_5_14_com", "maltreat_ov24_5_14_com",
  "maltreat_u24_ov15_com", "maltreat_ov24_ov15_com"
)], na.rm = TRUE)

# maltreat_hf
df$maltreat_hf <- rowSums(df[c(
  "maltreat_u24_u5_hf", "maltreat_ov24_u5_hf",
  "maltreat_u24_5_14_hf", "maltreat_ov24_5_14_hf",
  "maltreat_u24_ov15_hf", "maltreat_ov24_ov15_hf"
)], na.rm = TRUE)

# maltreat
df$maltreat <- rowSums(df[c("maltreat_hf", "maltreat_com")], na.rm = TRUE)

# pres_com
df$pres_com <- pmax(df$maltreat_com - df$conf_com, 0, na.rm = TRUE)

# pres_hf
df$pres_hf <- pmax(df$maltreat_hf - df$conf_hf, 0, na.rm = TRUE)

# pres
df$pres <- rowSums(df[c("pres_com", "pres_hf")], na.rm = TRUE)

# maladm
df$maladm <- rowSums(df[c("maladm_u5", "maladm_5_14", "maladm_ov15")], na.rm = TRUE)

# maldth
df$maldth <- rowSums(df[c(
  "maldth_u5", "maldth_1_59m", "maldth_10_14", "maldth_5_9",
  "maldth_5_14", "maldth_ov15", "maldth_fem_ov15", "maldth_mal_ov15"
)], na.rm = TRUE)




# ==================================================================
#   Step 12: Agregrate the data monthly at adm3
# ==================================================================
  
monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year, month, date) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")


write_excel(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(monthly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")

# ==================================================================
#   Step 13: Agregrate the data yearly at adm3
# ==================================================================

monthly_adm2_data <- df %>%
  group_by(adm1, adm2, adm3, year) %>%
  summarise(across(-c(adm1, adm2, year, month), sum, na.rm = TRUE), .groups = "drop")


write_excel(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
write_csv(yearly_adm2_data, "/output file path/aggregated_data_monthly_adm2_data.xlsx")
