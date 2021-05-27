##################
# DATA PREPARATION
##################

# Set working directory
setwd("~/Desktop/thesis/Oljefondet")

# Load libraries
library(readxl)
library(dplyr)
library(pdftools)
library(tidyverse)
library(flextable)
library(ggrepel)
library(ggforce)
library(xlsx)
library(ggh4x)
library(ggpubr)
library(broom)
library(rstatix)
library(corrplot)
library(RColorBrewer)
library(stargazer)
library(car)

# Load data set on Pension Funds Equity Investments portfolio  
oljefondet <- read_xlsx("EQ_2019_Country.xlsx")

# Which countries the fund invests in?
country_split <- oljefondet %>%
  group_by(Country) %>%
  summarise(Frequency = n()) %>%
  arrange(desc(Frequency))

# How many countries the fund invests in?
nrow(country_split)

#######################
# TRACE Matrix data set
#######################

# Set working directory
setwd("~/Desktop/thesis/2020 TRACE Matrix Information pack")

# Load file
trace_matrix <- pdf_text("2020 TRACE Matrix Country Risk Scores.pdf") %>%
  read_lines()

# Remove headers
trace_matrix <- trace_matrix[c(18:34, 54:72, 92:112, 132:152, 172:190, 210:229, 
                               249:267, 287:306, 326:344, 364:384, 404:419)]

# Remove spaces
trace_matrix <- gsub("\\s+", " ", trace_matrix)

##############################
# Fix lines with multiple rows
#############################

# United Arab Emirates
trace_matrix[41] <- substring(trace_matrix[41], 5, nchar(trace_matrix[41]))

trace_matrix[40] <- paste(" 40", trace_matrix[40], trace_matrix[42]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[40] <- paste(trace_matrix[40], trace_matrix[41])

# St. Vincent and the Grenadines
trace_matrix[53] <- substring(trace_matrix[53], 5, nchar(trace_matrix[53]))

trace_matrix[52] <- paste(" 50", trace_matrix[52], trace_matrix[54]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[52] <- paste(trace_matrix[52], trace_matrix[53])

# Antigue and Barbua
trace_matrix[64] <- substring(trace_matrix[64], 5, nchar(trace_matrix[64]))

trace_matrix[63] <- paste(" 59", trace_matrix[63], trace_matrix[65]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[63] <- paste(trace_matrix[63], trace_matrix[64])

# Trinidad and tobago
trace_matrix[69] <- substring(trace_matrix[69], 5, nchar(trace_matrix[69]))

trace_matrix[68] <- paste(" 62", trace_matrix[68], trace_matrix[70]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[68] <- paste(trace_matrix[68], trace_matrix[69])

# Sao Tome and Principe
trace_matrix[101] <- substring(trace_matrix[101], 5, nchar(trace_matrix[101]))

trace_matrix[100] <- paste(" 92", trace_matrix[100], trace_matrix[102]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[100] <- paste(trace_matrix[100], trace_matrix[101])

# Bosnia and Hercegovina
trace_matrix[143] <- substring(trace_matrix[143], 5, nchar(trace_matrix[143]))

trace_matrix[142] <- paste(" 132", trace_matrix[142], trace_matrix[144]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[142] <- paste(trace_matrix[142], trace_matrix[143])

# Central African Republic
trace_matrix[192] <- substring(trace_matrix[192], 5, nchar(trace_matrix[192]))

trace_matrix[191] <- paste(" 179", trace_matrix[191], trace_matrix[193]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[191] <- paste(trace_matrix[191], trace_matrix[192])

# Dem rep of Congo
trace_matrix[195] <- substring(trace_matrix[195], 5, nchar(trace_matrix[195]))

trace_matrix[194] <- paste(" 180", trace_matrix[194], trace_matrix[196]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[194] <- paste(trace_matrix[194], trace_matrix[195])

# THE CONGO
trace_matrix[198] <- substring(trace_matrix[198], 5, nchar(trace_matrix[198]))

trace_matrix[197] <- paste(" 181", trace_matrix[197], trace_matrix[199]) %>% 
  gsub("\\s+", " ", .)

trace_matrix[197] <- paste(trace_matrix[197], trace_matrix[198])

# Remove empty rows
trace_matrix <- trace_matrix[c(1:40, 43:52, 55:63, 66:68, 
                               71:100, 103:142, 145:191, 194, 197, 200:212)]
# Remove numbering
trace_matrix <- sub("^[^a-zA-Z]*", "", trace_matrix)

# Extract risk scores from TRACE matrix
scores <- sub("^\\D+", "", trace_matrix)

# Extract country from TRACE matrix
countries <- trimws(sub("^(\\D+).*", "\\1", trace_matrix))

# Convert scores to a data frame
df <- data.matrix(as.list(scores))

# Convert 
trace_matrix <- as.data.frame(matrix(unlist(strsplit(scores, "[ ]")), 
                                     ncol = 14, byrow = TRUE))

# Move country names as the first variable
trace_matrix <- trace_matrix %>% 
  mutate(Country = countries) %>%
  relocate(Country)

# Add column names
colnames(trace_matrix) <- c("Country", "Total Risk score", "Overall Opportunity risk",
                            "Interaction", "Expectation", "Leverage", "Overall Deterrence risk",
                            "Dissuasion", "Enforcement", "Overall Transparency risk", "Processes",
                            "Interests", "Overall Oversight risk", "Free press", "Civil society")

###########################
# Synchronize country names
###########################

oljefondet$Country <- gsub(pattern = "Russia", 
                           replacement = "Russian Federation", 
                           x = oljefondet$Country)

oljefondet$Country <- gsub(pattern = "Kyrgyzstan", 
                           replacement = "Kyrgyz Republic", 
                           x = oljefondet$Country)

# Join together the Oljefondet data set and the TRACE matrix
df <- left_join(oljefondet, trace_matrix, by = "Country")

# Convert risk scores to numeric
df <- df %>%
  mutate(`Total Risk score` = as.numeric(df$`Total Risk score`))

# Add risk categories
df <- df %>%
  mutate(risk_cat = case_when(`Total Risk score` < 22 ~ "Very Low",
                              `Total Risk score` < 38 ~ "Low",
                              `Total Risk score` < 56 ~ "Moderate",
                              `Total Risk score` < 74 ~ "High",
                              `Total Risk score` > 74 ~ "Very high"))

# Average total risk score for firms in portfolio
df %>%
  select(`Total Risk score`) %>%
  sum(.) / nrow(df)

# Calculate total fund value
fund_value_oljefondet <- df %>%
  select(`Market Value(USD)`) %>%
  sum(.)

# Average fund value 
average_fund_value <- fund_value / nrow(df)

#########################################
# Add ISO code for df and trace_matrix
#########################################

# Load country ISO codes
iso_codes <- ISOcodes::ISO_3166_1

iso_codes <- iso_codes %>%
  select(Alpha_2, Name) %>%
  rename(Country = Name)

# Add ISO code to df dataframe
df <- left_join(df, iso_codes, by = "Country")

# Syncro country names in df
df_fix <- df %>%
  filter(is.na(Alpha_2 == TRUE)) %>%
  mutate(Alpha_2 = case_when(Country == "Czech Republic" ~ "CZ",
                             Country == "Kyrgyz Republic" ~ "KG",
                             Country == "Moldova" ~ "MD",
                             Country == "South Korea" ~ "KR",
                             Country == "Taiwan" ~ "TW",
                             Country == "Tanzania" ~ "TZ",
                             Country == "Vietnam" ~ "VN"))
df <- df %>%
  filter(!is.na(Alpha_2 == TRUE))

df <- rbind(df, df_fix)

# Add ISO code to trace_matrix dataframe
trace_matrix <- left_join(trace_matrix, iso_codes, by = "Country")

# Fix missing ISO codes
trace_matrix_fix <- trace_matrix %>%
  filter(is.na(Alpha_2 == TRUE)) %>%
  mutate(Alpha_2 = case_when(Country == "Czech Republic" ~ "CZ",
                             Country == "Kyrgyz Republic" ~ "KG",
                             Country == "Moldova" ~ "MD",
                             Country == "South Korea" ~ "KR",
                             Country == "North Korea" ~ "KR",
                             Country == "Taiwan" ~ "TW",
                             Country == "Tanzania" ~ "TZ",
                             Country == "Vietnam" ~ "VN",
                             Country == "St. Vincent and the Grenadines" ~ "VC",
                             Country == "Slovak Republic" ~ "SK",
                             Country == "St. Lucia" ~ "LC",
                             Country == "St. Kitts and Nevis" ~ "KN",
                             Country == "Cape Verde" ~ "CV",
                             Country == "Kosovo" ~ "XK",
                             Country == "Micronesia" ~ "FM",
                             Country == "East Timor" ~ "TL",
                             Country == "Ivory Coast" ~ "CI",
                             Country == "Bolivia" ~ "BO",
                             Country == "Iran" ~ "IR",
                             Country == "Dem. Rep. of the Congo" ~ "CD",
                             Country == "Republic of the Congo (Brazzaville)" ~ "CG",
                             Country == "Laos" ~ "LA",
                             Country == "Syria" ~ "SY",
                             Country == "Venezuela" ~ "VE"))

trace_matrix <- trace_matrix %>%
  filter(!is.na(Alpha_2 == TRUE))

trace_matrix <- rbind(trace_matrix, trace_matrix_fix)

#############################################################
# Create dataframe for Observation and exclusion of companies
############################################################

# Read in data
exclusion <- readLines("https://www.nbim.no/en/the-fund/responsible-investment/exclusion-of-companies")


# Check index for first and last company
grep("Aboitiz Power Corp", exclusion)
grep("Zuari Agro Chemicals Ltd", exclusion)

# Narrow down the data
exclusion <- exclusion[144:644]

# Remove html tags
exclusion <-gsub("<.*?>","", exclusion)
exclusion <-gsub("&nbsp;","", exclusion)
exclusion <-gsub(" &amp;","", exclusion)

# Remove empty elements
exclusion <- exclusion[exclusion != ""]

# Remove Company columns name
exclusion <- gsub(".*Company", "", exclusion)

# Extract every second element (company names)
firms <- exclusion[seq(1, length(exclusion), 2)]

# Remove Category columns (not important)
exclusion <-gsub("Product-basedCategory","", exclusion)
exclusion <-gsub("Product basedCategory","", exclusion)
exclusion <-gsub("Conduct-basedCategory","", exclusion)
exclusion <-gsub("CategoryConduct-based","", exclusion)
exclusion <-gsub("CategoryProduct-based","", exclusion)

# Extract Criterion columns
criterion <- exclusion[seq(2, length(exclusion), 2)] %>% 
  gsub("Criterion.*", "", .)

# Fix mistakes
criterion <- gsub("Former GenCorp Inc", "Production of nuclear weapons", criterion)
criterion <- gsub("Former EADS Finance BV", "Production of nuclear weapons", criterion)
criterion <- gsub("Former EADS NV", "Production of nuclear weapons", criterion)
criterion <- gsub("Former Babcock Wilcox Co", "Production of nuclear weapons", criterion)
criterion <- gsub("Incl. subsidiaries Duke Energy Carolinas LLC, Duke Energy Progress LLC and Progress Energy Inc",
                  "Severe environmental damage", criterion)
criterion <- gsub("Former Posco Daewoo Corp", "Severe environmental damage", criterion)
criterion <- gsub("Former Sesa Sterlite", "Severe environmental damage", criterion)

criterion <- gsub(" Production of coal or coal-based energy", 
                  "Production of coal or coal-based energy", criterion)

criterion <- gsub(" Production of nuclear weapons", 
                  "Production of nuclear weapons", criterion)

# Extract Decision columns
decision <- gsub(".*Criterion", "", exclusion[seq(2, length(exclusion), 2)]) %>% 
  gsub("Decision.*", "", .)

# Fix mistakes
decision <- gsub("Production of nuclear weapons", "Exclusion", decision)
decision <- gsub("Severe environmental damage", "Exclusion", decision)

# Extract last 10 characters for publishing date and year
date <- sub('.*(?=.{10}$)', '', exclusion[seq(2, length(exclusion), 2)], perl=T)
year <- as.numeric(sub('.*(?=.{4}$)', '', date, perl=T))

# Create data frame
exclusion <- data.frame(firms, criterion, decision, date, year)

# Rename columns
colnames(exclusion) <- c("Firm name", "Criterion", "Decision", "Publishing date", "Year")

#############
# Data on CPI
#############

# set WD
setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Robustness check")

# Read in file
df_cpi <- read_xlsx("CPI2020.xlsx", sheet = 1)

# Normalize scores
df_cpi <- df_cpi %>%
  select(Country, `CPI score 2020`) %>%
  mutate(CPI_score = 100 - `CPI score 2020`) %>%
  select(-`CPI score 2020`)

# Syncronise country names
df_cpi <- df_cpi %>% 
  mutate(Country = str_replace(Country, "United States of America", "United States"), 
         Country = str_replace(Country, "Czechia", "Czech Republic"),
         Country = str_replace(Country, "Russia", "Russian Federation"),
         Country = str_replace(Country, "Korea, South", "South Korea"),
         Country = str_replace(Country, "Kyrgyzstan", "Kyrgyz Republic"))

##################
# data on WB - CoC
##################

# set WD
setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Robustness check")

# Read in file
df_coc <- read_xlsx("WB_Control of Corruption.xlsx", sheet = 1)

# Select variables
df_coc <- df_coc %>%
  select(`Country Name`, CoC_score) %>%
  rename(Country = `Country Name`)

# Syncronise country names
df_coc <- df_coc %>% 
  mutate(Country = str_replace(Country, "Egypt, Arab Rep.", "Egypt"), 
         Country = str_replace(Country, "Hong Kong SAR, China", "Hong Kong"),
         Country = str_replace(Country, "Taiwan, China", "Taiwan"),
         Country = str_replace(Country, "Korea, Dem. People’s Rep.", "South Korea"))

############
# EIU - GDI
############

df_gdi <- read_xlsx("_EIU-Democracy Indices - Dataset - v4.xlsx", 
                    sheet = "data-for-countries-etc-by-year")

df_gdi <- df_gdi %>%
  filter(time == "2019") %>%
  select(name, `Democracy index (EIU)`) %>%
  mutate(GDI_score = 100 - `Democracy index (EIU)`) %>%
  rename(Country = name) %>%
  select(-`Democracy index (EIU)`)

# Syncronise country names
df_gdi <- df_gdi %>% 
  mutate(Country = str_replace(Country, "Hong Kong, China", "Hong Kong"), 
         Country = str_replace(Country, "Russia", "Russian Federation"))

########################
# Data about Oljefondet
#######################

# Domains
df$`Overall Opportunity risk` <- as.numeric(df$`Overall Opportunity risk`)
df$`Overall Deterrence risk` <- as.numeric(df$`Overall Deterrence risk`)
df$`Overall Transparency risk` <- as.numeric(df$`Overall Transparency risk`)
df$`Overall Oversight risk` <- as.numeric(df$`Overall Oversight risk`)

# Sub-domains
df$Interaction <- as.numeric(df$Interaction)
df$Expectation <- as.numeric(df$Expectation)
df$Leverage <- as.numeric(df$Leverage)
df$Dissuasion <- as.numeric(df$Dissuasion)
df$Enforcement <- as.numeric(df$Enforcement)
df$Processes <- as.numeric(df$Processes)
df$Interests <- as.numeric(df$Interests)
df$`Free press` <- as.numeric(df$`Free press`)
df$`Civil society` <- as.numeric(df$`Civil society`)

# CPI
df <- left_join(df, df_cpi, by = "Country")
df$CPI_score <- as.numeric(df$CPI_score)

# CoC
df <- left_join(df, df_coc, by = "Country")
df$CoC_score <- as.numeric(df$CoC_score)

# GDI
df <- left_join(df, df_gdi, by = "Country")
df$GDI_score <- as.numeric(df$GDI_score)


df <- df %>%
  mutate(investment_percent = `Market Value(USD)` / fund_value_oljefondet,
         investment_risk_score_1 = investment_percent * `Total Risk score`,
         investment_risk_score_2 = investment_percent * `Overall Deterrence risk`,
         investment_risk_score_3 = investment_percent * `Overall Transparency risk`,
         investment_risk_score_4 = investment_percent * `Overall Oversight risk`,
         investment_risk_score_5 = investment_percent * `Overall Opportunity risk`,
         investment_risk_score_6 = investment_percent * Interaction,
         investment_risk_score_7 = investment_percent * Expectation,
         investment_risk_score_8 = investment_percent * Leverage,
         investment_risk_score_9 = investment_percent * Dissuasion,
         investment_risk_score_10 = investment_percent * Enforcement,
         investment_risk_score_11 = investment_percent * Processes,
         investment_risk_score_12 = investment_percent * Interests,
         investment_risk_score_13 = investment_percent * `Free press`,
         investment_risk_score_14 = investment_percent * `Civil society`,
         investment_risk_score_15 = investment_percent * CPI_score,
         investment_risk_score_16 = investment_percent * CoC_score,
         investment_risk_score_17 = investment_percent * GDI_score)

portfolio_oljefondet <- data.frame(rbind(sum(df$investment_risk_score_1), sum(df$investment_risk_score_2), 
                                         sum(df$investment_risk_score_3), sum(df$investment_risk_score_4), 
                                         sum(df$investment_risk_score_5), sum(df$investment_risk_score_6), 
                                         sum(df$investment_risk_score_7), sum(df$investment_risk_score_8),
                                         sum(df$investment_risk_score_9), sum(df$investment_risk_score_10),
                                         sum(df$investment_risk_score_11), sum(df$investment_risk_score_12),
                                         sum(df$investment_risk_score_13), sum(df$investment_risk_score_14),
                                         sum(df$investment_risk_score_15), sum(df$investment_risk_score_16),
                                         sum(df$investment_risk_score_17)))

rownames(portfolio_oljefondet) <- c("Total risk score", "Deterrence risk", "Transparency risk", 
                         "Oversight risk", "Opportunity risk", "Interaction", "Expectation",
                         "Leverage", "Dissuasion", "Enforcement", "Processes", "Interests",
                         "Free press", "Civil society", "CPI", "CoC", "GDI")

portfolio_oljefondet <- rownames_to_column(portfolio_oljefondet)

portfolio_oljefondet$scope <- "GPFG"

colnames(portfolio_oljefondet) <- c("Risk_score", "value", "scope")

portfolio_oljefondet$Risk_score <- factor(portfolio_oljefondet$Risk_score)

#############################
# Data about New Zealand fund
############################

# Set working directory
setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Datasets on pension funds")

# Read in file
zealand <- read_xlsx("nzsf.xlsx", range = "A9:D6649")

# Remove NAs
zealand <- zealand %>%
  na.omit()

# Convert variables to factor
zealand$Country <- factor(zealand$Country)
zealand$GICSIndustryDescr <- factor(zealand$GICSIndustryDescr)

# Syncronise country name
zealand <- zealand %>% 
  mutate(Country = str_replace(Country, "Russia Federation", "Russian Federation"))

# Join with TRACE matrix
zealand <- left_join(zealand, trace_matrix, by = "Country")

fund_value_zealand <- sum(zealand$`Value in New Zealand Dollars`)

# Domains
zealand$`Overall Opportunity risk` <- as.numeric(zealand$`Overall Opportunity risk`)
zealand$`Overall Deterrence risk` <- as.numeric(zealand$`Overall Deterrence risk`)
zealand$`Overall Transparency risk` <- as.numeric(zealand$`Overall Transparency risk`)
zealand$`Overall Oversight risk` <- as.numeric(zealand$`Overall Oversight risk`)
zealand$`Value in New Zealand Dollars` <- as.numeric(zealand$`Value in New Zealand Dollars`)
zealand$`Total Risk score` <- as.numeric(zealand$`Total Risk score`)

# Sub-domains
zealand$Interaction <- as.numeric(zealand$Interaction)
zealand$Expectation <- as.numeric(zealand$Expectation)
zealand$Leverage <- as.numeric(zealand$Leverage)
zealand$Dissuasion <- as.numeric(zealand$Dissuasion)
zealand$Enforcement <- as.numeric(zealand$Enforcement)
zealand$Processes <- as.numeric(zealand$Processes)
zealand$Interests <- as.numeric(zealand$Interests)
zealand$`Free press` <- as.numeric(zealand$`Free press`)
zealand$`Civil society` <- as.numeric(zealand$`Civil society`)

#CPI
zealand <- left_join(zealand, df_cpi, by = "Country")
zealand$CPI_score <- as.numeric(zealand$CPI_score)

#CoC
zealand <- left_join(zealand, df_coc, by = "Country")
zealand$CoC_score <- as.numeric(zealand$CoC_score)

#GDI
zealand <- left_join(zealand, df_gdi, by = "Country")
zealand$GDI_score <- as.numeric(zealand$GDI_score)

zealand <- zealand %>%
  mutate(investment_percent = `Value in New Zealand Dollars` / fund_value_zealand,
         investment_risk_score_1 = investment_percent * `Total Risk score`,
         investment_risk_score_2 = investment_percent * `Overall Deterrence risk`,
         investment_risk_score_3 = investment_percent * `Overall Transparency risk`,
         investment_risk_score_4 = investment_percent * `Overall Oversight risk`,
         investment_risk_score_5 = investment_percent * `Overall Opportunity risk`,
         investment_risk_score_6 = investment_percent * Interaction,
         investment_risk_score_7 = investment_percent * Expectation,
         investment_risk_score_8 = investment_percent * Leverage,
         investment_risk_score_9 = investment_percent * Dissuasion,
         investment_risk_score_10 = investment_percent * Enforcement,
         investment_risk_score_11= investment_percent * Processes,
         investment_risk_score_12 = investment_percent * Interests,
         investment_risk_score_13 = investment_percent * `Free press`,
         investment_risk_score_14 = investment_percent * `Civil society`,
         investment_risk_score_15 = investment_percent * CPI_score,
         investment_risk_score_16 = investment_percent * CoC_score,
         investment_risk_score_17 = investment_percent * GDI_score)

# Add risk categories
zealand <- zealand %>%
  mutate(risk_cat = case_when(`Total Risk score` < 22 ~ "Very Low",
                              `Total Risk score` < 38 ~ "Low",
                              `Total Risk score` < 56 ~ "Moderate",
                              `Total Risk score` < 74 ~ "High",
                              `Total Risk score` > 74 ~ "Very high"))

portfolio_zealand <- data.frame(rbind(sum(zealand$investment_risk_score_1), 
                                      sum(zealand$investment_risk_score_2),
                                      sum(zealand$investment_risk_score_3), 
                                      sum(zealand$investment_risk_score_4),
                                      sum(zealand$investment_risk_score_5), 
                                      sum(zealand$investment_risk_score_6, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_7, na.rm = TRUE), 
                                      sum(zealand$investment_risk_score_8, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_9, na.rm = TRUE), 
                                      sum(zealand$investment_risk_score_10, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_11, na.rm = TRUE), 
                                      sum(zealand$investment_risk_score_12, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_13, na.rm = TRUE), 
                                      sum(zealand$investment_risk_score_14, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_15, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_16, na.rm = TRUE),
                                      sum(zealand$investment_risk_score_17, na.rm = TRUE)))

rownames(portfolio_zealand) <- c("Total risk score", "Deterrence risk", "Transparency risk", 
                                 "Oversight risk", "Opportunity risk", "Interaction", "Expectation",
                                 "Leverage", "Dissuasion", "Enforcement", "Processes", "Interests",
                                 "Free press", "Civil society", "CPI", "CoC", "GDI")

portfolio_zealand <- rownames_to_column(portfolio_zealand)

portfolio_zealand$scope <- "NZ Super"

colnames(portfolio_zealand) <- c("Risk_score", "value", "scope")

portfolio_zealand$Risk_score <- factor(portfolio_zealand$Risk_score)

#######################
# Data for APFC fund
#######################

APFC <- pdf_text("APCF_equity_holdings.pdf") %>%
  read_lines()

# Remove empty elements from the vector
APFC <- APFC[APFC != ""]

# Remove unnecessary rows
index <- grep(pattern = "06/30/2020", APFC)
APFC = APFC[-index]

index <- grep(pattern = "Description", APFC)
APFC = APFC[-index]

index <- grep(pattern = "Unrealized", APFC)
APFC = APFC[-index]

index <- grep(pattern = "(USD)", APFC)
APFC = APFC[-index]

index <- c(1,2)
APFC = APFC[-index]

# remove excess white spaces
APFC <- gsub("\\s+", " ", APFC)

# Extract and save firm names
Name <- trimws(sub("[[:digit:]].*", "\\1", APFC))
Name <- Name %>% as_tibble(.,rownames= c("ID"))
colnames(Name) <- c("ID", "Name")

# Remove everything before first digit
APFC <- sub("^\\D+(\\d)", "\\1", APFC)

# Extract and save share amount
shares <- sub("\\s+.*", "\\1", APFC)
shares <- shares %>% as_tibble(.,rownames="ID")
colnames(shares) <- c("ID", "Shares")

# Remove share amount
APFC <- sub(".+? ", "", APFC)

# Extract and save book value
book_value <- sub("\\s+.*", "\\1", APFC)
book_value <- book_value %>% as_tibble(.,rownames="ID")
colnames(book_value) <- c("ID", "Book_value")

# Remove book value
APFC <- sub(".+? ", "", APFC)

# Extract and save market value
market_value <- sub("\\s+.*", "\\1", APFC)
market_value <- market_value %>% as_tibble(.,rownames="ID")
colnames(market_value) <- c("ID", "Market_value")

# Remove market value
APFC <- sub(".+? ", "", APFC)

# Extract and save performance
performance <- sub("\\s+.*", "\\1", APFC)
performance <- performance %>% as_tibble(.,rownames="ID")
colnames(performance) <- c("ID", "Performance")

# Remove performance
APFC <- sub(".+? ", "", APFC)

# Extract and save country
country <- sub("\\s+.*", "\\1", APFC)
country <- country %>% as_tibble(.,rownames="ID")
colnames(country) <- c("ID", "Country")

# Remove country
APFC <- sub(".+? ", "", APFC)

# convert sector to tibble
sector <- APFC %>% as_tibble(.,rownames="ID")
colnames(sector) <- c("ID", "Sector")

APFC <- cbind(Name, country, shares, book_value, market_value, 
              performance, sector) %>%
  select(Name, Country, Shares, Book_value, Market_value, Performance, Sector)

APFC$Country <- str_to_title(APFC$Country)

# Synconise country names
APFC <- APFC %>% 
  mutate(Country = str_replace(Country, "Hong", "Hong Kong"), 
         Country = str_replace(Country, "New", "New Zealand"),
         Country = str_replace(Country, "Russia", "Russian Federation"),
         Country = str_replace(Country, "South", "South Korea"),
         Country = str_replace(Country, "Cambodia", "Hong Kong"),
         Country = str_replace(Country, "Czech", "Czech Republic"))

APFC <- read_xlsx("APCF_to_clean.xlsx")

APFC <- APFC %>%
  mutate(Country = str_replace(Country, "australia", "Australia"),
         Country = str_replace(Country, "Puerto", "Puerto Rico"))

# Join with TRACE matrix
APFC <- left_join(APFC, trace_matrix, by = "Country")

APFC$Market_value <- as.numeric(gsub(",", "", APFC$Market_value))

fund_value_APFC <- sum(APFC$Market_value, na.rm = TRUE)

# Domains
APFC$`Overall Opportunity risk` <- as.numeric(APFC$`Overall Opportunity risk`)
APFC$`Overall Deterrence risk` <- as.numeric(APFC$`Overall Deterrence risk`)
APFC$`Overall Transparency risk` <- as.numeric(APFC$`Overall Transparency risk`)
APFC$`Overall Oversight risk` <- as.numeric(APFC$`Overall Oversight risk`)
APFC$`Total Risk score` <- as.numeric(APFC$`Total Risk score`)

# Sub-domains
APFC$Interaction <- as.numeric(APFC$Interaction)
APFC$Expectation <- as.numeric(APFC$Expectation)
APFC$Leverage <- as.numeric(APFC$Leverage)
APFC$Dissuasion <- as.numeric(APFC$Dissuasion)
APFC$Enforcement <- as.numeric(APFC$Enforcement)
APFC$Processes <- as.numeric(APFC$Processes)
APFC$Interests <- as.numeric(APFC$Interests)
APFC$`Free press` <- as.numeric(APFC$`Free press`)
APFC$`Civil society` <- as.numeric(APFC$`Civil society`)

# CPI
APFC <- left_join(APFC, df_cpi, by = "Country")
APFC$CPI_score <- as.numeric(APFC$CPI_score)

# CoC
APFC <- left_join(APFC, df_coc, by = "Country")
APFC$CoC_score <- as.numeric(APFC$CoC_score)

# GDI
APFC <- left_join(APFC, df_gdi, by = "Country")
APFC$GDI_score <- as.numeric(APFC$GDI_score)

APFC <- APFC %>%
  mutate(investment_percent = Market_value / fund_value_APFC,
         investment_risk_score_1 = investment_percent * `Total Risk score`,
         investment_risk_score_2 = investment_percent * `Overall Deterrence risk`,
         investment_risk_score_3 = investment_percent * `Overall Transparency risk`,
         investment_risk_score_4 = investment_percent * `Overall Oversight risk`,
         investment_risk_score_5 = investment_percent * `Overall Opportunity risk`,
         investment_risk_score_6 = investment_percent * Interaction,
         investment_risk_score_7 = investment_percent * Expectation,
         investment_risk_score_8 = investment_percent * Leverage,
         investment_risk_score_9 = investment_percent * Dissuasion,
         investment_risk_score_10 = investment_percent * Enforcement,
         investment_risk_score_11= investment_percent * Processes,
         investment_risk_score_12 = investment_percent * Interests,
         investment_risk_score_13 = investment_percent * `Free press`,
         investment_risk_score_14 = investment_percent * `Civil society`,
         investment_risk_score_15 = investment_percent * CPI_score,
         investment_risk_score_16 = investment_percent * CoC_score,
         investment_risk_score_17 = investment_percent * GDI_score)

# Add risk categories
APFC <- APFC %>%
  mutate(risk_cat = case_when(`Total Risk score` < 22 ~ "Very Low",
                              `Total Risk score` < 38 ~ "Low",
                              `Total Risk score` < 56 ~ "Moderate",
                              `Total Risk score` < 74 ~ "High",
                              `Total Risk score` > 74 ~ "Very high"))

portfolio_APFC <- data.frame(rbind(sum(APFC$investment_risk_score_1, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_2, na.rm = TRUE ),
                                   sum(APFC$investment_risk_score_3, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_4, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_5, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_6, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_7, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_8, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_9, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_10, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_11, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_12, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_13, na.rm = TRUE), 
                                   sum(APFC$investment_risk_score_14, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_15, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_16, na.rm = TRUE),
                                   sum(APFC$investment_risk_score_17, na.rm = TRUE)))

rownames(portfolio_APFC) <- c("Total risk score", "Deterrence risk", "Transparency risk", 
                              "Oversight risk", "Opportunity risk", "Interaction", "Expectation",
                              "Leverage", "Dissuasion", "Enforcement", "Processes", "Interests",
                              "Free press", "Civil society", "CPI", "CoC", "GDI")

portfolio_APFC <- rownames_to_column(portfolio_APFC)

portfolio_APFC$scope <- "APFC"

colnames(portfolio_APFC) <- c("Risk_score", "value", "scope")

portfolio_APFC$Risk_score <- factor(portfolio_APFC$Risk_score)

##############
# Data on ABP
##############

abp <- read_xlsx("ABP.xlsx")

# Join with TRACE matrix
abp <- left_join(abp, trace_matrix, by = "Country")

# Remove NAs
abp <- abp %>%
  na.omit()

# Calculate fund value
fund_value_abp <- sum(abp$`Invested Market Value`, na.rm = TRUE)

# Domains
abp$`Overall Opportunity risk` <- as.numeric(abp$`Overall Opportunity risk`)
abp$`Overall Deterrence risk` <- as.numeric(abp$`Overall Deterrence risk`)
abp$`Overall Transparency risk` <- as.numeric(abp$`Overall Transparency risk`)
abp$`Overall Oversight risk` <- as.numeric(abp$`Overall Oversight risk`)
abp$`Total Risk score` <- as.numeric(abp$`Total Risk score`)

# Sub-domains
abp$Interaction <- as.numeric(abp$Interaction)
abp$Expectation <- as.numeric(abp$Expectation)
abp$Leverage <- as.numeric(abp$Leverage)
abp$Dissuasion <- as.numeric(abp$Dissuasion)
abp$Enforcement <- as.numeric(abp$Enforcement)
abp$Processes <- as.numeric(abp$Processes)
abp$Interests <- as.numeric(abp$Interests)
abp$`Free press` <- as.numeric(abp$`Free press`)
abp$`Civil society` <- as.numeric(abp$`Civil society`)

#CPI
abp <- left_join(abp, df_cpi, by = "Country")
abp$CPI_score <- as.numeric(abp$CPI_score)

#CoC
abp <- left_join(abp, df_coc, by = "Country")
abp$CoC_score <- as.numeric(abp$CoC_score)

#GDI
abp <- left_join(abp, df_gdi, by = "Country")
abp$GDI_score <- as.numeric(abp$GDI_score)

abp <- abp %>%
  mutate(investment_percent = `Invested Market Value` / fund_value_abp,
         investment_risk_score_1 = investment_percent * `Total Risk score`,
         investment_risk_score_2 = investment_percent * `Overall Deterrence risk`,
         investment_risk_score_3 = investment_percent * `Overall Transparency risk`,
         investment_risk_score_4 = investment_percent * `Overall Oversight risk`,
         investment_risk_score_5 = investment_percent * `Overall Opportunity risk`,
         investment_risk_score_6 = investment_percent * Interaction,
         investment_risk_score_7 = investment_percent * Expectation,
         investment_risk_score_8 = investment_percent * Leverage,
         investment_risk_score_9 = investment_percent * Dissuasion,
         investment_risk_score_10 = investment_percent * Enforcement,
         investment_risk_score_11= investment_percent * Processes,
         investment_risk_score_12 = investment_percent * Interests,
         investment_risk_score_13 = investment_percent * `Free press`,
         investment_risk_score_14 = investment_percent * `Civil society`,
         investment_risk_score_15 = investment_percent * CPI_score,
         investment_risk_score_16 = investment_percent * CoC_score,
         investment_risk_score_17 = investment_percent * GDI_score)

abp <- abp %>%
  mutate(risk_cat = case_when(`Total Risk score` < 22 ~ "Very Low",
                              `Total Risk score` < 38 ~ "Low",
                              `Total Risk score` < 56 ~ "Moderate",
                              `Total Risk score` < 74 ~ "High",
                              `Total Risk score` > 74 ~ "Very high"))

portfolio_abp <- data.frame(rbind(sum(abp$investment_risk_score_1, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_2, na.rm = TRUE ),
                                   sum(abp$investment_risk_score_3, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_4, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_5, na.rm = TRUE),
                                   sum(abp$investment_risk_score_6, na.rm = TRUE),
                                   sum(abp$investment_risk_score_7, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_8, na.rm = TRUE),
                                   sum(abp$investment_risk_score_9, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_10, na.rm = TRUE),
                                   sum(abp$investment_risk_score_11, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_12, na.rm = TRUE),
                                   sum(abp$investment_risk_score_13, na.rm = TRUE), 
                                   sum(abp$investment_risk_score_14, na.rm = TRUE),
                                   sum(abp$investment_risk_score_15, na.rm = TRUE),
                                   sum(abp$investment_risk_score_16, na.rm = TRUE),
                                   sum(abp$investment_risk_score_17, na.rm = TRUE)))

rownames(portfolio_abp) <- c("Total risk score", "Deterrence risk", "Transparency risk", 
                              "Oversight risk", "Opportunity risk", "Interaction", "Expectation",
                              "Leverage", "Dissuasion", "Enforcement", "Processes", "Interests",
                              "Free press", "Civil society", "CPI", "CoC", "GDI")

portfolio_abp <- rownames_to_column(portfolio_abp)

portfolio_abp$scope <- "ABP"

colnames(portfolio_abp) <- c("Risk_score", "value", "scope")

portfolio_abp$Risk_score <- factor(portfolio_abp$Risk_score)

# Remove temporary data
rm(list = c("trace_matrix_fix", "df_fix", "scores", "country_split", "oljefondet", "countries",
            "criterion", "date", "decision", "firms", "year", "book_value", "country", "exclusion",
            "iso_codes", "market_value", "Name", "performance", "sector", "shares", "index"))

####################################################################################################
# End of data cleaning and preparation
###################################################################################################

##############
# VISUALIZATION
##############

# Merge funds together
scores <- rbind(portfolio_oljefondet, portfolio_zealand, 
                portfolio_APFC, portfolio_abp)

# Fix order of faceting variables
scores$Risk_score2 = factor(scores$Risk_score, 
                            levels=c("Dissuasion", "Enforcement", "Leverage", 
                                     "Expectation", "Civil society", "Free press",
                                     "Processes", "Interaction", "Interests", 
                                     "CPI", "CoC", "GDI"))

# Add domain names
scores <- scores %>%
  mutate(domain = case_when(Risk_score2 %in% c("Interaction", "Expectation", "Leverage") ~ "Opportunity",
                            Risk_score2 %in% c("Processes", "Interests") ~ "Transparency",
                            Risk_score2 %in% c("Dissuasion", "Enforcement") ~ "Deterrence",
                            Risk_score2 %in% c("Free press", "Civil society") ~ "Oversight"))

scores$domain <- factor(scores$domain)

# Reorder groups
scores$scope <- factor(scores$scope, levels = c("GPFG", "ABP", "APFC", "NZ Super"))
scores$Risk_score2 <- factor(scores$Risk_score2)

#########
# PLOT 1 
########

scores$Risk_score <- factor(scores$Risk_score, levels = c("Total risk score", "Deterrence risk",
                                                          "Opportunity risk", "Oversight risk",
                                                          "Transparency risk"))

# Calculate average risk scores across all funds
mean_df <- scores %>%
  filter(Risk_score %in% c("Total risk score", "Deterrence risk", "Transparency risk",
                           "Oversight risk", "Opportunity risk")) %>%
  select(scope, Risk_score, value) %>%
  group_by(Risk_score) %>%
  summarise(value = mean(value))
  

# Plot weighted portfolio scores per domains (bar graph)
scores %>%
  filter(Risk_score %in% c("Total risk score", "Deterrence risk", "Transparency risk",
                           "Oversight risk", "Opportunity risk")) %>%
  select(Risk_score, value, scope) %>%
  ggplot(aes(x= as.numeric(value), y= reorder(Risk_score, -value), fill = scope))  +
  geom_col(colour = "black", width= 0.6, 
           position = position_dodge(width= 0.9)) +  
  geom_vline(data = mean_df, aes(xintercept = value), colour = "red",
             linetype = "dashed", size = 1) +
  theme_bw() + 
  geom_text(aes(label = round(value, 1)), 
            vjust=-0.5, color="black", size=3.6, 
            position = position_dodge(0.9)) + 
       labs(x = "",
       y = "", 
       fill = NULL) + 
  theme(legend.position = "bottom",
        axis.text.y = element_text(size = 12),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        strip.text.x = element_text(size = 14),
        legend.text = element_text(size = 16)) + 
  facet_nested(.~ Risk_score, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("GPFG" = "lightblue4", "ABP" = "lightblue3", 
                               "APFC" = "lightblue2", "NZ Super" = "lightblue1")) +  
  scale_x_continuous(expand = c(0, 0), limits = c(0, 40)) + 
  coord_flip()
  

##################
# Robustness check
#################

scores2 <- scores %>%
  filter(Risk_score2 %in% c("CPI", "CoC") | Risk_score == "Total risk score") %>%
  mutate(Risk_score2 = str_replace(Risk_score2, "CPI", "TI-CPI"),
         Risk_score2 = str_replace(Risk_score2, "CoC", "WB-CoC"),
         Risk_score2 = replace_na(Risk_score2, "TRACE Matrix")) %>%
  select(scope, Risk_score2, value)

scores2$Risk_score2 <- factor(scores2$Risk_score, 
                             levels = c("TRACE Matrix", "TI-CPI", "WB-CoC"))

# Calculate average risk scores across all funds
mean_df2 <- scores2 %>%
  group_by(Risk_score2) %>%
  summarise(value = mean(value))

scores2 %>%
  ggplot(aes(x= as.numeric(value), y= reorder(Risk_score2, -value), fill = scope))  +
  geom_col(colour = "black", width= 0.6, 
           position = position_dodge(width= 0.9)) +  
  geom_vline(data = mean_df2, aes(xintercept = value), colour = "red",
             linetype = "dashed", size = 1) + 
  theme_bw() + 
  geom_text(aes(label = round(value, 1)), 
            vjust=-0.5, color="black", size=3.2, 
            position = position_dodge(0.9)) + 
  labs(x = "",
       y = "", 
       fill = NULL) + 
  theme(legend.position = "bottom",
        axis.text.y = element_text(size = 12),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        strip.text.x = element_text(size = 14),
        legend.text = element_text(size = 16)) + 
  facet_grid(.~ Risk_score2, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("GPFG" = "lightblue4", "ABP" = "lightblue3", 
                               "APFC" = "lightblue2", "NZ Super" = "lightblue1")) +  
  scale_x_continuous(expand = c(0, 0), limits = c(0, 40)) + 
  coord_flip()

########
# PLOT 2
########

###############################################
# Plot for country risk scores across portfolio
###############################################

gpfg <- df %>%
  select(Country, `Total Risk score`, `Market Value(USD)`) %>%
  rename(value = `Market Value(USD)`) %>%
  mutate(scope = "GPFG")

nzsuper <- zealand %>%
  select(Country, `Total Risk score`, `Value in New Zealand Dollars`) %>%
  rename(value = `Value in New Zealand Dollars`) %>%
  mutate(scope = "NZ Super")

apfc <- APFC %>%
  select(Country, Market_value, `Total Risk score`) %>%
  rename(value = Market_value) %>%
  mutate(scope = "APFC")

ABP <- abp %>%
  select(Country, `Invested Market Value`, `Total Risk score`) %>%
  rename(value = `Invested Market Value`) %>%
  mutate(scope = "ABP")

portfolios <- rbind(gpfg, nzsuper, apfc, ABP)

########
# PLOT 4
########

##################################################################
# Share of count of firms in the different risk categories per portfolio
#################################################################

#####
# GPFG
#####

df$risk_cat <- factor(df$risk_cat)

# Calculate category shares
oljefondet_category_shares <- df %>%
  group_by(risk_cat) %>%
  select(Name, `Market Value(USD)`) %>%
  summarise(n = n(),
            share = round(n / nrow(df), 3),
            value = sum(`Market Value(USD)`),
            percent = round(value / fund_value_oljefondet, 5))

# Add scope variable
oljefondet_category_shares$scope <- rep("GPFG", 4)

###############
# NZ Super Fund
###############
zealand$risk_cat <- factor(zealand$risk_cat)

# Calculate category shares
zealand_category_shares <- zealand %>%
  group_by(risk_cat) %>%
  select(`Security Name`, `Value in New Zealand Dollars`) %>%
  summarise(n = n(),
            share = round(n / nrow(zealand), 3),
            value = sum(`Value in New Zealand Dollars`),
            percent = round(value / fund_value_zealand, 5))

# Add scope variable
zealand_category_shares$scope <- rep("NZ Super", 4)

######
# APFC
######
APFC$risk_cat <- factor(APCF$risk_cat)

APFC_category_shares <- APFC %>%
  group_by(risk_cat) %>%
  select(Name, Market_value) %>%
  na.omit() %>%
  summarise(n = n(),
            share = round(n / nrow(APFC), 3),
            value = sum(Market_value),
            percent = round(value / fund_value_APFC, 5))

# Add scope variable
APFC_category_shares$scope <- rep("APFC", 4)

#########
# ABP
########

abp$risk_cat <- factor(abp$risk_cat)

abp_category_shares <- abp %>%
  group_by(risk_cat) %>%
  select(Company, `Invested Market Value`) %>%
  na.omit() %>%
  summarise(n = n(),
            share = round(n / nrow(abp), 3),
            value = sum(`Invested Market Value`),
            percent = round(value / fund_value_abp, 5))

abp_category_shares$scope <- rep("ABP", 4)

shares <- rbind(oljefondet_category_shares, zealand_category_shares, 
                APFC_category_shares, abp_category_shares)

# Reorder factor levels
shares$risk_cat <- factor(shares$risk_cat, levels = c("High", "Moderate", "Low", "Very Low"))
shares$scope <- factor(shares$scope, levels = c("GPFG", "ABP", "APFC", "NZ Super"))

# Share of firm counts per portfolio
shares %>%
  ggplot(aes(x= reorder(scope, share), y= share, fill = risk_cat)) + 
  geom_col(position = "fill", width = 0.5) + 
  geom_text(aes(label = paste(share*100, "%")), 
            position=position_stack(vjust=0.5), size = 4, 
            vjust = ifelse(shares$percent < 1.5, -0.15, 1)) + 
  theme_bw() + 
  scale_y_continuous(labels = scales::percent) + 
  labs(x= "", y="", fill = NULL) +
  theme(legend.position = "bottom",
        axis.text.y = element_text(size = 14),
        legend.text = element_text(size = 14),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        strip.text.x = element_text(size = 16)) + 
  facet_grid(~scope, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very Low" = "lightblue1"))


###########
# PLOT 5
###########

# Fixing percent decimal values
shares <- shares %>%
  mutate(percent_new = ifelse(percent < 0.00156, round(percent, 5)*100, round(percent, 4)*100))

# Share of investment value per portfolio
shares %>%
  ggplot(aes(x= reorder(scope, percent), y= percent, fill = risk_cat)) + 
  geom_col(position = "fill", width = 0.5) + 
  geom_text(aes(label = paste(percent_new, "%")), 
            position=position_stack(vjust=0.5), size = 3,
            vjust = ifelse(shares$percent_new == 0.049, -1, 0)) +
  theme_bw() + 
  scale_y_continuous(labels = scales::percent) + 
  labs(x= "", y="", fill = NULL) +
  facet_grid(~scope, scales = "free") + 
  theme(legend.position = "bottom",
        axis.text.y = element_text(size = 12),
        legend.text = element_text(size = 12),
        axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        strip.text.x = element_text(size = 14)) + 
  facet_grid(~scope, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very Low" = "lightblue1"))

# Plot 6
shares_gathered <- shares %>%  
  filter(risk_cat == "High") %>%
  select(scope, share, percent) %>%
  gather(., key = variable, value = value, share:percent)

shares_gathered$variable <- factor(shares_gathered$variable,
                                   levels = c("share", "percent"))

shares_gathered %>%
  ggplot(aes(x= reorder(scope, -value), y= value, fill = variable)) +
  geom_col(position = position_dodge(width = 0.9), width = 0.6) + 
  scale_y_continuous(labels = scales::percent_format(accuracy = 0.1),
                     limits = c(0, 0.015),
                     expand = c(0, 0)) +
  theme_bw() + 
  labs(x= "", y= "") + 
  geom_text(aes(label = paste(round(value*100, 2), "%")), vjust = -0.3, 
            position = position_dodge(width = 0.9), size = 3.5) +
  theme(axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        axis.text.y = element_text(size = 12),
        legend.position = "bottom",
        legend.text = element_text(size = 12),
        strip.text.x = element_text(size = 12)) + 
  scale_fill_manual(name = NULL, 
                    values = c("share" = "lightblue4", 
                               "percent" = "lightblue2"),
                    labels = c("Firm count share", 
                               "Portfolio value share")) + 
  facet_grid(~scope, scales = "free")

######
# Investment share in high risk countries in USD
######
oljefondet_plot7 <- df %>%
  filter(risk_cat == "High") %>%
  mutate(total = sum(`Market Value(USD)`)) %>%
  group_by(Country, total) %>%
  summarise(value = sum(`Market Value(USD)`)) %>%
  mutate(freq = value / total,
         scope = "GPFG")

abp_plot7 <- abp %>%
  filter(risk_cat == "High") %>%
  mutate(total = sum(`Invested Market Value`)* 1.1220 * 1000000) %>%
  group_by(Country, total) %>%
  summarise(value = sum(`Invested Market Value`) * 1.1220 * 1000000) %>%
  mutate(freq = value / total,
         scope = "ABP")

apfc_plot7 <- APFC %>%
  filter(risk_cat == "High") %>%
  mutate(total = sum(Market_value)) %>%
  group_by(Country, total) %>%
  summarise(value = sum(Market_value)) %>%
  mutate(freq = value / total,
         scope = "APFC")

nzsuper_plot7 <- zealand %>%
  filter(risk_cat == "High") %>%
  mutate(total = sum(`Value in New Zealand Dollars`)* 0.6722) %>%
  group_by(Country, total) %>%
  summarise(value = sum(`Value in New Zealand Dollars`) * 0.6722) %>%
  mutate(freq = value / total,
         scope = "NZ Super")

plot7 <- rbind(oljefondet_plot7, apfc_plot7, nzsuper_plot7, abp_plot7)

plot7$Country <- factor(plot7$Country)
plot7$scope <- factor(plot7$scope, levels = c("GPFG", "ABP", "APFC", "NZ Super"))

plot7 %>%
  ggplot(aes(x= scope , y= freq, fill = Country)) +
  geom_col(position = "fill") +
  facet_grid(~scope, scales = "free") +
  theme_bw() + 
  labs(fill = "Country (risk score)", x= "", y= "") + 
  scale_y_continuous(labels = scales::percent_format(accuracy = 1)) +
  geom_text(aes(label = ifelse(freq > 0.016507870 | freq < 0.001451783, 
                               paste(round(freq, 4)*100, "%"), 
                               "")), 
                position=position_stack(vjust=0.5), size = 3.5) + 
  scale_fill_brewer(palette = "Blues", 
                    labels = (c("Bangladesh" = "Bangladesh (66)",
                                                   "Gabon" = "Gabon (66)",
                                                   "Egypt" = "Egypt (64)",
                                                   "Liberia" = "Liberia (63)",
                                                   "Nigeria" = "Nigeria (61)",
                                                   "Pakistan" = "Pakistan (61)",
                                                   "Vietnam" = "Vietnam (58)")),
                    limits = c("Bangladesh", "Gabon", "Egypt", "Liberia", "Nigeria",
                               "Pakistan", "Vietnam")) +
  theme(axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        axis.text.y = element_text(size = 12),
        legend.text = element_text(size = 10),
        strip.text.x = element_text(size = 12))
  

totals <- plot7 %>%
  group_by(scope) %>%
  summarize(value = sum(value))

plot7 %>%
  ggplot(aes(x= scope, y= value / 1000000, fill = scope)) +
  geom_col(width = 0.6) +
  facet_grid(~scope, scales = "free") +
  theme_bw() + 
  theme(axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
        axis.text.y = element_text(size = 12),
        legend.position = "none",
        strip.text.x = element_text(size = 12)) +
  labs(fill = NULL, x= "", y= "Million USD") + 
  scale_fill_manual(name = NULL, 
                    values = c("GPFG" = "lightblue4", "ABP" = "lightblue3", 
                               "APFC" = "lightblue2", "NZ Super" = "lightblue1")) + 
  ylim(c(0, 1500)) + 
  geom_text(data = totals, aes(label = round(value / 1000000, 2)), vjust = -0.3)

###################
#Exclusion graph
##################
levels(exclusion$Criterion) <- sub("Serious violations of individuals' rights in situations of war or conflict",
                                   "Violation of individuals rights in war",
                                   levels(exclusion$Criterion))



exclusion_df <- exclusion %>%
  filter(Criterion %in% c("Production of coal or coal-based energy",
                          "Production of nuclear weapons",
                          "Production of tobacco",
                          "Severe environmental damage",
                          "Unacceptable greenhouse gas emissions",
                          "Violation of human rights",
                          "Gross corruption",
                          "Violation of individuals rights in war"),
         Decision == "Exclusion") %>%
  mutate(Criterion = str_replace(Criterion, "Serious violations of individuals' rights in situations of war or conflict",
                                 "Serious violations of individuals' rights in war or conflict")) %>%
  group_by(Criterion) %>%
  summarise(n= n()) %>%
  arrange(desc(n))

legend_ord <- levels(with(exclusion_df, reorder(Criterion, -n)))

exclusion_df$Criterion <- factor(exclusion_df$Criterion)

exclusion_df <- exclusion_df %>%
  mutate(Criterion = fct_relevel(Criterion, 
                                 "Production of coal or coal-based energy",
                                 "Severe environmental damage",
                                 "Production of tobacco",
                                 "Production of nuclear weapons",
                                 "Violation of human rights",
                                 "Unacceptable greenhouse gas emissions",
                                 "Gross corruption",
                                 "Violation of individuals rights in war"))


exclusion_df %>%
  ggplot(aes(x= reorder(Criterion, -n), y= n, fill = Criterion)) +
  geom_col(position = "dodge", width = 0.7) + 
  theme_minimal() +
  geom_text(aes(label = n), size = 4, vjust = -0.2) + 
  theme(legend.position = "bottom", 
        axis.text.x = element_blank(),
        axis.ticks.x = element_blank(),
        legend.text = element_text(size = 10),
        axis.text.y = element_text(size = 10)) + 
  guides(fill = guide_legend(nrow = 2)) + 
  labs(fill = NULL, x= "", y= "") + 
  ylim(c(0, 80)) +
  scale_fill_brewer(breaks = legend_ord,
                    labels = c("Production of coal or coal-based energy" = "Coal production",
                               "Severe environmental damage" = "Environmental damage",
                               "Production of tobacco" = "Tobacco production",
                               "Production of nuclear weapons" = "Nuclear weapon production",
                               "Violation of human rights" = "Human rights violation",
                               "Unacceptable greenhouse gas emissions" = "Greenhouse emissions",
                               "Gross corruption" = "Gross corruption",
                               "Violation of individuals rights in war" = 
                                 "Violation of individuals rights in war"),
                    palette = "Blues",
                    direction = -1)

########################################
# Calculate operating country risk score
########################################

setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/FBRI/Subsidiary Files")

# Save path
directory = getwd()

# Extract firm names from file names
firm_names <- list.files(path = directory, pattern =".xlsx")

# Initiate an empty data frame
operating_country_scores <- data.frame(Name = character(0),
                                        subsidiary1 = character(0),
                                        risk_score1 = numeric(0),
                                        subsidiary2 = character(0),
                                        risk_score2 = numeric(0), 
                                        subsidiary3 = character(0),
                                        risk_score3 = numeric(0))

# Initiate for loop to calculate risk scores
for (firm in firm_names) {
  
  file <- read_xlsx(firm, sheet = 2)
  
  file <- file %>%
    filter(Country != "n.a.")
  
  file <- file %>%
    rename(Alpha_2 = Country) %>%
    select(Alpha_2)
  
  file <- left_join(file, trace_matrix, by = "Alpha_2")
  
  file$`Total Risk score` <- as.numeric(file$`Total Risk score`)
  
  file <- file %>%
    na.omit()
  
  output <- file %>%
    select(Country, `Total Risk score`) %>%
    group_by(Country, `Total Risk score`) %>%
    summarise(subsidiary_count= n()) %>% 
    arrange(desc(`Total Risk score`)) %>%
    head(3)
  
  if (nrow(output) == 1) {
    Name <- sub("\\..*", "", firm)
    subsidiary1 = output[[1, 1]]
    risk_score1 = output[[1, 2]]
    subsidiary2 = NA
    risk_score2 = NA 
    subsidiary3 = NA
    risk_score3 = NA 
    
    } else if (nrow(output) == 2) {
    
    Name <- sub("\\..*", "", firm)
    subsidiary1 = output[[1, 1]]
    risk_score1 = output[[1, 2]]
    subsidiary2 = output[[2, 1]]
    risk_score2 = output[[2, 2]] 
    subsidiary3 = NA
    risk_score3 = NA 
    
    } else {
    
    Name <- sub("\\..*", "", firm)
    subsidiary1 = output[[1, 1]]
    risk_score1 = output[[1, 2]]
    subsidiary2 = output[[2, 1]]
    risk_score2 = output[[2, 2]] 
    subsidiary3 = output[[3, 1]]
    risk_score3 = output[[3, 2]] 
  }
  
  operating_country_scores <- rbind(operating_country_scores, 
                                     data.frame(Name, subsidiary1, 
                                                risk_score1, subsidiary2, 
                                                risk_score2, subsidiary3, 
                                                risk_score3))
}

# Calculate the average of 3 riskiest operating country risk score
operating_country_scores <- operating_country_scores %>%
  rowwise() %>%
  mutate(operating_country_score = round(mean(c(risk_score1, risk_score2, risk_score3), 
                                                 na.rm = TRUE), 2))

####################################
# Calculate subsidiary sector scores
####################################

setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Subsidiaries & SIC")

weights <- read.xlsx("Sector_Score.xlsx", sheetIndex = 1)

weights <- weights %>%
  rename(OIC = Own.Industry.Code) %>%
  select(OIC, Final.Scores, Industry)

setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/FBRI/Subsidiary Files")

# Save path
directory = getwd()

# Extract firm names from file names
firm_names <- list.files(path = directory, pattern =".xlsx")

# Initiate an empty data frame
operating_subsidiary_scores <- data.frame(Name = character(0),
                                          SIC_code1 = numeric(0),
                                          SIC1_description = character(0),
                                          Risk_score1= numeric(0),
                                          SIC_code1 = numeric(0),
                                          SIC2_description = character(0), 
                                          Risk_score2= numeric(0),
                                          SIC_code3 = numeric(0),
                                          SIC3_description = character(0),
                                          Risk_score3 = numeric(0))

for (firm in firm_names) {
  
  setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/FBRI/Subsidiary Files")
  
  file <- read.xlsx(firm, sheetIndex = 2)
  
  setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Subsidiaries & SIC")
  
  mapping <- read_xlsx("SIC_Mapping_Complete.xlsx", sheet = 1)
  
  mapping <- mapping %>%
    select(`INDUSTRY GROUP...9`, OIC, `CATEGORY DESCRIPTION`) %>%
    na.omit() %>%
    rename(sic = `INDUSTRY GROUP...9`) %>%
    distinct()
  
  file <- file %>%
    rename(sic = US.SIC..Core.code) %>%
    select(sic) %>%
    na.omit()
  
  file <- left_join(file, mapping, by = "sic")
  
  file <- left_join(file, weights, by = "OIC")
  
  file$Final.Scores <- as.numeric(file$Final.Scores)

  output <- file %>%
    na.omit() %>%
    distinct(Final.Scores, .keep_all = TRUE) %>%
    arrange(desc(Final.Scores)) %>%
    head(3)

  if (nrow(output) == 1) {
    Name <- sub("\\..*", "", firm)
    SIC_code1 = output[[1, 1]]
    SIC1_description = output[[1, 5]]
    Risk_score1 = output[[1, 4]]
    SIC_code2 = NA
    SIC2_description = NA
    Risk_score2 = NA
    SIC_code3 = NA
    SIC3_description = NA
    Risk_score3 = NA
  
    } else if (nrow(output) == 2) {
  
    Name <- sub("\\..*", "", firm)
    SIC_code1 = output[[1, 1]]
    SIC1_description = output[[1, 5]]
    Risk_score1 = output[[1, 4]]
    SIC_code2 = output[[2, 1]]
    SIC2_description = output[[2, 5]]
    Risk_score2 = output[[2, 4]]
    SIC_code3 = NA
    SIC3_description = NA
    Risk_score3 = NA
  
    } else {
  
    Name <- sub("\\..*", "", firm)
    SIC_code1 = output[[1, 1]]
    SIC1_description = output[[1, 5]]
    Risk_score1 = output[[1, 4]]
    SIC_code2 = output[[2, 1]]
    SIC2_description = output[[2, 5]]
    Risk_score2 = output[[2, 4]]
    SIC_code3 = output[[3, 1]]
    SIC3_description = output[[3, 5]]
    Risk_score3 = output[[3, 4]]

    }
  
  operating_subsidiary_scores <- rbind(operating_subsidiary_scores, 
                                       data.frame(Name, 
                                                  SIC_code1, SIC1_description, Risk_score1,
                                                  SIC_code2, SIC2_description, Risk_score2, 
                                                  SIC_code3, SIC3_description, Risk_score3))
}

# Calculate the average of 3 riskiest SIC code risk score
operating_subsidiary_scores <- operating_subsidiary_scores %>%
  rowwise() %>%
  mutate(subsidiary_score = round(mean(c(Risk_score1, Risk_score2, Risk_score3), 
                                               na.rm = TRUE), 2))


#################
# Random sampling
#################

# Sectors with upper risk levels
upper <- c("Industrials", "Oil & Gas", "Utilities", "Basic Materials", 
           "Health Care", "Consumer Services", "Telecommunications")

# Sectors with moderate risk levels
lower <- c("Industrials", "Financials", "Consumer Goods", 
           "Consumer Services", "Technology", "Basic Materials")

#############################################
# LOWER industry risk + VERY LOW country risk
#############################################
set.seed(6)

low_verylow <- df %>%
  filter(Industry %in% lower,
         `Total Risk score` < 22) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# LOWER industry risk + LOW country risk
#############################################
set.seed(6)

low_low <-df %>%
  filter(Industry %in% lower,
         `Total Risk score` > 22 & `Total Risk score` < 38) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# LOWER industry risk + MODERATE country risk
#############################################
set.seed(6)

low_moderate <- df %>%
  filter(Industry %in% lower,
         `Total Risk score` > 38 & `Total Risk score` < 56) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# LOWER industry risk + HIGH country risk
#############################################
set.seed(6)

low_high <- df %>%
  filter(Industry %in% lower,
         `Total Risk score` > 56) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# UPPER industry risk + VERY LOW country risk
#############################################
set.seed(6)

high_verylow <- df %>%
  filter(Industry %in% upper,
         `Total Risk score` < 22) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# UPPER industry risk + LOW country risk
#############################################
set.seed(6)

high_low <- df %>%
  filter(Industry %in% upper,
         `Total Risk score` > 22 & `Total Risk score` < 38) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# UPPER industry risk + MODERATE country risk
#############################################
set.seed(6)

high_moderate <- df %>%
  filter(Industry %in% upper,
         `Total Risk score` > 38 & `Total Risk score` < 56) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

#############################################
# UPPER industry risk + HIGH country risk
#############################################
set.seed(6)

high_high <- df %>%
  filter(Industry %in% upper,
         `Total Risk score` > 56) %>%
  select(Country, Name, Industry, `Market Value(USD)`, 
         `Total Risk score`, risk_cat) %>%
  sample_n(., 40)

################
# Analysis
################
setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/FBRI")

df_80 <- read.xlsx("80_randomly_picked_firms.xlsx", sheetIndex = 2)

df_80 <- df_80 %>%
  filter(is.na(ID) == FALSE)

df_80$Country.risk <- factor(df_80$Country.risk, 
                             levels = c("High", "Moderate", "Low", "Very low"),
                             labels = c("High", "Moderate", "Low", "Very low"))

df_80$Industry.risk <- factor(df_80$Industry.risk)

df_80$Anti.corruption.effort.score <- as.numeric(df_80$Anti.corruption.effort.score, na.rm = TRUE)
df_80$Country.risk.score <- as.numeric(df_80$Country.risk.score)
df_80$FBRI <- as.numeric(df_80$FBRI)
df_80$Listing.Country.risk.score <- as.numeric(df_80$Listing.Country.risk.score)

# Normality plots per risk category
df_80 %>%
  ggplot(aes(x= FBRI)) +
  geom_histogram(aes(y = ..density..), fill = "lightblue2", colour = "black", bins = 40) +
  geom_density(alpha=.2, colour = "red", size = 1.2) + 
  theme_bw() + 
  facet_wrap(. ~ Country.risk, ncol = 2, scales = "free") + 
  theme(strip.text.x = element_text(size = 14)) + 
  labs(x= "", y= "")

# Normality plots FBRI
df_80 %>%
  ggplot(aes(x= FBRI)) +
  geom_histogram(aes(y = ..density..), fill = "lightblue2", colour = "black", bins = 40) +
  geom_density(alpha=.2, colour = "red", size = 1.2) + 
  theme_bw() + 
  labs(x= "FBRI scores", y= "Density")

# Scatterplots with regression lines
log.model.df <- data.frame(x = df_80$Listing.Country.risk.score,
                           y = exp(fitted(exp_model2)))

####
ggplot(df_80, aes(x = Listing.Country.risk.score, y = FBRI)) + geom_point() +
  stat_smooth(method = 'lm', aes(colour = 'Linear'), se = FALSE) +
  stat_smooth(method = 'lm', formula = y ~ poly(x,3), aes(colour = 'Polynomial'), se= FALSE) +
  labs(colour = NULL, x= "Listing country risk score", y= "FBRI") +
  geom_line(data = log.model.df, aes(x, y, color = "Exponential"), size = 1, linetype = 1) +
  theme_bw() +
  theme(legend.position = "bottom",
        legend.text = element_text(size = 12),
        axis.title.x = element_text(size = 12),
        axis.text.x = element_text(size = 12),
        axis.text.y = element_text(size = 12),
        axis.title.y = element_text(size = 12)) +
  scale_color_manual(limits = c("Linear", "Exponential", "Polynomial"), 
                     values = c("red", "grey", "blue")) +
  scale_x_continuous(limits = c(20, 70)) +
  guides(colour = guide_legend(override.aes = list(size=3)))

# Regressions
lm_model <- lm(FBRI ~ Listing.Country.risk.score, data = df_80)
exp_model <- lm(FBRI ~ exp(Listing.Country.risk.score), data = df_80)
exp_model2 <- lm(log(FBRI) ~ Listing.Country.risk.score, data = df_80)
poly_model <- lm(FBRI ~ poly(Listing.Country.risk.score,3), data = df_80)

# Regression output table
stargazer(lm_model, exp_model2, poly_model,
          out = "results.doc",
          align=TRUE,
          covariate.labels = c("Listing country risk score",
                               "Listing country risk score (poly I.)",
                               "Listing country risk score (poly II.)",
                               "Listing country risk score (poly III.)",
                               "Constant"),
          type = "text",
          no.space=TRUE,
          column.labels =  c("Linear", "Exponential", "Polynomial"))


leveneTest(df_80$FBRI, df_80$Country.risk, center = mean)

####################
# Country risk score
###################

# Summary statistics
df_80 %>% 
  group_by(Country.risk) %>%
  get_summary_stats(Country.risk.score, type = "mean_sd")

# ANOVA Test
df_80 %>% anova_test(Country.risk.score ~ Country.risk)

# Run Pairwise T-Test
df_80 %>%
  pairwise_t_test(Country.risk.score ~ Country.risk, 
                  p.adjust.method = "bonferroni")

my_comparisons <- list(c("High", "Moderate"), 
                       c("High", "Low"), 
                       c("High", "Very low"),
                       c("Moderate", "Low"),
                       c("Moderate", "Very low"),
                       c("Low", "Very low"))

##############
# FBRI scores
##############

# Summary statistics
df_80 %>% 
  group_by(Country.risk) %>%
  get_summary_stats(FBRI, type = "mean_sd")

# ANOVA Test
df_80 %>% anova_test(FBRI ~ Country.risk)

# FBRI Boxplot
df_80 %>%
  group_by(Country.risk) %>%
  ggplot(aes(x= Country.risk, y= FBRI)) +
  stat_boxplot(geom = "errorbar", position = position_dodge(0.9), width = 0.2) +
  geom_boxplot(outlier.colour = NA, position = position_dodge(0.9), 
               fill = "lightblue1", width = 0.3) + 
  stat_summary(fun = mean, colour="darkred", geom="point", size = 2.5,
               position = position_dodge(0.9), show.legend = FALSE) + 
  stat_summary(fun = mean, colour ="black", geom ="text",
               aes(label = round(..y.., digits = 2)), position = position_dodge(0.9),
               vjust = 1.5, size = 3.5) +
  theme_bw() + 
  labs(x= "", y="",
       caption = "Figure 1: FBRI scores") + 
  theme(legend.position = "bottom",
        legend.text = element_text(size = 13),
        axis.text.x = element_text(size = 13),
        axis.text.y = element_text(size = 13)) + 
  scale_y_continuous(limits = c(0, 100),
                     breaks = c(0, 25, 50, 75, 100)) + 
  stat_compare_means(comparisons = my_comparisons, 
                     method = "t.test",
                     label.y = c(65, 72, 79, 86, 93, 99),
                     label = "p.signif")

# FBRI Violin
df_80 %>%
  group_by(Country.risk) %>%
  ggplot(aes(x= Country.risk, y= FBRI)) +
  geom_violin(position = position_dodge(0.9), 
               fill = "lightblue1", width = 0.3) + 
  stat_summary(fun = mean, colour="darkred", geom="point", size = 2.5,
               position = position_dodge(0.9), show.legend = FALSE) + 
  stat_summary(fun = mean, colour ="black", geom ="text",
               aes(label = round(..y.., digits = 2)), position = position_dodge(0.9),
               vjust = 1.5, size = 3.8) +
  theme_bw() + 
  labs(x= "", y="",
       caption = "Figure 1: FBRI scores") + 
  theme(legend.position = "bottom",
        legend.text = element_text(size = 13),
        axis.text.x = element_text(size = 13),
        axis.text.y = element_text(size = 13)) + 
  scale_y_continuous(limits = c(0, 100),
                     breaks = c(0, 25, 50, 75, 100)) + 
  stat_compare_means(comparisons = my_comparisons, 
                     method = "t.test",
                     label.y = c(75, 80, 85, 90, 95, 100),
                     label = "p.signif",
                     symnum.args = list(cutpoints = c(0, 0.01, 0.05, 0.1, 1),
                                        symbols = c("***", "**", "*", "ns")))

compare_means(formula = FBRI ~ Country.risk,
              data = df_80,
              method = "t.test",
              symnum.args = list(cutpoints = c(0, 0.01, 0.05, 0.1, 1),
                                 symbols = c("***", "**", "*", "ns"))) %>%
  select(group1, group2, p, p.signif, method) %>%
  mutate(p = round(p, 3)) %>%
  flextable %>% autofit



# Operating country score
data <- df_80 %>%
  group_by(Country.risk) %>%
  summarize("Country.risk.score" = mean(Country.risk.score))

ggline(df_80, 
          x= "Country.risk", 
          y = "Country.risk.score", 
          add = c("mean_sd"),
          point.color = "lightblue3",
          color = "lightblue3",
          point.size = 11,
          width = 0.4) +
  labs(x= "", y= "Operating country score") + 
  theme_bw() + 
  theme(legend.text = element_text(size = 12),
        axis.text.x = element_text(size = 12),
        axis.text.y = element_text(size = 12),
        axis.title.y = element_text(size = 12)) + 
  geom_text(data = data,
            aes(label = round(Country.risk.score, 2)) , size = 3.5) + 
  scale_y_continuous(expand = c(0, 0), limits = c(0, 100))


# Bribery prevention

setwd("~/Norges Handelshøyskole/Arne Christoph Portmann - Master Thesis/Quantitative Analyses/Anti corruption score")

df_prevention <- read_xlsx("anti_corruption__scores.xlsx", sheet = 3)

df_prevention %>%
  group_by(`Country risk`) %>%
  summarise(score = mean(`Final score`)) %>%
  ggplot(aes(x= reorder(`Country risk`, -score), y=score, fill = `Country risk`)) +
  geom_col(position = position_dodge(0.9), width = 0.5) +
  theme_bw() + 
  geom_jitter(data = df_prevention, aes(x= `Country risk`, y= `Final score`), 
              alpha = 0.3, show.legend = FALSE, size = 1.3, width = 0.1) + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very low" = "lightblue1")) + 
  labs(x= "", y= "Bribery prevention score") + 
  theme(legend.position = "none",
        axis.text.x = element_text(size = 14),
        axis.text.y = element_text(size = 14),
        axis.title.y = element_text(size = 14)) + 
  geom_text(aes(label = score), vjust = 1.3, size = 6) + 
  scale_y_continuous(limits = c(0, 100),
                     expand = c(0.0, 0.05),
                     oob = scales::squish)
  

df_prevention <- df_prevention %>%
  select(`Country risk`, `Industry risk`, Company, `Item 1`:`Item 6`) %>%
  gather(data = .,
         key = Question,
         value = Score,
         `Item 1`:`Item 6`)

df_prevention$`Country risk` <- factor(df_prevention$`Country risk`,
                                       levels = c("Very low", "Low", "Moderate", "High"))


df_prevention$Question <- factor(df_prevention$Question, labels = c("1) Public commitment",
                                                                    "2) Third party compliance",
                                                                    "3) Anti-bribery training",
                                                                    "4) Gifts & Expenses",
                                                                    "5) Whistleblowing",
                                                                    "6) Corporate structure"))


df_prevention %>%
  group_by(`Country risk`, Question) %>%
  summarise(Score = mean(Score)) %>% 
  ggplot(aes(x= Question, y=Score, fill = `Country risk`)) +
  geom_col(position = position_dodge(0.9), width = 0.5) +
  theme_bw() +
  theme(legend.position = "bottom",
        axis.text.x = element_blank(),
        axis.ticks.x = element_blank(),
        strip.text.x = element_text(size = 12),
        legend.text = element_text(size = 14),
        axis.text.y = element_text(size = 12),
        axis.title.y = element_text(size = 14)) +
  labs(x= "", y="") + 
  facet_wrap( ~ Question, scales = "free", ncol = 2) + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very low" = "lightblue1")) + 
  geom_text(aes(label = Score), 
            position = position_dodge(0.9), vjust = 1.3, size = 3.5)



df_prevention %>%
  filter(Question %in% c("Item 5", "Item 6")) %>%
  group_by(`Country risk`, Question) %>%
  count(Score) %>%
  ggplot(aes(x=Score, y=n, fill = `Country risk`)) +
  geom_col(position = position_dodge(preserve = "single")) + 
  theme_bw() + 
  facet_grid( ~ Question, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very low" = "lightblue1")) + 
  scale_y_continuous(limits = c(0, 12, 3),
                     expand = c(0, 0)) + 
  theme(legend.position = "bottom",
        strip.text.x = element_text(size = 14),
        legend.text = element_text(size = 14),
        axis.text.y = element_text(size = 12),
        axis.title.y = element_text(size = 14),
        axis.title.x = element_text(size = 14),
        axis.text.x = element_text(size = 12)) +
  labs(y= "Firm count", x="Score")

df_80$Listing.Country.risk.score <- as.numeric(df_80$Listing.Country.risk.score)

# Ownership
df_80 <- df_80 %>%
  mutate(Ownership.Score = ifelse(Ownership.Score %in% c(50,100), 1, 0))

df_80 %>%
  group_by(Country.risk) %>%
  summarise(Mean = mean(Ownership.Score)) %>%
  ggplot(aes(x=Country.risk, y=Mean, fill = Country.risk)) +
  geom_col(width = 0.6) +
  theme_bw() +  
  scale_y_continuous(labels = scales::percent_format(accuracy = 1),
                     expand = c(0, 0),
                     limits = c(0, 0.5)) +
  labs(x= "", y="") + 
  geom_text(aes(label = paste0(Mean*100,"%")), size = 4.5, vjust = 1.3) + 
  theme(legend.position = "none",
        axis.text.y = element_text(size = 12),
        axis.text.x = element_blank(),
        strip.text.x = element_text(size = 14),
        axis.ticks.x = element_blank()) + 
  facet_grid( ~ Country.risk, scales = "free") + 
  scale_fill_manual(name = NULL, 
                    values = c("High" = "lightblue4", "Moderate" = "lightblue3", 
                               "Low" = "lightblue2", "Very low" = "lightblue1"))




# Correlation
df_80 %>%
  select(Listing.Country.risk.score, Country.risk.score, 
         Subsidiary.risk.score, PEP.score, Previous.allegations.score,
         Anti.corruption.effort.score, Ownership.Score) %>%
  rename("Listing country risk" = Listing.Country.risk.score,
         "Operating countries" = Country.risk.score,
         "Subsidiary industries" = Subsidiary.risk.score,
         "Bribery prevention" = Anti.corruption.effort.score,
         "Ownership" = Ownership.Score,
         "Previous allegations" = Previous.allegations.score,
         "Political exposure" = PEP.score) %>% 
  cor(.) %>% 
  corrplot(., method="color",  
           type="upper", 
           addCoef.col = "black",
           tl.col="black", 
           tl.srt=45,
           sig.level = 0.01, 
           diag=FALSE)



plot7 <- left_join(plot7, trace_matrix, by = "Country")

plot7 %>%
  select(Country, `Total Risk score`) %>%
  distinct(Country, .keep_all = T) %>%
  arrange(desc(`Total Risk score`))
  
