req_lev <- c('M', 'R', 'O')
sample_type <- c('Water', 'Sediment')
detection_type <- 'multi taxon detection'
project_id <- 'gbr2022'
assay_name <- c('MiFish', 'crust16S')
studyMetadata_user <- 'User Not Named' 
sampleMetadata_user <- 'User Not Named'

input_file_name <- "eDNA_data_checklist_v7_20241004.xlsx"

#sheet_name <- "checklist" #changed in v7
sheet_name <- "list_v7"
input <- readxl::read_excel(input_file_name, sheet = sheet_name) %>% 
  dplyr::mutate(requirement_level_code = dplyr::recode(requirement_level, #This column wasn't in v7
                                                       'Mandatory' = "M", 
                                                       'Recommended' = "R", 
                                                       'Optional'= "O")) %>% 
  tidyr::separate_longer_delim(cols = data_type, delim = " | ") #This separates rows that are relevant for multiple data types


# second version ----------------------------------------------------------


req_lev <- c('M')
sample_type <- c('Water', 'Sediment', 'Air', 'Other')
detection_type <- 'Targeted taxon detection'
project_id <- 'test2'
assay_name <- c('MiFish', 'crust16S')
studyMetadata_user <- 'User Not Named' 
sampleMetadata_user <- 'User Not Named'

input_file_name <- "eDNA_data_checklist_v7_20241004.xlsx"

#sheet_name <- "checklist" #changed in v7
sheet_name <- "list_v7"
input <- readxl::read_excel(input_file_name, sheet = sheet_name) %>% 
  dplyr::mutate(requirement_level_code = dplyr::recode(requirement_level, #This column wasn't in v7
                                                       'Mandatory' = "M", 
                                                       'Recommended' = "R", 
                                                       'Optional'= "O")) %>% 
  tidyr::separate_longer_delim(cols = data_type, delim = " | ") #This separates rows that are relevant for multiple data types

