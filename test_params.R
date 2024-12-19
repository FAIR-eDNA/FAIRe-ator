req_lev <- c('M', 'R', 'O')
sample_type <- c('Water', 'Sediment')
detection_type <- 'multi taxon detection'
project_id <- 'TEST1234'
assay_name <- c('MiFish', 'ABCD')
studyMetadata_user <- NULL 
sampleMetadata_user <- NULL

eDNA_temp_gen_fun(
  req_lev = req_lev,
  sample_type = sample_type,
  detection_type = detection_type,
  project_id = project_id,
  assay_name = assay_name,
  studyMetadata_user = studyMetadata_user,
  sampleMetadata_user = sampleMetadata_user
)

# second version ----------------------------------------------------------


req_lev <- c('M')
sample_type <- c('Water', 'Sediment', 'Air', 'Other')
detection_type <- 'Targeted taxon detection'
project_id <- 'test2'
assay_name <- c('MiFish', 'crust16S')
studyMetadata_user <- NULL 
sampleMetadata_user <- NULL

input_file_name <- "eDNA_data_checklist_v7_20241004.xlsx"

#sheet_name <- "checklist" #changed in v7
sheet_name <- "list_v7"
input <- readxl::read_excel(input_file_name, sheet = sheet_name) %>% 
  dplyr::mutate(requirement_level_code = dplyr::recode(requirement_level, #This column wasn't in v7
                                                       'Mandatory' = "M", 
                                                       'Recommended' = "R", 
                                                       'Optional'= "O")) %>% 
  tidyr::separate_longer_delim(cols = data_type, delim = " | ") #This separates rows that are relevant for multiple data types

