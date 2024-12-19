#' This function is to generate FAIR eDNA data templates, based on the sample types, detection type, and requirement levels of your choice. 
#'
#' Instructions 
#' 
#' Step 1: Save the input file in the working directory
#' The current version is data_types_fig_eDNA_data_checklist_v7_20241004.xlsx 
#' 
#' Step 2: Run the eDNA_temp_gen_fun function with the below arguments
#' 
#' We need to describe the expected output
#' Arguments
#' 
#' @req_lev Requirement level(s) of each fields to include in the data template. Select one or more from "M", "R", and "O" (Mandatory, Recommended, and Optional, respectively). Default is c("M", "R", "O")
#' @sample_type A (list of) Sample type(s). Select one or more from "Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", and "SymbiontAssociated", or "other". "other" will include all the sample-type-specific fields in the output. 
#' @detection_type An approach applied to detect taxon/taxa. Select one "targeted taxon detection" (e.g., q/d PCR based detection) or "multi taxon detection" (e.g., metabarcoding)
#' @project_id A brief, concise project identifier with no spaces or special characters. This ID will be used in data file names as 'project_id'.
#' @assay_name A brief, concise assay name(s) with no spaces or special characters. This will be used in data file names as 'assay_name'. 
#' @studyMetadata_user (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the studyMetadata.
#' @sampleMetadata_user (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the sampleMetadata.
#' 
#' @examples
#' eDNA_temp_gen_fun(req_lev = c('M', 'R', 'O'), 
#'                   sample_type = c('Water', 'Sediment'), 
#'                   detection_type='multi taxon detection', 
#'                   project_id = 'gbr2022', 
#'                   assay_name = c('MiFish', 'crust16S')) 

#TO-DO: include option of whether sequencing was done to confirm species
eDNA_temp_gen_fun = function(req_lev = c('M', 'R', 'O'), 
                             sample_type, 
                             detection_type, 
                             project_id, 
                             assay_name, 
                             studyMetadata_user = NULL, 
                             sampleMetadata_user = NULL) {
  
  # install packages --------------------------------------------------------
  
  packages <- c("readxl", "openxlsx", "RColorBrewer", "dplyr")
  for (i in packages) {
    if (!require(i, character.only = TRUE)) {
      install.packages(i, dependencies = TRUE)
      library(i, character.only = TRUE)
    }
  }
  

  # set input and output ----------------------------------------------------

  #input_file_name <- "eDNA_data_checklist_v7_20241004.xlsx"
  #sheet_name <- "checklist" #changed in v7
  #sheet_name <- "list_v7"
  # input <- readxl::read_excel(input_file_name, sheet = sheet_name) %>% 
  #   dplyr::mutate(requirement_level_code = dplyr::recode(requirement_level, #This column wasn't in v7
  #                                                        'Mandatory' = "M", 
  #                                                        'Recommended' = "R", 
  #                                                        'Optional'= "O")) %>% 
  #   tidyr::separate_longer_delim(cols = data_type, delim = " | ") #This separates rows that are relevant for multiple data types
  
  t <- tempfile()
  input_file_name <- googledrive::drive_get(id = "1RiJYCI-cSWRcCYJXLNSiFKgqfJslqKf_") %>% pull(name)
  googledrive::drive_download(file = googledrive::as_id("1RiJYCI-cSWRcCYJXLNSiFKgqfJslqKf_"), path = t)
  
  input <- readxl::read_excel(t, sheet = "checklist") %>% 
      tidyr::separate_longer_delim(cols = data_type, delim = " | ") #This separates rows that are relevant for multiple data types
  
  # create a directory for output templates
  if(!dir.exists(paths = "./template")){dir.create(paste('template', project_id, sep='_'))}
  
  wb <- createWorkbook() # create a excel workbook
  
  # Make README -------------------------------------------------------------
  
  readme1 <- c('The templates were generated using the eDNA checklist version of;', 
               input_file_name, 
               '',
               'Date/Time generated;', 
               format(Sys.time(),
                      '%Y-%m-%dT%H:%M:%S')) 
  
  
  
  readme2 <- c('The templates were generated based on the below arguments;',
               paste('project_id =', project_id),
               paste('assay_name =', paste(assay_name, collapse = ' | ')),
               paste('detection_type =', detection_type),
               paste('req_lev =', paste(req_lev, collapse = ' | '))
  )
  
  if (any(sample_type %>% tolower() == 'other')) {
    readme2 <- c(readme2, 
                 paste('sample_type =', paste(sample_type, collapse = ' | '), 
                       '(Note: this option provides sample-type-specific fields for ALL sample types)')
    )
  } else {
    readme2 <- c(readme2, 
                 paste('sample_type =', paste(sample_type, collapse = ' | '), 
                       '(Note: this option provides sample-type-specific fields for the selected sample type(s))')
    )
  }
  
  if (!is.null(studyMetadata_user)) {
    readme2 <- c(readme2, paste('studyMetadata_user =', paste(studyMetadata_user, collapse = ' | ')))
  }
  if (!is.null(sampleMetadata_user)) {
    readme2 <- c(readme2, paste('sampleMetadata_user =', paste(sampleMetadata_user, collapse = ' | ')))
  }
  
  readme3 <- c('Requirement levels;', 'M = Mandatory', 'R = Recommended', 'O = Optional')
  
  readme4 <- c('List of files;',
               paste0('studyMetadata_', project_id),
               paste0('sampleMetadata_', project_id)             
  )
  if (detection_type=='multi taxon detection') {
    readme4=c(readme4,
              paste('rawOTU', project_id, assay_name, '<library_id>', sep = '_'),
              paste('otu', project_id, assay_name, '<library_id>', sep = '_'),
              paste('dnaSeqData', project_id, assay_name, '<library_id>', sep = '_'),
              'Note: rawOTU, otu and dnaSeqData should be produced for each assay_name and library_id',
              'Note: <library_id> in the file names should match with library_id in your sampleMetadata'
    )
    
  } else if (detection_type=='targeted taxon detection') {
    readme4=c(readme4, 
              paste('amplificationData', project_id, assay_name, sep = '_'),
              'Note: amplificationData should be produced for each assay_name'
    )
  }
  
  readme.df=data.frame(c(readme1, '', readme2, '', readme3, '', readme4))
  
  
  addWorksheet(wb = wb, sheetName='README')
  writeData(wb, 'README', readme.df, colNames = F) 
  
  # Fill colours in README
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFCC00"), rows = grep('M = Mandatory', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFFF99"), rows = grep('R = Recommended', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#CCFF99"), rows = grep('O = Optional', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  

  # filter checklist ------------------------------------------------
  
  #remove terms based on sample type
  samp_type_all <- c("Water", 
                     "Soil", 
                     "Sediment", 
                     "Air", 
                     "HostAssociated", 
                     "MicrobialMatBiofilm", 
                     "SymbiontAssociated")
  
  samp_type_row2keep <- if(any(sample_type %>% tolower() == 'other')) {
    NA
    } else {
  samp_type_row2keep <- apply(input[,sample_type], 1, function(x) any(!is.na(x) & !is.null(x)))
  }
    
  #remove terms based on detection type
  section2keep <- if(detection_type == 'multi taxon detection') {
    c('Library preparation/sequencing', 'Bioinformatics', 'OTU/ASV')
  } else if(detection_type == 'Targeted taxon detection') {
    'Targeted taxon detection'
  } else if (detection_type %>% tolower() == 'other') {
    NA
  }
  
  
  # filter and split by data type and section
  input <- input %>% 
    filter(requirement_level_code %in% req_lev &
             (section %in% c(section2keep) |
                data_type %in% c("sampleMetadata") |
                samp_type_row2keep)
           ) %>% 
    select(data_type, section, requirement_level_code, example, field_type, field_name) %>% 
    group_by(data_type, .add = TRUE) %>% 
    {setNames(group_split(.), 
              group_keys(.)[[1]])}
  
#Make other [which ones?] sheets
  purrr::imap(input, function(x,y){
    
    rlc <- x[["requirement_level_code"]] %>% sort() #for use below
    
    x <- x %>% 
      select(requirement_level_code, section, field_type, example, field_name) %>% 
      arrange(section, requirement_level_code) %>% 
      t
    
    y <- names(input[y])
    addWorksheet(wb = wb, sheetName = y)
    
    writeData(wb, y, x, colNames = F, rowNames = TRUE)


    
    # Fill colours in README
    addStyle(wb, sheet = y, 
             style = createStyle(fgFill = "#FFCC00"), 
             rows = 1, 
             cols = grep(x = rlc, pattern = 'M') + 1, 
             gridExpand = T, 
             stack = T)
    
    addStyle(wb, sheet = y, 
             style = createStyle(fgFill = "#FFFF99"), 
             rows = 1, 
             cols = grep(x = rlc, pattern = 'R') + 1, 
             gridExpand = T, 
             stack = T)
    
    addStyle(wb, sheet = y, 
             style = createStyle(fgFill = "#CCFF99"), 
             rows = 1, 
             cols = grep(x = rlc, pattern = 'O') + 1, 
             gridExpand = T, 
             stack = T)
    
    addStyle(wb, sheet = y, 
             style = createStyle(fgFill = "gray"), 
             rows = 2, 
             cols = 1:length(rlc) + 1,
             gridExpand = T, 
             stack = T)
    
    addStyle(wb, y, 
             style = createStyle(textDecoration = "bold"),
             rows = 5,
             cols = 1:length(rlc) + 1,
             gridExpand = T)
    
  })
  
  for (i in assay_name) {
    saveWorkbook(wb, paste0(paste('template', project_id, sep='_'),
                            "/",
                            Sys.Date(), '_tester_', project_id, '_', i, '_ENTER library_id HERE.xlsx'), overwrite = T)
  }
  
}
