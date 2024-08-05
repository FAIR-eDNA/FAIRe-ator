#### This function is to generate FAIR eDNA data templates, based on the sample types, detection type, and requirement levels of your choice. 

#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
#### Instruction 

### Step 1. Set working directory
#i.e., 
setwd("C:/Users/tak025/OneDrive - CSIRO/Documents/Miwa/IWY_eDNAdata/data_std/data_checklist/temp_generator")

### Step 2: Save the input file in the directory
# The current version is data_types_fig_eDNA_data_checklist_v4_20240716.xlsx 

### Step 3: Read the eDNA_temp_gen_fun function 

### Step 4: Run the function 
#i.e., eDNA_temp_gen_fun(req_lev = c('M', 'R'), sample_type = c('Water', 'Sediment'), detection_type='multispecies detection') 


#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

eDNA_temp_gen_fun = function(input_file_name, req_lev, sample_type, detection_type) {
  # install packages
  packages <- c("readxl", "openxlsx", "RColorBrewer")
  for (i in packages) {
    if (!require(i, character.only = TRUE)) {
      install.packages(i, dependencies = TRUE)
      library(i, character.only = TRUE)
    }
  }
  
  input_file_name <- "data_types_fig_eDNA_data_checklist_v4_20240716.xlsx"
  sheet_name <- "checklist"
  data <- read_excel(input_file_name, sheet = sheet_name)
  
  ### Make row2keep and row2rm lists
  ## requirement_level
  for (i in req_lev) {
    ls_temp <- grep(i, data$requirement_level_code)
    if (i == req_lev[1]) req_lev_row2keep <- ls_temp else req_lev_row2keep <- unique(c(req_lev_row2keep, ls_temp))
  }
  
  ## detection_type
  unique(data$section)
  if (detection_type == 'targeted species detection') { #TO-DO: include option of whether sequencing was done to confirm species
    detect_type_row2rm <- which(data$section %in% c('Library preparation/sequencing', 'Bioinformatics', 'Sequence info'))
  } else if (detection_type == 'multispecies detection') {
    detect_type_row2rm <- which(data$section == 'Targeted species detection')
  } else if (detection_type == 'other') {
    detect_type_row2rm <- NA
  }
  detect_type_row2rm
  
  ## sample_type
  (samp_type_row2rm <- which(rowSums(data[,sample_type]) == 0)) #remove rows that have 0 in all sample_type column
  data[samp_type_row2rm, sample_type] #all zero. good
  
  #### studyMetadata ####
  data_type <- 'studyMetadata'
  data_type_row2keep <- grep(data_type, data$data_type)
  
  ## Make a checklist that has only the specified data_type, req_lev, detection_type, and sample_type 
  intersect(data_type_row2keep, req_lev_row2keep) #list of rows that occur in both data_type_row2keep and req_lev_row2keep
  unique(c(samp_type_row2rm, detect_type_row2rm)) #list of rows that occur in either samp_type_row2rm or detect_type_row2rm
  (row2keep_ls <- setdiff(intersect(data_type_row2keep, req_lev_row2keep), unique(c(samp_type_row2rm, detect_type_row2rm)))) #take out the rows that are in *row2rm
  length(row2keep_ls)
  data[samp_type_row2rm,]$requirement_level #There was no change with setdif() as all the samp_type_row2rm were optional
  data_shortls <- data[row2keep_ls,]
  unique(data_shortls$data_type) #check if correct
  unique(data_shortls$requirement_level_code)#check if correct
  unique(data_shortls$section)#check if correct
  
  ## make a template 
  colnames(data_shortls)
  col_ls <- c("requirement_level_code", "field_type", sample_type, "section", "example","field_name")
  temp <- data_shortls[,col_ls]
  temp$entries <- NA
  temp[1,'entries'] <- paste('e.g.,', temp[1,'example'])
  # remove sample_type cols if all NA
  for (i in sample_type){
    if (all(is.na(temp[,i]))) temp[,i] <- NULL
  }
  # write the template in .xlsx format
  (output_file_name <- paste('template', data_type, gsub(' ','',detection_type), paste(sample_type,collapse = ''), paste(req_lev,collapse = ''), sep = '_'))
  wb <- createWorkbook()
  addWorksheet(wb = wb, sheetName=data_type)
  writeData(wb, data_type, temp, colNames = T) 
  # bold text
  bold_style <- createStyle(textDecoration = "bold") 
  addStyle(wb, sheet = data_type, style = bold_style,
           rows = 1:(nrow(temp)+1), cols = which(colnames(temp) == 'field_name'), gridExpand = T) #+1 to account for col headers
  addStyle(wb, sheet = data_type, style = bold_style,
           rows = 1, cols = 1:ncol(temp), gridExpand = T)
  # fill colours based on section and requirement level
  section_ls <- unique(data_shortls$section)
  colour_ls <- brewer.pal(8,'Pastel1')
  colour_ls2 <- colour_ls[1:length(section_ls)]
  if (length(colour_ls)>8) colour_ls2 <- c(colour_ls, colour_ls[9:(length(section_ls)-8)])
  for (i in 1:nrow(temp)) {
    for (j in 1:ncol(temp)) {
      for (k in 1:length(section_ls)) {#fill colours based on sections
        if (!is.na(temp[i,j]) && temp[i,j] == unique(section_ls)[k]) {
          addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), 
                   rows = i+1, col = j, gridExpand = TRUE, stack = TRUE) #+1 to account for column headers
        }
      }
      # fill red, yellow, green for 'mandated', 'recommended', 'optional'
      if (!is.na(temp[i,j]) && temp[i,j] =='M') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                 rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='R') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                 rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='O') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                 rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  #### sampleMetadata ####
  data_type <- 'sampleMetadata'
  data_type_row2keep <- grep(data_type, data$data_type)
  
  ## Make a checklist that has only the specified data_type, req_lev, detection_type, and sample_type 
  intersect(data_type_row2keep, req_lev_row2keep) #list of rows that occur in both data_type_row2keep and req_lev_row2keep
  unique(c(samp_type_row2rm, detect_type_row2rm)) #list of rows that occur in either samp_type_row2rm or detect_type_row2rm
  (row2keep_ls <- setdiff(intersect(data_type_row2keep, req_lev_row2keep), unique(c(samp_type_row2rm, detect_type_row2rm)))) #take out the rows that are in *row2rm
  length(row2keep_ls)
  data[samp_type_row2rm,]$requirement_level #There was no change with setdif() as all the samp_type_row2rm were optional
  data_shortls <- data[row2keep_ls,]
  unique(data_shortls$data_type) #check if correct
  unique(data_shortls$requirement_level_code)#check if correct
  unique(data_shortls$section)#check if correct
  
  colnames(data_shortls)
  col_ls <- c("requirement_level_code", "field_type", sample_type, "section", "example","field_name")
  temp <- data.frame(t(data_shortls[,col_ls]))
  colnames(temp) <- data_shortls$field_name
  colnames(temp)[1] <- '# field_name'
  # remove sample_type rows if they are all NA (i.e. when req_lev = 'M') 
  for (i in sample_type){ 
    if (all(is.na(temp[i,]))) temp5 <- temp[-which(rownames(temp)==i),]
  }
  # add a row on the bottom showing example of samp_name
  temp2 <- rbind(temp,temp[1,]) # temp[1,] can be any row
  temp2[nrow(temp2),] <- NA # remove the entry of the newly added, bottom row
  temp2[nrow(temp2),1] <- paste('# e.g.,', temp2['example', 1]) # add samp_name example
  
  # add "#" at the front of the rows of fields' information
  temp2[1:(nrow(temp2)-2),1] <- paste('#', rownames(temp2[1:(nrow(temp2)-2),]))
  temp2[,1:5]
  
  # add a few rows at front wiht # for description
  (ls <- c(paste('# This template is from the checklist', input_file_name),
        '# Requirement level M=Mandatory, R=Recommended, O=Optional'))
  descr.df <- data.frame(matrix(nrow=length(ls), ncol=ncol(temp2)))
  colnames(descr.df) <- colnames(temp2)
  descr.df[,1] <- ls
  temp3 <- rbind(descr.df, temp2)
  
  addWorksheet(wb=wb, sheetName=data_type)
  writeData(wb, data_type, temp3, colNames = F) 
  
  # bold text
  addStyle(wb, sheet = data_type, style = bold_style,
           rows=which(rownames(temp3)=='field_name'), cols=1:ncol(temp3), gridExpand = T)
  # fill colours based on section and requirement level
  section_ls <- unique(data_shortls$section)
  colour_ls <- brewer.pal(8,'Pastel1')
  colour_ls2 <- colour_ls[1:length(section_ls)]
  for (i in 1:nrow(temp3)) {
    for (j in 1:ncol(temp3)) {
      for (k in 1:length(section_ls)) {#fill colours based on sections
        if (!is.na(temp3[i,j]) && temp3[i,j] == unique(section_ls)[k]) {
          addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), 
                   rows = i, col = j, gridExpand = TRUE, stack = TRUE)
        }
      }
      # fill red, yellow, green for 'mandated', 'recommended', 'optional'
      if (!is.na(temp3[i,j]) && temp3[i,j] =='M') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp3[i,j]) && temp3[i,j] =='R') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp3[i,j]) && temp3[i,j] =='O') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  #### README ####
  # TO DO
  #data_type='README'
  #addWorksheet(wb=wb, sheetName=data_type)
  
  #### otu/amplification data ####
  # TO DO
  
  #### DNA sequence data ####
  # TO DO
  
  ## Save
  output_file_name <- paste('template', gsub(' ','',detection_type), paste(req_lev,collapse = ''), paste(sample_type,collapse = ''), sep = '_')
  saveWorkbook(wb, paste0(output_file_name, '.xlsx'), overwrite = T)  
}


## req_lev = c("M", "O", "R")
## sample_type = c("Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", "SymbiontAssociated")
## detection_type = "multispecies detection", "targeted species detection", "other"
# Note for detection_type - Multiple answers NOT permitted. If 'other' is selected, there is no specific fields to include/exclude so the output (template) will have both targeted and multi species detection fields

eDNA_temp_gen_fun(input_file_name = "list_v5_eDNA_data_checklist_v4_20240716.csv", 
             req_lev = c('M', 'R'), sample_type = c('Water', 'Sediment'), detection_type='multispecies detection') 
eDNA_temp_gen_fun(input_file_name = "list_v5_eDNA_data_checklist_v4_20240716.csv", 'M', c('Water','Sediment'), detection_type='multispecies detection') 
eDNA_temp_gen_fun(input_file_name = "list_v5_eDNA_data_checklist_v4_20240716.csv", 
             c('M', 'R', 'O'), c('Water','Sediment'), detection_type='targeted species detection') 

