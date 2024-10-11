#' This function is to generate FAIR eDNA data templates, based on the sample types, detection type, and requirement levels of your choice. 
#'
#' Instructions 
#' 
#' Step 1: Save the input file in the working directory
#' The current version is data_types_fig_eDNA_data_checklist_v6_20240821.xlsx 
#' 
#' Step 2: Run the eDNA_temp_gen_fun function with the below arguments
#' 
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
#' eDNA_temp_gen_fun(req_lev = c('M', 'R', 'O'), sample_type = c('Water', 'Sediment'), detection_type='multi taxon detection', project_id = "gbr2022", assay_name = c("MiFish", "crust16S")) 

eDNA_temp_gen_fun = function(req_lev = c('M', 'R', 'O'), sample_type, detection_type, project_id, assay_name, 
                             studyMetadata_user = NULL, sampleMetadata_user = NULL) {
  # install packages
  packages <- c("readxl", "openxlsx", "RColorBrewer")
  for (i in packages) {
    if (!require(i, character.only = TRUE)) {
      install.packages(i, dependencies = TRUE)
      library(i, character.only = TRUE)
    }
  }
  
  input_file_name <- "eDNA_data_checklist_v7_20241004.xlsx"
  sheet_name <- "checklist"
  data <- read_excel(input_file_name, sheet = sheet_name)
  
  # create a directory for output templates
  dir.create(paste('template', project_id, sep='_'))
  
  #README
  readme1 <- c('The templates were generated using the eDNA checklist version of;', input_file_name, '',
               'Date/Time generated;', format(Sys.time(), '%Y-%m-%dT%H:%M:%S'))
  readme2 <- c('The templates were generated based on the below arguments;',
               paste('project_id =', project_id),
               paste('assay_name =', paste(assay_name, collapse = ' | ')),
               paste('detection_type =', detection_type),
               paste('req_lev =', paste(req_lev, collapse = ' | '))
               )
  if (sample_type == 'other') {
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
  
  write.table(readme.df, paste0('template_', project_id, '/README.txt'), row.names = F, col.names = F, quote = F)

  ### Make row2keep and row2rm lists
  ## requirement_level
  for (i in req_lev) {
    ls_temp <- grep(i, data$requirement_level_code)
    if (i == req_lev[1]) req_lev_row2keep <- ls_temp else req_lev_row2keep <- unique(c(req_lev_row2keep, ls_temp))
  }
  
  ## detection_type
  unique(data$section)
  if (detection_type == 'targeted taxon detection') { #TO-DO: include option of whether sequencing was done to confirm species
    detect_type_row2rm <- which(data$section %in% c('Library preparation/sequencing', 'Bioinformatics', 'OTU/ASV'))
  } else if (detection_type == 'multi taxon detection') {
    detect_type_row2rm <- which(data$section == 'Targeted taxon detection')
  } else if (detection_type == 'other') {
    detect_type_row2rm <- NA
  }
  detect_type_row2rm
  
  ## sample_type
  samp_type_all=c("Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", "SymbiontAssociated")
  if (sample_type=='other') {
    samp_type_row2rm=NULL
  } else {
    (samp_type_row2rm <- which(rowSums(data[,sample_type]) == 0)) #remove rows that have 0 in all sample_type column
    data[samp_type_row2rm, sample_type] #all zero. good
  }
  
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
  if (sample_type=='other') {
    col_ls <- c("requirement_level_code", "field_type", "controlled_vocabulary_options", samp_type_all, "section", "example","field_name")
  } else {
    col_ls <- c("requirement_level_code", "field_type", "controlled_vocabulary_options", sample_type, "section", "example","field_name")  
  }
  temp <- data_shortls[,col_ls]
  temp$study_level <- NA

  # remove sample_type cols if all NA
  if (sample_type=='other') {
    for (i in samp_type_all) {
      if (all(is.na(temp[,i]))) temp[,i] <- NULL
    }
  } else {
    for (i in sample_type){
      if (all(is.na(temp[,i]))) temp[,i] <- NULL
    }  
  }
  
  
  # Pre-fill project_id, assay_name, detection_type
  temp[which(temp$field_name=='project_id'),'study_level'] = project_id
  temp[which(temp$field_name=='assay_name'),'study_level'] = paste(assay_name, collapse =" | ")
  temp[which(temp$field_name=='detection_type'),'study_level'] = detection_type 
  temp[which(temp$field_name=='checkls_ver'),'study_level'] = input_file_name 
  
  if (length(assay_name)>1) {
    for (i in 1:length(assay_name)) {
      temp[,paste0('assay',i)]=NA
      temp[which(temp$field_name=='assay_name'),paste0('assay',i)] = assay_name[i]
    }
     
  }
  
  ## add user defined fields
  if (!is.null(studyMetadata_user)) {
    user.df <- data.frame(matrix(ncol=ncol(temp), nrow=length(studyMetadata_user)))
    colnames(user.df) <- colnames(temp)
    user.df$field_name <- studyMetadata_user 
    user.df$requirement_level_code <- 'O'
    user.df$field_type <- 'free text'
    user.df$section <- 'User defined'
    temp <- rbind(temp, user.df)
  }

  ## Add req_lev full name for the first time appearing
  if (any(temp$requirement_level_code == "M")) {
    temp[grep('M', temp$requirement_level_code)[1], 'requirement_level_code'] = 'M=Mandatory'
  }
  if (any(temp$requirement_level_code == "R")) {
    temp[grep('R', temp$requirement_level_code)[1], 'requirement_level_code'] = 'R=Recommended'
  }
  if (any(temp$requirement_level_code == "O")) {
    temp[grep('O', temp$requirement_level_code)[1], 'requirement_level_code'] = 'O=Optional'
  }
  
  wb <- createWorkbook() # create a excel workbook
  
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
  colour_ls2 <- rep(colour_ls, length.out = length(section_ls))
  if (length(colour_ls)>8) colour_ls2 <- c(colour_ls, colour_ls[9:(length(section_ls)-8)])
  for (i in 1:nrow(temp)) {
    for (j in 1:ncol(temp)) {
      for (k in 1:length(section_ls)) {#fill colours based on sections
        if (!is.na(temp[i,j]) && temp[i,j] == unique(section_ls)[k]) {
          addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), rows = i+1, col = j, gridExpand = TRUE, stack = TRUE) #+1 to account for column headers
        }
      }
      # fill red, yellow, green for 'mandated', 'recommended', 'optional'
      if (!is.na(temp[i,j]) && temp[i,j] =='M') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='M=Mandatory') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='R') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='R=Recommended') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='O') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp[i,j]) && temp[i,j] =='O=Optional') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), rows = i+1, cols = j, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  #README 
  addWorksheet(wb = wb, sheetName='README')
  writeData(wb, 'README', readme.df, colNames = F) 
  # Fill colours in README
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFCC00"), rows = grep('M = Mandatory', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFFF99"), rows = grep('R = Recommended', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#CCFF99"), rows = grep('O = Optional', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  
  saveWorkbook(wb, paste0('template_', project_id, '/', data_type, '_', project_id,'.xlsx'), overwrite = T)
  
  #### sampleMetadata ####
  data_type <- 'sampleMetadata'
  data_type_row2keep <- grep(data_type, data$data_type)
  
  ## Make a checklist that has only the specified data_type, req_lev, detection_type, and sample_type 
  intersect(data_type_row2keep, req_lev_row2keep) #list of rows that occur in both data_type_row2keep and req_lev_row2keep
  unique(c(samp_type_row2rm, detect_type_row2rm)) #list of rows that occur in either samp_type_row2rm or detect_type_row2rm
  (row2keep_ls <- setdiff(intersect(data_type_row2keep, req_lev_row2keep), unique(c(samp_type_row2rm, detect_type_row2rm)))) #take out the rows that are in *row2rm
  length(row2keep_ls)
  #data[samp_type_row2rm,]$requirement_level #There was no change with setdif() as all the samp_type_row2rm were optional
  data_shortls <- data[row2keep_ls,]
  unique(data_shortls$data_type) #check if correct
  unique(data_shortls$requirement_level_code)#check if correct
  unique(data_shortls$section)#check if correct
  
  colnames(data_shortls)
  if (sample_type=='other') {
    col_ls <- c("requirement_level_code", "field_type", "controlled_vocabulary_options", samp_type_all, "section", "example","field_name")
  } else {
    col_ls <- c("requirement_level_code", "field_type", "controlled_vocabulary_options", sample_type, "section", "example","field_name")  
  }
  temp <- data.frame(t(data_shortls[,col_ls]))
  colnames(temp) <- data_shortls$field_name
  colnames(temp)[1] <- '# field_name'
  # remove sample_type rows if they are all NA (i.e. when req_lev = 'M') 
  if (sample_type=='other') {
    for (i in samp_type_all) {
      if (all(is.na(temp[i,]))) temp <- temp[-which(rownames(temp)==i),]
    }
  } else {
    for (i in sample_type){
      if (all(is.na(temp[i,]))) temp <- temp[-which(rownames(temp)==i),]
    }  
  }

  # add a row on the bottom showing example of samp_name
  temp2 <- rbind(temp,temp[1,]) # temp[1,] can be any row
  temp2[nrow(temp2),] <- NA # remove the entry of the newly added, bottom row
  temp2[nrow(temp2),1] <- paste('# e.g.,', temp2['example', 1]) # add samp_name example
  
  # add "#" at the front of the rows of fields' information
  temp2[1:(nrow(temp2)-2),1] <- paste('#', rownames(temp2[1:(nrow(temp2)-2),]))
  temp2[,1:5]
  
  ## add user defined fields
  if (!is.null(sampleMetadata_user)) {
    user.df <- data.frame(matrix(nrow=nrow(temp2), ncol=length(sampleMetadata_user)))
    colnames(user.df) <- sampleMetadata_user
    rownames(user.df) <- rownames(temp2)
    user.df['field_name',] <- sampleMetadata_user 
    user.df['requirement_level_code',] <- 'O'
    user.df['field_type',] <- 'free text'
    user.df['section',] <- 'User defined'
    temp2 <- cbind(temp2, user.df)
  }
  
  ## Add req_lev full name for the first time appearing
  if (any(temp2['requirement_level_code',] == "M")) {
    temp2['requirement_level_code', grep('M', temp2['requirement_level_code',])[1]] = 'M=Mandatory'
  }
  if (any(temp2['requirement_level_code',] == "R")) {
    temp2['requirement_level_code', grep('R', temp2['requirement_level_code',])[1]] = 'R=Recommended'
  }
  if (any(temp2['requirement_level_code',] == "O")) {
    temp2['requirement_level_code', grep('O', temp2['requirement_level_code',])[1]] = 'O=Optional'
  }

  wb <- createWorkbook() # create a excel workbook
  
  addWorksheet(wb=wb, sheetName=data_type)
  writeData(wb, data_type, temp2, colNames = F) 
  
  # bold text
  addStyle(wb, sheet = data_type, style = bold_style,
           rows=which(rownames(temp2)=='field_name'), cols=1:ncol(temp2), gridExpand = T)
  # fill colours based on section and requirement level
  section_ls <- unique(data_shortls$section)
  colour_ls <- brewer.pal(8,'Pastel1')
  colour_ls2 <- rep(colour_ls, length.out = length(section_ls))
  for (i in 1:nrow(temp2)) {
    for (j in 1:ncol(temp2)) {
      for (k in 1:length(section_ls)) {#fill colours based on sections
        if (!is.na(temp2[i,j]) && temp2[i,j] == unique(section_ls)[k]) {
          addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), 
                   rows = i, col = j, gridExpand = TRUE, stack = TRUE)
        }
      }
      # fill red, yellow, green for 'mandated', 'recommended', 'optional'
      if (!is.na(temp2[i,j]) && temp2[i,j] =='M') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp2[i,j]) && temp2[i,j] =='M=Mandatory') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp2[i,j]) && temp2[i,j] =='R') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp2[i,j]) && temp2[i,j] =='R=Recommended') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp2[i,j]) && temp2[i,j] =='O') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
      if (!is.na(temp2[i,j]) && temp2[i,j] =='O=Optional') {
        addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                 rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
      }
    }
  }
  
  #README 
  addWorksheet(wb = wb, sheetName='README')
  writeData(wb, 'README', readme.df, colNames = F) 
  # Fill colours in README
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFCC00"), 
           rows = grep('M = Mandatory', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#FFFF99"), 
           rows = grep('R = Recommended', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  addStyle(wb, sheet = 'README', style = createStyle(fgFill = "#CCFF99"), 
           rows = grep('O = Optional', readme.df[,1]), cols = 1, gridExpand = T, stack = T)
  
  
  saveWorkbook(wb, paste0('template_', project_id, '/', data_type, '_', project_id,'.xlsx'), overwrite = T)
  
  #### DNA sequence data ####
  if (detection_type == 'multi taxon detection') {
    data_type <- 'dnaSeqData' 
    data_type_row2keep <- grep(data_type, data$data_type)
    
    ## Make a checklist that has only the specified data_type, req_lev and detection_type 
    row2keep_ls <- intersect(data_type_row2keep, req_lev_row2keep) #list of rows that occur in both data_type_row2keep and req_lev_row2keep
    length(row2keep_ls)
    data_shortls <- data[row2keep_ls,]
    unique(data_shortls$data_type) #check if correct
    unique(data_shortls$requirement_level_code)#check if correct
    unique(data_shortls$section)#check if correct
    
    colnames(data_shortls)
    col_ls <- c("requirement_level_code", "field_type", "example","field_name") #Section is not included here as they are all OTU/ASV
    temp <- data.frame(t(data_shortls[,col_ls]))
    colnames(temp) <- data_shortls$field_name
    colnames(temp)[1] <- '# field_name'
    
    # add "#" at the front of the rows of fields' information
    temp[1:(nrow(temp)-1),1] <- paste('#', rownames(temp[1:(nrow(temp)-1),]))
    
    wb <- createWorkbook() # create a excel workbook
    addWorksheet(wb=wb, sheetName=data_type)
    writeData(wb, data_type, temp, colNames = F) 
    
    # bold text
    addStyle(wb, sheet = data_type, style = bold_style,
             rows=which(rownames(temp)=='field_name'), cols=1:ncol(temp), gridExpand = T)
    
    for (i in 1:nrow(temp)) {
      for (j in 1:ncol(temp)) {
        for (k in 1:length(section_ls)) {#fill colours based on sections
          if (!is.na(temp[i,j]) && temp[i,j] == unique(section_ls)[k]) {
            addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), 
                     rows = i, col = j, gridExpand = TRUE, stack = TRUE)
          }
        }
        # fill red, yellow, green for 'mandated', 'recommended', 'optional'
        if (!is.na(temp[i,j]) && temp[i,j] =='M') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
        if (!is.na(temp[i,j]) && temp[i,j] =='R') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
        if (!is.na(temp[i,j]) && temp[i,j] =='O') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
      }
    }
    for (i in assay_name) {
      saveWorkbook(wb, paste0('template_', project_id, '/', data_type, '_', project_id, '_', i, '_ENTER library_id HERE.xlsx'), overwrite = T)
    }
    
  } 
  
  
  #### Amplification data ####
  if (detection_type == 'targeted taxon detection') {
    data_type <- 'amplificationData'
    data_type_row2keep <- grep(data_type, data$data_type)
    ## Make a checklist that has only the specified data_type, req_lev and detection_type 
    row2keep_ls <- intersect(data_type_row2keep, req_lev_row2keep) #list of rows that occur in both data_type_row2keep and req_lev_row2keep
    length(row2keep_ls)
    data_shortls <- data[row2keep_ls,]
    unique(data_shortls$data_type) #check if correct
    unique(data_shortls$requirement_level_code)#check if correct
    unique(data_shortls$section)#check if correct
    
    colnames(data_shortls)
    col_ls <- c("requirement_level_code", "field_type", "controlled_vocabulary_options", "example","field_name") #Section is not included here as they are all OTU/ASV
    temp <- data.frame(t(data_shortls[,col_ls]))
    colnames(temp) <- data_shortls$field_name
    colnames(temp)[1] <- '# field_name'
    
    # add "#" at the front of the rows of fields' information
    temp[1:(nrow(temp)-1),1] <- paste('#', rownames(temp[1:(nrow(temp)-1),]))
    
    wb <- createWorkbook() # create a excel workbook
    addWorksheet(wb=wb, sheetName=data_type)
    writeData(wb, data_type, temp, colNames = F) 
    
    # bold text
    addStyle(wb, sheet = data_type, style = bold_style,
             rows=which(rownames(temp)=='field_name'), cols=1:ncol(temp), gridExpand = T)
    
    for (i in 1:nrow(temp)) {
      for (j in 1:ncol(temp)) {
        for (k in 1:length(section_ls)) {#fill colours based on sections
          if (!is.na(temp[i,j]) && temp[i,j] == unique(section_ls)[k]) {
            addStyle(wb, sheet=data_type, style = createStyle(fgFill = colour_ls2[k]), 
                     rows = i, col = j, gridExpand = TRUE, stack = TRUE)
          }
        }
        # fill red, yellow, green for 'mandated', 'recommended', 'optional'
        if (!is.na(temp[i,j]) && temp[i,j] =='M') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFCC00"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
        if (!is.na(temp[i,j]) && temp[i,j] =='R') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#FFFF99"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
        if (!is.na(temp[i,j]) && temp[i,j] =='O') {
          addStyle(wb, sheet = data_type, style = createStyle(fgFill = "#CCFF99"), 
                   rows = i, cols = j, gridExpand = TRUE, stack = TRUE)
        }
      }
    }
    
  
    for (i in assay_name) {
      saveWorkbook(wb, paste0('template_', project_id, '/', data_type, '_', project_id, '_', i, '.xlsx'), overwrite = T)
    }
    
    
      
  }

}


##### Examples ####
#Rachel's data
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('Water'), detection_type = 'multi taxon detection', 
                   project_id='verte_quadeloupe', assay_name = c("vert01", "tele01", "mamm01", "cetac"))
eDNA_temp_gen_fun(sample_type=c('Water'), detection_type = 'multi taxon detection', 
                  project_id='verte_quadeloupe', assay_name = c("vert01", "tele01", "mamm01", "cetac"))

#Katrina's data
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('Water'), detection_type = 'multi taxon detection', 
                   project_id='IOT-eDNA', assay_name = c("16SFish", "COILeray"))

#Bruce's plankton data
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('other'), detection_type = 'multi taxon detection', 
                  project_id='Plankton-Recorder', assay_name = 'COI-metazoan')

# Louie's sediment data (qPCR)
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('Sediment'), detection_type = 'targeted taxon detection', 
                  project_id='AEP_Fish_sedDNA', assay_name = c('eESLU1', 'eCOAR7'))

# Neha's water data (qPCR)
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('Water'), detection_type = 'targeted taxon detection', 
                  project_id='Rockfish_targeted_qPCR', assay_name = c('eSEMA3', 'eSEPA9', 'eSERU5'))

# test
eDNA_temp_gen_fun(req_lev=c('M', 'R', 'O'), sample_type=c('Sediment'), detection_type = 'multi taxon detection', 
                  project_id='test_Luke', assay_name = 'COI-metazoan', studyMetadata_user = 'expedition_id', 
                  sampleMetadata_user = c('line_id', 'moonphase'))
eDNA_temp_gen_fun(req_lev=c('M'), sample_type=c('other'), detection_type = 'targeted taxon detection', 
                  project_id='testtest', assay_name = 'COI-metazoan')
