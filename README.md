# FAIRe-ator (FAIR eDNA template generator)
This R function generates the FAIR eDNA metadata templates, customised to fit to users eDNA study designs (e.g., sample type, assay type, requirement level)

**STEP 1.** Set your working directory in R. 


**STEP 2.** Download input files (checklist and FULL template) and save them in your working directory.

- [FAIRe_checklist_v1.0.2.xlsx](https://raw.githubusercontent.com/FAIR-eDNA/FAIRe_checklist/main/FAIRe_checklist_v1.0.2.xlsx)
- [FAIRe_checklist_v1.0.2_FULLtemplate.xlsx](https://raw.githubusercontent.com/FAIR-eDNA/FAIRe_checklist/main/FAIRe_checklist_v1.0.2_FULLtemplate.xlsx)

Note: Previous versions of the checklist, FULL templates, and change logs are available [here](https://github.com/FAIR-eDNA/FAIRe_checklist/tree/main)


**STEP 3.** Source the `FAIReator.R` script to load the function into your R environment.

```
source("https://raw.githubusercontent.com/FAIR-eDNA/FAIRe-ator/refs/heads/main/FAIReator.R")

```

The FAIReator function requires the following parameters:

1. `sample_type` A (list of) Sample type(s). Select one or more from "Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", and "SymbiontAssociated", or "other". "other" will include all the sample-type-specific fields in the output. 
1. `assay_type` An approach applied to detect taxon/taxa. Select one "targeted" (e.g., q/d PCR based detection) or "metabarcoding" (e.g., metabarcoding).
1. `project_id` A brief, concise project identifier with no spaces or special characters. This ID will be used in data file names as 'project_id'.
1. `assay_name` A brief, concise assay name(s) with no spaces or special characters. This will be used in data file names as 'assay_name'. 

The three optional parameters are:

1. `req_lev` Requirement level(s) of each fields to include in the data template. Select one or more from "M", "HR", "R", and "O" (Mandatory, Highly recommended, Recommended, and Optional, respectively). Default is c("M", "HR", "R", "O").
1. `projectMetadata_user` (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the projectMetadata.
1. `sampleMetadata_user` (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the sampleMetadata.


**STEP 4.** Run the function, according to your needs. For example:

```
FAIReator(
  sample_type = c('Water', 'Sediment'),
  assay_type = 'metabarcoding',
  project_id = 'gbr2022',
  assay_name = c('MiFish', 'crust16S')
)
                   
```

Note: Make sure the two input files are closed (not open in another program), otherwise the script will not run properly.

The expected outputs are:

- A new folder called "template_<project_id>" (e.g., template_gbr2022) in your working directory.
- An Excel file called "<project_id>.xlsx" (e.g., gbr2022.xlsx) in the new template folder, containing multiple worksheets in the file, including README, projectMetadata, sampleMetadata and other assay_type specific tables (e.g., OTU table if assay_type = "metabarcoding")

