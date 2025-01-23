# FAIRe-ator (FAIR eDNA template generator)
This R function generates the FAIR eDNA metadata templates, customised to fit to users eDNA study designs (e.g., sample type, assay type, requirement level)

To use this function, save the following two files in your R working directory:

- FAIRe_checklist_v1.0.xlsx
- FAIRe_checklist_v1.0_FULLtemplate.xlsx

Download and source the `FAIReator.R` script to load the function into your R environment.

The FAIReator function requires the following parameters:

1. `req_lev` Requirement level(s) of each fields to include in the data template. Select one or more from "M", "R", and "O" (Mandatory, Recommended, and Optional, respectively). Default is c("M", "R", "O")
1. `sample_type` A (list of) Sample type(s). Select one or more from "Water", "Soil", "Sediment", "Air", "HostAssociated", "MicrobialMatBiofilm", and "SymbiontAssociated", or "other". "other" will include all the sample-type-specific fields in the output. 
1. `assay_type` An approach applied to detect taxon/taxa. Select one "targeted" (e.g., q/d PCR based detection) or "metabarcoding" (e.g., metabarcoding)
1. `project_id` A brief, concise project identifier with no spaces or special characters. This ID will be used in data file names as 'project_id'.
1. `assay_name` A brief, concise assay name(s) with no spaces or special characters. This will be used in data file names as 'assay_name'. 

The two optional parameters are:

1. `projectMetadata_user` (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the projectMetadata.
1. `sampleMetadata_user` (optional) A user-defined field or list of fields that are not listed in the FAIR eDNA metadata checklist. These fields will be appended to the end of the sampleMetadata.


Run the function, according to your needs. For example:


```
FAIReator(
  req_lev = c('M', 'HR', 'R', 'O'),
  sample_type = c('Water', 'Sediment'),
  assay_type = 'metabarcoding',
  project_id = 'gbr2022',
  assay_name = c('MiFish', 'crust16S')
)
                   
```