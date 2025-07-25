test_that("FAIReator works", {
  test <- FAIReator(
    sample_type = c("Water", "Sediment"),
    assay_type = "metabarcoding",
    project_id = "gbr2022",
    assay_name = c("MiFish", "crust16S"),
    return_wb_object = TRUE
  )

  system_file_path <- system.file("extdata",
    package = "fairr"
  )
  check_file <- paste0(system_file_path, "/demo_data/test-output.rds")
  check <- readRDS(check_file)
  test_project <- openxlsx::readWorkbook(
    test,
    sheet = "projectMetadata"
  )
  check_project <- openxlsx::readWorkbook(
    check,
    sheet = "projectMetadata"
  )
  # This can varying depending on testing environment so make the same
  test_project$project_level[14] <- ""
  check_project$project_level[14] <- ""
  testthat::expect_equal(test_project, check_project)
})
