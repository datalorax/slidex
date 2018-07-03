context("test-convert_pptx.R")
path <- system.file("examples", "slidedemo.pptx", package = "slidex")
out <- convert_pptx(path)

test_that("Basic files are where they're supposed to be", {
  expect_equal(file.exists("slidedemo/slidedemo.Rmd"), TRUE)
  expect_equal(dir.exists("slidedemo/assets"), TRUE)
  expect_equal(dir.exists("slidedemo/xml"), FALSE)
})

unlink("slidedemo", recursive = TRUE, force = TRUE)

test_that("Error is thrown when file is not pptx", {
  expect_error(convert_pptx(substr(path, 1, nchar(path) - 1), "Daniel"))
  expect_equal(dir.exists("slidedemo_xml"), FALSE)
  expect_equal(file.exists("slidedemo.Rmd"), FALSE)
  expect_equal(dir.exists("assets"), FALSE)
})

test_that("Error is thrown when language encoding is not English", {
  frpptx <- system.file("examples", "slidex-fr.pptx", package = "slidex")

  expect_error(convert_pptx(frpptx, "Daniel"))
  expect_equal(dir.exists("slidex-fr"), FALSE)
})
