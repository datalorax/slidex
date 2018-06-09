context("test-convert_pptx.R")

test_that("Basic files are where they're supposed to be", {
  path <- system.file("examples", "slidedemo.pptx", package = "slidex")
  out <- convert_pptx(path, "Daniel")

  #expect_equal(file.exists("slidedemo.Rmd"), TRUE)
  expect_equal(dir.exists("assets"), TRUE)
  expect_equal(dir.exists("slidedemo_xml"), FALSE)

  expect_error(convert_pptx(pptx, "Daniel"))
  expect_equal(dir.exists("slidedemo_xml"), FALSE)
})

unlink("assets", recursive = TRUE, force = TRUE)
unlink("slidedemo.Rmd", recursive = TRUE, force = TRUE)

test_that("Error is thrown when file is not pptx", {
  path <- system.file("examples", "slidedemo.pptx", package = "slidex")

  expect_error(convert_pptx(substr(path, 1, nchar(path) - 1), "Daniel"))
  expect_equal(dir.exists("slidedemo_xml"), FALSE)
  expect_equal(file.exists("slidedemo.Rmd"), FALSE)
  expect_equal(dir.exists("assets"), FALSE)
})

test_that("Error is thrown when language encoding is not English", {
  frpptx <- system.file("examples", "slidex-fr.pptx", package = "slidex")

  expect_error(convert_pptx(frpptx, "Daniel"))
  expect_equal(dir.exists("slidex-fr_xml"), FALSE)
  expect_equal(file.exists("slidex-fr.Rmd"), FALSE)
  expect_equal(dir.exists("assets"), FALSE)
})
