library(testthat)
library(mockery)

context("theme_bsol")

test_that("it returns a ggplot2 theme", {
  theme <- theme_bsol()

  expect_s3_class(theme, c("theme", "gg"), TRUE)
})

test_that("it calls theme_classic", {
  m <- mock(cycle = TRUE)

  stub(theme_bsol, "theme_classic", m)

  theme_bsol()
  theme_bsol("")

  expect_called(m, 2)
  expect_args(m, 1, base_family = "Open Sans")
  expect_args(m, 2, base_family = "")
})
