library(tidyverse)
library(webshot)
library(magick)
library(flextable)
library(ggplot2)
library(rmarkdown)

# setwd
setwd("C:/Users/Stephen/Desktop/R/rmarkdown")

# convert hmtl table to png with webshot, then convert png to pdf with magick

# save r script to tempfile for rendering as an rmarkdown doc
r_script_name <- tempfile(fileext = ".R")
cat("mtcars %>% head() %>% regulartable()", file = r_script_name)

# create temporary html file for output of rendering
html_name <- tempfile(fileext = ".html")

# render r script as if it were rmarkdown
# http://brooksandrew.github.io/simpleblog/articles/render-reports-directly-from-R-scripts/
render(input = r_script_name, output_format = "html_document", 
       output_file = html_name)

# get a png from the html file with webshot
# note that png and pdf will have higher resolution the higher the zoom
# 1 vs 5 is very noticeable; 10 vs 20 is less so until you enlarge pdf to see fine details
webshot(html_name, zoom = 20, delay = .5, file = "regulartable.png", 
        selector = "table")

# read png into magick
table_png_image <- image_read("regulartable.png")
table_png_image

# output png as pdf
image_write(table_png_image, path = "regulartable.pdf", format = "pdf")



