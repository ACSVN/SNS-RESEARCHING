#install.packages("magrittr")
#install.packages("flextable")
#install.packages("ggplot2")

library(officer)
library(magrittr) 
library(flextable)
library(ggplot2)
library(imager)
library(stringr)


# ---------- set project directory (change wd from "/source/*" to "/*") ---------- #
setwd(str_replace(getwd(), "/sources", ""))


# ---------- set argument from bat file ---------- #
args <- commandArgs(trailingOnly = TRUE)
client_name <- as.character(args[1])
Twitter_id <- as.character(args[2])
screen_switch <- as.character(args[3])
# client_name <- "test-client"
# Twitter_id <- "test"
# screen_switch <- 1


# ---------- set path for images, powerpoint file ---------- #
cappng_dir <- str_c(getwd(), "/images/Cap.png")
toppng_dir <- str_c(getwd(), "/images/top.png")
ppt_file <- str_c(getwd(), "/output/", Twitter_id, "/Twitter-PPT.pptx")


# ---------- set trimming position ---------- #
imgTrimPosition <- array(0, dim = c(2, 2))

if (screen_switch == 1) {
  imgTrimPosition[1, 1] <- 560
  imgTrimPosition[2, 1] <- 150
  imgTrimPosition[1, 2] <- 1100
  imgTrimPosition[2, 2] <- 1000
  width_size <- 3
  left_position <- 1
} else {
  imgTrimPosition[1, 1] <- 170
  imgTrimPosition[1, 2] <- 130
  imgTrimPosition[2, 1] <- 780
  imgTrimPosition[2, 2] <- 680
  width_size <- 4.8
  left_position <- 0
}

# ---------- trimming img and save as "top.png" ---------- #
img = load.image(cappng_dir)
img <- imsub(img, x < imgTrimPosition[1, 2], y < imgTrimPosition[2, 2]) #(First) trim right & bottom 
img <- imsub(img, x > imgTrimPosition[1, 1], y > imgTrimPosition[2, 1]) #(second) trim left & top
save.image(img, file = toppng_dir)


# ---------- create powerpoint object ---------- #
mypptx <- read_pptx(ppt_file)


# ---------- Page 1 ---------- #
mypptx <- on_slide(mypptx, index = 1)
slide_summary(mypptx, index = 1)
mypptx <- ph_with_text(mypptx, type = "ctrTitle", str = client_name, index = 1)


# ---------- Page 2 ---------- #
mypptx <- on_slide(mypptx, index = 2)
mypptx <- ph_with_img_at(x = mypptx, src = toppng_dir, height = 5, width = width_size, 
                         left = left_position, top = 1.2, rot = 45)


# ---------- Page 3 ---------- #
mypptx <- on_slide(mypptx, index = 3)
mypptx <- ph_with_img_at(x = mypptx, src = toppng_dir, height = 5, width = width_size, 
                         left = left_position, top = 1.2, rot = 45)


print(mypptx, ppt_file) %>%
  invisible()