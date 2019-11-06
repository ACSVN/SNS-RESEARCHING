# install.packages("officer")
# install.packages("magrittr")
# install.packages("flextable")
# install.packages("ggplot2")
# install.packages("imager")
# install.packages("stringr")

library(officer)
library(magrittr)
library(flextable)
library(ggplot2)
library(imager)
library(stringr)



# ---------- set project directory (change wd from "/source/*" to "/*") ---------- #
setwd(str_replace(getwd(), "/sources", ""))
setwd(str_replace(getwd(), "/bat", ""))
# getwd()


# ---------- set argument from bat file ---------- #

args <- commandArgs(trailingOnly = TRUE)
client_name <- as.character(args[1])
Twitter_id <- as.character(args[2])
screen_switch <- as.character(args[3])



# ---------- function to trim cap.png ---------- #

func_trimImage <- function(loadFile, saveFile, topLeft_x, topLeft_y, bottomRight_x, bottomRight_y) {
  img = load.image(loadFile)
  img <- imsub(img, x < bottomRight_x, y < bottomRight_y)
  img <- imsub(img, x > topLeft_x, y > topLeft_y)
  save.image(img, file = saveFile)
}


cap_png <- str_c(getwd(), "/images/Cap.png")
top_png <- str_c(getwd(), "/images/top.png")

if (screen_switch == 1) {
  func_trimImage(cap_png, top_png, 560, 150, 1100, 1000)
} else {
  func_trimImage(cap_png, top_png, 170, 130, 780, 680)
}



# ---------- function to paste for powerpoint object ---------- #
func_pasteTitle <- function(pptObj, slideNum, title) {
  pptObj <- on_slide(pptObj, index = slideNum)
  pptObj <- ph_with_text(pptObj, type = "ctrTitle", str = title, index = slideNum)
}

func_pasteImage <- function(pptObj, slideNum, Image, h, w, l, t, r) {
  pptObj <- on_slide(pptObj, index = slideNum)
  pptObj <- ph_with_img_at(x = pptObj, src = Image, height = h, width = w, left = l, top = t, rot = r)
}


ppt_file <- str_c(getwd(), "/output/", Twitter_id, "/Twitter-PPT.pptx")
mypptx <- read_pptx(ppt_file)

func_pasteTitle(mypptx, 1, client_name)
if (screen_switch == 1) {
  func_pasteImage(mypptx, 2, top_png, 5.0, 3.0, 1.0, 1.2, 45)
  func_pasteImage(mypptx, 3, top_png, 5.0, 3.0, 1.0, 1.2, 45)
} else {
  func_pasteImage(mypptx, 2, top_png, 5.0, 4.8, 0.0, 1.2, 45)
  func_pasteImage(mypptx, 3, top_png, 5.0, 4.8, 0.0, 1.2, 45)
}

print(mypptx, ppt_file) %>%
  invisible()
