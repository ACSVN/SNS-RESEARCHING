#install.packages("magrittr")
#install.packages("flextable")
#install.packages("ggplot2")

library(officer)
library(magrittr) 
library(flextable)
library(ggplot2)
library(imager)

##-------------------------create img
img = load.image("C:/SNS-TWITTER/images/Cap.png")
img <-imsub(img, x<1100, y<1000)
img <-imsub(img, x>560, y>150)


save.image(img, file="C:/SNS-TWITTER/images/top.png")
##-----------------------get and set arg
args <- commandArgs(trailingOnly = TRUE)
client_name <- as.character(args[1])
Twitter_id <- as.character(args[2])

#client_name = "sample"
#client_name <- paste0(client_name,"Report")

##-----------------------Page 1
pptxpath <- paste0(getwd(),"/output/", "Twitter-PPT.pptx")
mypptx <- read_pptx(pptxpath)
mypptx <- on_slide( mypptx, index = 1)
slide_summary(mypptx, index = 1)
mypptx <- ph_with_text(mypptx, type = "ctrTitle", str = client_name, index = 1)

##-----------------------Page 2
mypptx <- on_slide( mypptx, index = 2)
topcaption <- ("C:/SNS-TWITTER/images/top.png")
mypptx <- ph_with_img_at(x = mypptx, src = topcaption, height = 5, width = 3,
                         left = 1, top = 1.2, rot = 45 )
##-----------------------Page 3
mypptx <- on_slide( mypptx, index = 3)
mypptx <- ph_with_img_at(x = mypptx, src = topcaption, height = 5, width = 3,
                         left = 1, top = 1.2, rot = 45 )
#top10caption <- ("C:/Rproject/result/top10.png")
#mypptx <- on_slide( mypptx, index = 4)
#mypptx <- ph_with_img_at(x = mypptx, src = top10caption, height = 5, width = 10,
#                         left = 0, top = 2, rot = 45 )
#worst10caption <- ("C:/Rproject/result/worst10.png")
#mypptx <- on_slide( mypptx, index = 5)
#mypptx <- ph_with_img_at(x = mypptx, src = worst10caption, height = 5, width = 10,
#                         left = 0, top = 2, rot = 45 )
#heikinchicaption <- ("C:/Rproject/result/heikinchi.png")
#mypptx <- on_slide( mypptx, index = 6)
#mypptx <- ph_with_img_at(x = mypptx, src = heikinchicaption, height = 5, width = 10,
#                         left = 0, top = 2, rot = 45 )
#chuohchicaption <- ("C:/Rproject/result/chuohchi.png")
#mypptx <- on_slide( mypptx, index = 7)
#mypptx <- ph_with_img_at(x = mypptx, src = chuohchicaption, height = 5, width = 10,
#                         left = 0, top = 2, rot = 45 )

print(mypptx, pptxpath) %>%
  invisible()