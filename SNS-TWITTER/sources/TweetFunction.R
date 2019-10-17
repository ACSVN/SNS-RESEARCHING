# install.packages("rtweet")
# install.packages("tidytext")
# install.packages("stringr")
# ------------------------------------------------------- #

## load twitter library - the rtweet library is recommended now over twitteR
library(rtweet)
## text mining library
library(tidytext)
## concat string
library(stringr)
# ------------------------------------------------------- #

##
## func_setPremiumAccount auto set premium addixcs account
##
func_setPremiumAccount <- function(){
  appname <- "ADDIXCS-Twitter"
  consumer_key <- ""
  consumer_secret <-""
  access_token <- ""
  access_secret <- ""
  
  twitter_token <- create_token(app = appname,
                                consumer_key = consumer_key,
                                consumer_secret = consumer_secret,
                                access_token = access_token,
                                access_secret = access_secret)
  
  message("Finished set token file")
}


##
## func_setTokenAccount will get parameters, then set token account for program, 
## after that will create the token file
## @params appName: is name of application which we register on twitter develop
## @params consumerKey: is a key of twitter created
## @params consumerSecret: is a key of twitter created
## @params accessToken: is a key which we created in application on twitter develop
## @params accessCecret: is a key which we created in application on twitter develop
## @params tokenName: is name of file we want to create
##
func_setTokenAccount <- function(appName, consumerKey, consumerSecret, 
                                 accessToken, accessCecret, tokenName){
  ## create token
  twitter_token <- create_token(app = appName,
                                consumer_key = consumerKey,
                                consumer_secret = consumerSecret,
                                access_token = accessToken,
                                access_secret = accessCecret)
  
  ## path of home directory 
  home_directory <- path.expand(paste0(getwd(), "/tokens"))  #"~"
  
  ## combine with name for token
  file_name <- file.path(home_directory, tokenName)
  
  ## save token to home directory
  saveRDS(twitter_token, file = file_name)
  
  message("Finished set and create token file")
}

##
## func_getTokenAccount will set your token into .Renviromen
## (when you chen token file you must close R program and open R again)
## params tokenName is name of token file
##
func_getTokenAccount <- function(tokenName){
  # --------------- Using token
  ## path of home directory 
  home_directory <- path.expand(paste0(getwd(), "/tokens"))  #"~"
  
  ## combine with name for token
  file_name <- file.path(home_directory, tokenName)
  
  ## save token to home directory
  # saveRDS(twitter_token, file = file_name)
  
  ## assuming you followed the procodures to create "file_name"
  ## from the previous code chunk, then the code below should
  ## create and save your environment variable.
  cat(paste0("TWITTER_PAT=", file_name),
      file = file.path(home_directory, ".Renviron"),
      append = TRUE)
  
  message("Finished set tweeter token file")
}

##
## func_getUserTweet_premium use premium account (search in 30 days)
## @params username: is name which we want search ()
## @params numberRecord: is number what you want to get it
##
func_getUserTweet_30Days_premium <- function(username, numberRecord){
  infoSearch = str_c("from:", username)
  
  gmtDate <- as.POSIXlt(Sys.time(), "GMT") # the current time in UTC
  todate <- format(gmtDate, "%Y%m%d%H%M")
  
  no = as.numeric(substr(todate, 11, 12)) - 1
  print(no)
  mydate = ""
  if(no > 10){
    mydate <- paste0(substr(todate, 1, 10), no)
  }else{
    mydate <- paste0(substr(todate, 1, 10), "0", no)
  }
  
  userInfo <- search_30day(q = infoSearch, n = numberRecord, 
                           toDate = mydate, env_name = "DevSearch30Days")
  
  userData <- data.frame(
    userInfo$created_at[1:numberRecord],      # Created date
    unlist(userInfo$media_url, use.names=FALSE)[1:numberRecord],       # Creative
    userInfo$text[1:numberRecord],            # Content
    userInfo$favorite_count[1:numberRecord],  # Like count
    userInfo$reply_count[1:numberRecord],     # Reply count
    userInfo$retweet_count[1:numberRecord],   # Retweet count
    userInfo$followers_count[1:numberRecord]  # Follower count
  )
  colnames(userData) <- c("Date", "Creative", "Content", "Like", "Reply", "Retweet", "Follower")
  
  message("Finished get data from Tweeter")
  
  return(userData)
}

##
## func_getUserTweet_standard will get all information of 1 user
## @params username: is name which we want search ()
## @params numberRecord: is number what you want to get it
##

func_getUserTweet_standard <- function(username, numberRecord){
  
  infoSearch = str_c("from:", username)
  
  userInfo <- search_tweets( 
    q = infoSearch, #screen_name="@realDonaldTrump",
    n = numberRecord, type = "recent", include_rts = FALSE,
    geocode = NULL, max_id = NULL, parse = TRUE, token = NULL,
    retryonratelimit = FALSE, verbose = TRUE, exclude_replies=TRUE,
    include_entities=TRUE
  )
  
  userData <- data.frame(
    userInfo$created_at,      # Created date
    unlist(userInfo$media_url, use.names=FALSE),       # Creative
    userInfo$text,            # Content
    userInfo$favorite_count,  # Like count
    userInfo$reply_count,     # Reply count
    userInfo$retweet_count,   # Retweet count
    userInfo$followers_count  # Follower count
  )
  colnames(userData) <- c("Date", "Creative", "Content", "Like", "Reply", "Retweet", "Follower")
  
  message("Finished get data from Tweeter")
  
  return(userData)
}
