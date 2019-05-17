library(officer)
library(magrittr) # need to run every time you start R and want to use %>%
library(dplyr)
library(stringi)

#defining the document name
my_doc <- read_docx() 



barPlot <- function(ip){
  src1 <- tempfile(fileext = ".png")
  png(filename = src1, width = 5, height = 6, units = 'in', res = 300)
  barplot(ip)
  dev.off()
  src1 <<- src1
}

boxPlot <- function(ip){
  src2 <- tempfile(fileext = ".png")
  png(filename = src2, width = 5, height = 6, units = 'in', res = 300)
  boxplot(ip)
  dev.off()
  src2 <<- src2
}

scatterPlot <- function(ip){
  src3 <- tempfile(fileext = ".png")
  png(filename = src3, width = 5, height = 6, units = 'in', res = 300)
  scatter.smooth(ip)
  dev.off()
  src3 <<- src3
}



writing <- function(what="text",vars.name.pass=NULL,vars.val.pass=NULL,analysis=NULL){
  
  
    
  
  if(what=="text"){
    
    if(analysis=="mean"){
      TextInfomation1 <<-  paste(vars.name.pass, " mean is:", mean(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation1, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
        
    }
    
    if(analysis=="sd"){
      
      TextInfomation2 <<-  paste(vars.name.pass, " sd is:", sd(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation2, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
      
    }
    
    if(analysis=="median"){
      
      TextInfomation3 <<-  paste(vars.name.pass, " median is:", median(data[,vars.name.pass]))
      dyn.command <- paste(dyn.command,'body_add_par(TextInfomation3, style = "Normal") %>%') 
      dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
      
    }
    
  }
  
  if(what=="scatter"){
    src3 <<- scatterPlot(data[,paste(vars.name.pass)])
    Textscatter <<-  'Scatter Plot is: '
    dyn.command <- paste(dyn.command,'body_add_par(Textscatter, style = "Normal") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src3, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="bar"){
    src1 <<- barPlot(data[,paste(vars.name.pass)])
    Textbar <<-  'Bar Plot is: '
    dyn.command <- paste(dyn.command,'body_add_par(Textbar, style = "Normal") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src1, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="box"){
    src2 <<- boxPlot(data[,paste(vars.name.pass)])
    Textbox <<-  'Box Plot is: '
    dyn.command <- paste(dyn.command,'body_add_par(Textbox, style = "Normal") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command,'body_add_img(src = src2, width = 5, height = 6, style = "centered") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }
  
  if(what=="table"){
    TableToSave <<- head(data)
    Texthead <<-  'A short example of your data to analyze: '
    dyn.command <- paste(dyn.command,'body_add_par(Texthead, style = "Normal") %>%')
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
    dyn.command <<- paste(dyn.command, 'body_add_table(TableToSave, style = "table_template") %>%') 
    dyn.command <<- paste(dyn.command,'body_add_par("", style = "Normal") %>%') # blank paragraph
  }

  
}

#Example with iris dataset
data <- iris


for(var in 1:(ncol(data)-1)){
  
  dyn.command <<- 'my_doc <- my_doc %>%'
  
  if(var ==1){
    writing(what = "table",vars.name.pass = NULL,analysis = NULL)
  }
  
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "mean")
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "sd")
  writing(what = "text",vars.name.pass = paste(names(data)[var]),analysis = "median")
  writing(what = "bar",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  writing(what = "box",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  writing(what = "scatter",vars.name.pass = paste(names(data)[var]),analysis = NULL)
  
  dyn.command <<- noquote(stri_replace_last_fixed(dyn.command, " %>%", ""))

  #browser()
  
  eval(parse(text=dyn.command))
  
  print(my_doc, target = "Temp.docx")
  
}


