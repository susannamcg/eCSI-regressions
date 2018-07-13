#note: 2126 not available for FY14, asking that question will crash the program
library(openxlsx)

##PART 0: Obtain get data based on user input  
  
  cutoff <- .3 #use only questions that have single regression r vals over this cutoff
  nind <- 3 #number of to satisfaction indicators to print at the end
  
  userInput = winDialogString("Rail Line:", "line......")
  railLine = gsub(" ","",userInput) #removes all spaces in the user input
  dataSet <- read.csv(paste(railLine, ".csv", sep = ""))
  #questions = winDialogString("Questions (seperate with ', '):", "questions...")
  
  #Uncommebtn below and comment above to edit more easily (removes user input)
  #userInput = "Ethan Allen"
  #railLine = "EthanAllen"
  #dataSet = read.csv("EthanAllen.csv")
  questions = "1010, 1020, 1050, 1060, 1070, 1080, 1090, 1100, 1110, 2030, 2040, 2050, 2060, 2070, 2080, 2090, 2100, 2110, 2120, 2124, 2126, 2130, 2420, 2430, 2440, 2450, 2460, 3010, 3020, 3030, 3040"
  
  #format questions: a list with 'X'es before numeric questions (the way r reads the titles)
  questions <- c(strsplit(questions, ", ")[[1]])
  for(i in 1:length(questions)){
    if(questions[i] != "ArrP" & questions[i] != "DepP" & questions[i] != "Miles"){
      questions[i] = paste("X",questions[i], sep = "")
      #ensure all data columns are read as integers
      dataSet[,which(names(dataSet) == questions[i])] = as.integer(as.character(dataSet[,which(names(dataSet) == questions[i])]))
    }
  }
  dataSet$X4060 = as.integer(as.character(dataSet$X4060))
  
##PART 1: plot the effect of multiple variables on overall satisfaction (X4060)
  
  #to ensure that the axes are scaled for the data used in the regression 
  #(0-100 for most questions, no. of minutes for ArrP)
  if("ArrP" %in% questions){
    xmin = min(dataSet$ArrP)
    if(max(dataSet$ArrP > 100)){
      xmax = max(dataSet$ArrP)
    }
  } else{ 
    xmin = 0
    xmax = 100
  }
  x <- xmin:xmax
  
  #setting up the axes
  par(mfrow = c(1,1))
  NumColors = length(questions)
  rbow<-rainbow(NumColors)
  cols = rbow
  
  plot(x = NULL, 
       xlim = c(min(x), max(x)), ylim = c(0,100), 
       xlab = "Question Response", ylab = "Overall Satisfaction", 
       main = paste("overall satisfaction on", userInput, sep = " ")
  )
  
  #Loading the question names
  qnames <- read.csv("names.csv")
  qnames <- qnames[3:67,3]
  qbase = names(dataSet[22:87])
  rlist <- data.frame()
  
  #Plotting the questions
  for(i in 1:length(questions)){
    cq <- dataSet[names(dataSet) == questions[i]]
    
    #remove NULL question values from satisfaction and question data
    #(doing this in a for loop to minimize the erased data - this way, only null
    #data for a specific question is erased, not all null data). 
    #I erase data for all questions that have null data in any of the selected questions in part 2 - 
    #it might be better to do that here just because it is such a similar task and unnecessary to do it twice
    sat <- as.integer(dataSet$X4060)[cq != "NULL"]
    cq <- as.numeric(as.character(cq[cq != "NULL"]))
    
    #compute and plot single regressions for current question
    cqfit <-lm (sat ~ cq)
    coef <- coefficients(cqfit)
    fcqfit <- function(x) {coef[1] + x * coef[2]}
    adjrsq <- summary(cqfit)$adj.r.squared
    
    plot(fcqfit, col = cols[i], xlim = c(min(x), max(x)), ylim = c(0,100), add = TRUE)
    par(new = TRUE)
    points(sat ~ cq, col = cols[i])
    
    #match question code to question text
    if(substr(questions[i],1,1) != "X"){cq = questions[i]
    }else{
      cq <- as.character(qnames[qbase == questions[i]])
    }
    
    #creating a data frame of adjusted r squared values
    rval <- data.frame("Question" = questions[i], "Adj r squared" = adjrsq, "Text" = cq)
    rlist <- rbind(rlist, rval)
    
  }
  #clean up things created in loops
  legend("topleft", c(questions[1:i]), fill = c(cols[1:i]))
  cat("\n")
  
  #edit the list of questions to remove those with single regression Rsq vals below the cutoff
  rabove = c(which(rlist$Adj.r.squared <= cutoff))
  questions = questions[-rabove]
  
  #order and print list of r squared values
  rlist <- rlist[order(-rlist$Adj.r.squared),]
  print(rlist, right = F)
  
  #prevent program from crashing if there are no significant questions!
  if(length(questions) == 0){print("No questions had sufficient significance")
  }else{
    
    
    
##PART 2: Doing multiple regressions on most important questions from part 1
    
    #check the data to use only complete data in the regressions
    #*potential issue: might significantly reduce the sample size if too much data is incomplete
    valid <- as.matrix(dataSet$X4060 != "NULL")
    for(i in 1:length(valid)){
      if(is.na(valid[i])){
        valid[i] = FALSE
      }
    }
    for(i in 1:length(valid)){
      if(valid[i] == TRUE){
        for(j in 1:length(questions)){
          entry = dataSet[i,which(names(dataSet) == questions[j])]
          if(is.na(entry)){
            valid[i] = FALSE
          }
          break
        }
      }
    }
    
    #construct a data frame of the cleaned data for all questions:
    reg <- dataSet$X4060
    for (i in 1:length(questions)){
      qdata <- as.numeric(as.character(dataSet[,which(names(dataSet) == questions[i])]))
      reg = cbind(reg, qdata)
    }
    reg <- as.data.frame(reg)[valid,]
    names(reg) <- c("satisfaction", questions)
    
    #to generate all possible combinations of multiple regressions (no cross terms)
    fitslist = data.frame() 
    n = length(reg)-1
    for (i in 2:n){
      combo = combn(names(reg)[2:length(reg)],i)
      #print(i)
      for(j in 1:ncol(combo)){
        #print(paste("j:", j))
        vec = combo[,j]
        p <- paste(vec, collapse ="+")
        dep <- paste("satisfaction ~",p)
        crfit <- lm(as.formula(dep), data = reg)
        
        #creating a data frame of fits and their adjusted squared values, to later identify the best fit
        adjrsq <- summary(crfit)$adj.r.squared
        newline <- data.frame(i, j, adjrsq)
        fitslist <- rbind(fitslist, newline)
      }
    }
    
    #order list of fits high to low by adj r sq and extract i, j vals to identify the best fit
    fitslist <- fitslist[order(-fitslist$adjrsq),]
    i = fitslist[1,1]
    j = fitslist[1,2]
    combo = combn(names(reg)[2:length(reg)],i)
    vec = combo[,j]
    p <- paste(vec, collapse ="+")
    dep <- paste("satisfaction ~",p)
    bestfit <- lm(as.formula(dep), data = reg)
    
    #plot the fit indicators, order summary by coefficients, print summary:
    par(mfrow = c(2,2))
    plot(bestfit)
    bestfit$coefficients = bestfit$coefficients[order(-bestfit$coefficients)]
    print(summary(bestfit))
    
    #Saving the top n specifief indicators of satisfaction on specified train 
    adjrsq = 100* as.numeric(format(round(summary(crfit)$adj.r.squared, 2), nsmall = 2))
    
    topinds <- data.frame()
    n = nind
    for(i in 1:n){
      if(names(coefficients(bestfit))[i] == "(Intercept)"){
        n = n + 1  
      }else{
        ind <- data.frame("Indicator" = qnames[qbase == names(coefficients(bestfit))[i]], 
                          "Correlation" = coefficients(bestfit)[i])
        topinds <- rbind(topinds, ind)
      }
    }
    
    if(nrow(topinds) < nind){
      print(paste("Warning: did not find", nind, "indicators."))
      nind <- length(topinds)
    }
    
    print(paste("Top ", nind, "indicators of passenger satisfaction on ", userInput, "with", adjrsq, "% confidence:"))
    print(topinds)
  }
  
##PART 3: Exporting the data to Excel
  #Remove 'X'es from the question labels
  rlist$Question = as.character(rlist$Question)
  for(i in 1:length(rlist$Question)){
    rlist$Question[i] = substr(rlist$Question[i], 2, 99)}
  
  
  #write multiple regression coefficients and stats data into an excel spreadsheet
  coefs = as.data.frame(unclass(summary(bestfit))$coefficients)
  
  #adding question name and numbers to coefficient data frame
  Question = c()
  n = length(rownames(coefs))
  for(i in 1:n){
    if(rownames(coefs)[i] == "(Intercept)"){Question = append(Question, "Intercept")
    }else{
      Question = append(Question, substr(rownames(coefs)[i],2,99))
    }
  }
  Text = c()
  n = length(rownames(coefs))
  for(i in 1:n){
    if(rownames(coefs)[i] == "(Intercept)"){Text = append(Text, "")
    }else{
      Text = append(Text, as.character(qnames[qbase == rownames(coefs)[i]]))
    }
  }
  
  coefs = cbind(Question, coefs, Text)
  adjrsqline = data.frame("Adjusted r squared of fit:", summary(bestfit)$adj.r.sq, "", "", "", "")
  titles = c ("Question Num.","Coefficient Est.", "Std. Error", "t value", "Pr(>|t|)", "Question Text")
  fitsum = rbind(coefs, setNames(adjrsqline, names(coefs)))
  
  
  ## Create new workbooks
   wb <- createWorkbook() 
   
     ## Create the worksheets
  addWorksheet(wb, sheetName = "Single regression coefficients" )
  addWorksheet(wb, sheetName = "Multiple regression fit" )
  addWorksheet(wb, sheetName = "Top Indicators")
   
     ## Write the data
  writeData(wb, "Single regression coefficients", rlist)
  writeData(wb, "Multiple regression fit", fitsum)
  writeData(wb, "Top Indicators", topinds)
  
  saveWorkbook(wb, file = paste(userInput, " - eCSI Regressions", ".xlsx", sep = ""))
