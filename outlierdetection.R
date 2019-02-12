
########################################################################
# Load libraries
########################################################################

suppressWarnings(suppressMessages(library(readxl,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(stringi,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(dplyr,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(XLConnect,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(gtools,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(tibble,quietly=TRUE))) 
suppressWarnings(suppressMessages(library(tictoc,quietly=TRUE))) 

########################################################################
# Setup
########################################################################

id_panel=c(1,3,5,6,7,8,10,
           11,12,13,14,15,
           17,19,20)                          # IDs of panelists
nick_panel=c("Crist","Sol","Agnès","Mercè",
             "Ester","Tere","AnnaB","AnaCr",
             "Bene","Hay","Jean","MJose",
             "Mont","Ruth","Yol")             # Names of panelists  
                                              # matching IDs

PLeader_1="Mercè"                             # Leaders of panel 1 and 2
PLeader_2="AnaCR"

########################################################################
# Outlier test functino
########################################################################

process_outliers=function(){
  
  # Choose file + random, set wd as current folder
  
  file=choose.files()
  path_=strsplit(file,"[\\\\]")[[1]]
  setwd(paste(path_[1:(length(path_)-1)],collapse="\\"))
  filename=last(path_)
  
  random=choose.files()
  tic()                                             # Start timer
  WB=loadWorkbook(file)                             # Start results WB
  a=readWorksheet(WB,sheet=getSheets(WB)[1])
  m_temp=read_xlsx(file)
  a$Marca.temporal=m_temp$`Marca temporal`          # Update timeframe
  a$Panelista=sapply(1:length(a$Panelista),function(x){
    i=which(id_panel==a$Panelista[x])
    a$Panelista[x]=nick_panel[i]
  })                                                # Change Panelist ID to name
  
  # Add panelist names
  # Add nº of panel (1 or 2). Yield error if no panel leader is found
  if (PLeader_1 %in% a$Panelista){
    Panel=rep(1,dim(a)[1])
  } else if (PLeader_2 %in% a$Panelista){
    Panel=rep(2,dim(a)[1])
  } else{
    Panel=rep(0,dim(a)[1])
    print(paste0("Error: ", PLeader_1, " and ", Pleader2, " are not found, don't know panel."))
  }
  a=add_column(a,Panel,.after="Sesión")
  
  # Turn characteristic columns to numeric
  ichar=6:(dim(a)[2]-1)                   
  a[,ichar]=sapply(a[,ichar], as.numeric) 
  
  # Turn codes into actual product names
  
  rand=loadWorkbook(random)
  n_sheets=length(getSheets(rand))
  codigos=NULL
  
  # Create dictionary (code -> name of product)
  for (i in 1:n_sheets){
    u=readWorksheet(rand,sheet=getSheets(rand)[i])
    u=unlist(u,use.names=F)
    codigos=c(codigos,u[grep(":",u)])
    sesplit=strsplit(u[grep(":",u)],"\\:")
    ses=which(a$Código %in%unlist(sesplit))
    a$Sesión[ses]=i
  }
  
  # Compare codes to dictionary containing code -> name
  aux=unlist(stri_split_fixed(str=codigos,pattern=": ",n=2))
  dim(aux)=c(2,length(aux)/2)
  aux[1,]=toupper(aux[1,])
  
  # Check if any code is wrong (if it appears less than 3 times)
  # Can be immediately changed
  
  numcods=a$Código
  if (length(which(table(numcods)<3))>0){
    cat("There are some wrong codes\n")
    cat("Number of repetitions per code: \n")
    print(table(numcods))
    a$Código=fix(numcods)
  }
  
  cods=sapply(1:dim(a)[1],function(x){
    i=which(aux[1,]==toupper(a$Código[x]))
    a$Código[x]=aux[2,i]
  })
  
  a$Código=as.factor(cods)
  an=a
  aorig=an
  aorig=aorig[with(aorig,order(Código,Panelista,Sesión)),]
  levs=levels(a$Código)
  
  
  # Mean and SD per product
  m=a[,c(3,ichar)]%>%group_by(Código)%>%summarise_all(funs(mean),na.rm=TRUE)
  sd=a[,c(3,ichar)]%>%group_by(Código)%>%summarise_all(funs(sd),na.rm=TRUE)
  
  # Filter +-2SD
  count_oversd=0
  count_overmean=0
  for(i in 1:nrow(a)){
    prod=which(levs==a[i,3][[1]])
    sdp=sd[prod,-1]
    mp=m[prod,-1]
    
    overmean=which(abs(a[i,ichar]-mp)>1)
    oversd=which(abs((a[i,ichar]-mp))>2*sdp)
    
    count_overmean=count_overmean+length(overmean)
    count_oversd=count_oversd+length(oversd)
    
    remove=oversd
    an[i,(5+remove)]=NA
  }
  # Position of NAs in original sheet
  
  ######### Up to this point, 'an' is == to 'a' with removed +-2devs.
  
  an=an[with(an,order(Código,Panelista,Sesión)),]
  orig_na=which(is.na(an),arr.ind=TRUE)
  # New means and SDs
  
  mn=an[,c(3,ichar)]%>%group_by(Código)%>%summarise_all(funs(mean),na.rm=TRUE)
  sdn=an[,c(3,ichar)]%>%group_by(Código)%>%summarise_all(funs(sd),na.rm=TRUE)
  mn$Código=as.character.factor(mn$Código)
  sdn$Código=as.character.factor(sdn$Código)
  # Create new matrix by product, including mean and SD
  at=an[FALSE,]
  for (x in 1:length(levs)){
    i=which(an$Código==levs[x])
    a=mn[x,]
    a$Código="Mean"
    a[,-1]=round(a[,-1],2)
    b=sdn[x,]
    b$Código="SD"
    b[,-1]=round(b[,-1],2)
    at=smartbind(at,an[i,],a,b)
  }
  
  # Add number of NAs per column
  na_count <-sapply(at[,ichar], function(y) sum(length(which(is.na(y)))))
  n=as.data.frame(cbind(na_count,na_count))
  
  na_count$Código="Errors deleted"
  
  at=smartbind(at,t(na_count))
  at[,ichar]=sapply(at[,ichar], as.numeric)
  
  # Create new sheet with processed matrix
  
  createSheet(WB,"processed")
  writeWorksheet(WB,at,sheet="processed")
  writeWorksheet(WB,aorig,sheet=getSheets(WB)[1])
  
  # Style NAs painting them blue
  
  style <- createCellStyle(WB,name="NA")
  setFillPattern(style,fill=XLC$"FILL.SOLID_FOREGROUND")
  setFillForegroundColor(style,color=XLC$COLOR.LIGHT_TURQUOISE)
  col=which(is.na(at),arr.ind=TRUE)
  col=col[-which(col[,2]<6 |col[,2]==nrow(at)),]
  rowIndex=col[,1]+1
  colIndex=col[,2]
  setCellStyle(WB,sheet="processed",row=rowIndex,
               col=colIndex,cellstyle=style)
  rowIndex2=orig_na[,1]+1
  colIndex2=orig_na[,2]
  setCellStyle(WB,sheet=getSheets(WB)[1],row=rowIndex2,
               col=colIndex2,cellstyle=style)
  
  # Adapt width to length of characters
  
  width_vec <- apply(at, 2, function(x) max(nchar(as.character(x)) + 2, na.rm = TRUE))
  width_vec_header <- nchar(colnames(at))  + 2
  max_vec_header <- pmax(width_vec, width_vec_header)
  setColumnWidth(WB, sheet="processed", column = 1:ncol(at), width =max_vec_header*256 )
  
  exit_aux=strsplit(filename,"\\.")[[1]]
  exitfile=paste(exit_aux[1],"_processed.",exit_aux[2],sep="")
  saveWorkbook(WB,exitfile)
}
process_outliers()                            # Main call to function
toc()                                         # End of timer

########### AUTOMATIC OUTLIER DETECTION

# 1) Run process_outliers code

# 2) File search prompt. Select original excel sheet with Panel data.

# 3) Second file search prompt. Select randomization excel sheet.

# 4) Function returns processed excel to same folder where raw exists.


