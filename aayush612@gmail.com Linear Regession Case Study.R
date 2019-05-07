#----Getting the data for analysis--------------

setwd("C:/Users/AAYUSH GARG/Downloads/BAnalytics/Final Case Study 2 - Linear Regression")
require(car)
require(openxlsx)
mydata<-inp
mydata<- openxlsx::read.xlsx("Linear Regression Case.xlsx", sheet=1)
inp<-mydata
#contains indexs of categorical variables( used a Java code to find out the
# indexs of the categorical variable present in the xlsx sheet)

num<-c(F,T,T,T,F,T,T,F,T,T,T,T,T,T,F,F,T,F,F,F,F,F,T,T,T,F,T,F,F,F,F,F,F,F,F,F,T,T,T,T,T,
       T,T,F,T,T,T,T,T,F,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,T,F,F,F,F,T,
       T,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,F,F,F,F,T,T,T,T,T,T,T,T,T,T,F,T,T,T,
       T,T,T,T,T,T,T,T,T)

require(magrittr)

#----Stats Descriptive---------------------------

myStats<- function(x,nums=TRUE){
      if(nums==TRUE){
            quan<- quantile(x,na.rm=T,probs = c(0.05,0.10,0.25,0.5,0.75,0.90,0.95))
            nmiss<- sum(is.na(x))
            x<-x[!is.na(x)]
            stddev<- sd(x)
            LL<- mean(x)-3*stddev
            UL<- mean(x)+3*stddev
            outlier_x<- x> UL | x<LL
            pmiss<- nmiss*100/(nmiss+length(x))
            return(c(quan=quan,nmiss=nmiss,pmiss=pmiss,stddev=stddev,
                     L=length(x),outliers=sum(outlier_x)))
      }
      else{
            nmiss=sum(is.na(x))
            tab<- table(x)
            mode<- names(tab[tab==max(tab)])[1]
            pmis<- nmiss*100/(nmiss+length(x))
            return(c(nmiss=nmiss,mode=mode,pmiss=pmis,L=length(x)))
      }
}
h<- function(x,var=1,names="x"){ #plots hist if var=1 breaks=1000 default else specify var
      if(var==1){
            breaks=1000
      }
      else
            breaks=var
      hist(x,breaks = breaks,main  = names)
}

mydata$custid<- NULL #removing cust id column
num=num[-1]       #updating the pointer of index
num_stats<- t(data.frame(apply(mydata[,(!num)],2,myStats,nums=T)))
cat_stats<- t(data.frame(apply(mydata[,num],2,myStats,nums=F)))
num_stats %>% View
cat_stats %>% View

# write.csv(num_stats,)
# Marking al the colum with missing continous values
miss_dat<-c(18,19,21,49,87,88,91,93,96,98,106,108,101,102,103)
t(data.frame(apply(mydata[,miss_dat],2,myStats))) %>% View
#Deleting columns with large missing data, threshold 28%
mydata[,c(91,93,96,98,106,108,101,103)]<- NULL
num<-num[-c(91,93,96,98,106,108,101,103)] #removing the index of deleted variables
myd<-mydata# backup with rempved contin variables

#fixing missing value of categorical variable "townsize" with the mode value
mode_townsize<-table(mydata[,2]) %>% max
mydata$townsize[is.na(mydata$townsize)]<- mode_townsize
      
#----Correcting the outliers----

write.csv(num_stats,"Stats.csv")
#="mydata$"& $A2 &"= ifelse(mydata$" & $A2 & "<" & $B2 & ","& $B2 & ", ifelse(mydata$" & $A2 & ">"& $H2 &","&  $H2 & "," & "mydata$"& $A2& "))"

#Continous Outlier Treatment capping at 5% and 95% interval

mydata$age= ifelse(mydata$age<20,20, ifelse(mydata$age>76,76,mydata$age))
mydata$ed= ifelse(mydata$ed<9,9, ifelse(mydata$ed>20,20,mydata$ed))
mydata$income= ifelse(mydata$income<13,13, ifelse(mydata$income>147,147,mydata$income))
mydata$lninc= ifelse(mydata$lninc<2.56494935746154,2.56494935746154, ifelse(mydata$lninc>4.99043258677874,4.99043258677874,mydata$lninc))
mydata$debtinc= ifelse(mydata$debtinc<1.9,1.9, ifelse(mydata$debtinc>22.2,22.2,mydata$debtinc))
mydata$creddebt= ifelse(mydata$creddebt<0.101088,0.101088, ifelse(mydata$creddebt>6.3730104,6.3730104,mydata$creddebt))
mydata$lncreddebt= ifelse(mydata$lncreddebt< -2.29160361221836,-2.29160361221836, ifelse(mydata$lncreddebt>1.85229733265072,1.85229733265072,mydata$lncreddebt))
mydata$othdebt= ifelse(mydata$othdebt<0.2876923,0.2876923, ifelse(mydata$othdebt>11.8159808,11.8159808,mydata$othdebt))
mydata$lnothdebt= ifelse(mydata$lnothdebt< -1.24348344030873,-1.24348344030873, ifelse(mydata$lnothdebt>2.46958637465373,2.46958637465373,mydata$lnothdebt))
mydata$spoused= ifelse(mydata$spoused< -1,-1, ifelse(mydata$spoused>18,18,mydata$spoused))
mydata$reside= ifelse(mydata$reside<1,1, ifelse(mydata$reside>5,5,mydata$reside))
mydata$carvalue= ifelse(mydata$carvalue< -1,-1, ifelse(mydata$carvalue>72,72,mydata$carvalue))
mydata$commutetime= ifelse(mydata$commutetime<16,16, ifelse(mydata$commutetime>35,35,mydata$commutetime))
mydata$carditems= ifelse(mydata$carditems<5,5, ifelse(mydata$carditems>16,16,mydata$carditems))
mydata$cardspent= ifelse(mydata$cardspent<91.3045,91.3045, ifelse(mydata$cardspent>782.3155,782.3155,mydata$cardspent))
mydata$card2items= ifelse(mydata$card2items<1,1, ifelse(mydata$card2items>9,9,mydata$card2items))
mydata$card2spent= ifelse(mydata$card2spent<14.8195,14.8195, ifelse(mydata$card2spent>419.447,419.447,mydata$card2spent))
mydata$tenure= ifelse(mydata$tenure<4,4, ifelse(mydata$tenure>72,72,mydata$tenure))
mydata$longmon= ifelse(mydata$longmon<2.9,2.9, ifelse(mydata$longmon>36.7575,36.7575,mydata$longmon))
mydata$lnlongmon= ifelse(mydata$lnlongmon<1.06471073699243,1.06471073699243, ifelse(mydata$lnlongmon>3.60434189192823,3.60434189192823,mydata$lnlongmon))
mydata$longten= ifelse(mydata$longten<12.62,12.62, ifelse(mydata$longten>2567.65,2567.65,mydata$longten))
mydata$lnlongten= ifelse(mydata$lnlongten<2.53527150100047,2.53527150100047, ifelse(mydata$lnlongten>7.85074476077837,7.85074476077837,mydata$lnlongten))
mydata$tollmon= ifelse(mydata$tollmon<0,0, ifelse(mydata$tollmon>43.5,43.5,mydata$tollmon))
mydata$tollten= ifelse(mydata$tollten<0,0, ifelse(mydata$tollten>2620.2125,2620.2125,mydata$tollten))
mydata$equipmon= ifelse(mydata$equipmon<0,0, ifelse(mydata$equipmon>49.0525,49.0525,mydata$equipmon))
mydata$equipten= ifelse(mydata$equipten<0,0, ifelse(mydata$equipten>2600.99,2600.99,mydata$equipten))
mydata$cardmon= ifelse(mydata$cardmon<0,0, ifelse(mydata$cardmon>42,42,mydata$cardmon))
mydata$cardten= ifelse(mydata$cardten<0,0, ifelse(mydata$cardten>2455.75,2455.75,mydata$cardten))
mydata$wiremon= ifelse(mydata$wiremon<0,0, ifelse(mydata$wiremon>51.305,51.305,mydata$wiremon))
mydata$wireten= ifelse(mydata$wireten<0,0, ifelse(mydata$wireten>2687.9225,2687.9225,mydata$wireten))
mydata$hourstv= ifelse(mydata$hourstv<12,12, ifelse(mydata$hourstv>28,28,mydata$hourstv))

#imputing missing values
require(Hmisc)
mydata[,!num]<-data.frame(apply(mydata[,!num],2,function(x){ x<- impute(x)}))

#----Completed Outlier Treatment and imputing missing variables----#

indnum<-c(1:length(num))
indnum<-indnum[!num]

#Observing the histograms plots in 2x4 matrix
par(mfrow=c(2,4)) # setting to display 8 plots per page
for(i in seq(1,sum(!num))){
      h(mydata[,indnum[i]],names=names(mydata)[indnum[i]])
}

#----Dependent Variable Creation and Modification ----

mydata$total_spent<- mydata$cardspent+ mydata$card2spent
mydata$ln_spent<-log(mydata$total_spent)
num<-c(num,F,F)#adding total_spent , ln_spent
#h(mydata[,"ln_spent"])

# summary(aov(ln_spent~pets,data=mydata))[4]

indcat<-seq(1,ncol(mydata))
indcat<-indcat[num]


#----Performing ANOVA Test on Categorical Variables to test its signnificance----
anovadf<-data.frame(variable="variable",F_value=1,Pvalue=1)[-1,]
for(i in c(1:length(indcat))){
      anoF=anova(lm(paste("ln_spent~mydata$",names(mydata)[indcat[i]],sep=""),data = mydata))[1,4]
      ano=anova(lm(paste("ln_spent~mydata$",names(mydata)[indcat[i]],sep=""),data = mydata))[1,5]
      nm=names(mydata)[indcat[i]]
      
      #constructed a table with anova values to filter significant variables
      #for model
      anovadf<-rbind(anovadf,data.frame(nm,anoF,ano))
      cat(nm,anoF,ano,"\n")
}
mydd<-mydata #backup point 2

#converted categorical vriabls into factors
mydata[,num]<-data.frame(apply(mydata[,num],2,factor))

#----Finding correlations between continous vars/Multicollinearity check----#
cordat<-mydata[,!num]
cordat<-cordat[,-c(12:19,6,8,22,23,24,25,27,29,31,32)]#removing variables whose
            #transformed variable exists in dataset such as wiremon(lnwiremon)
corrm<- cor(cordat)

require(corrplot)
# corrplot(corrm) #Observing Correlations
require(psych)
require(GPArotation)
par(mfrow=c(1,1))
scree(corrm,factors = T, pc=T)
eig<- eigen(corrm)
frame<- data.frame(val=eig$values,cumsumval=cumsum(eig$values),
                   percent=cumsum(eig$values)/sum(eig$values))
FA<- fa(corrm,nfactors = 9,rotate = "varimax", fm="ml")
FA<- fa.sort(FA)
Loadings<-data.frame(FA$loadings[,])
# View(Loadings)
write.csv(Loadings,"D:/LoadingLr.csv")

#----Bivariate Analysis----#
cat("The following shows correlations with the dependent variable ln_spent")
cor.mat<- cor(cordat,cordat$ln_spent)
cor.mat

#----Creation of Dummy Variables of significant predictors----

mydata$internet3<- ifelse(mydata$internet==3,1,0)
mydata$internet1<- ifelse(mydata$internet==1,1,0)
mydata$internet2<- ifelse(mydata$internet==2,1,0)
mydata$internet4<- ifelse(mydata$internet==4,1,0)
mydata$reason2<- ifelse(mydata$reason==2,1,0)
mydata$reason3<- ifelse(mydata$reason==3,1,0)
mydata$reason4<- ifelse(mydata$reason==4,1,0)
mydata$reason9<- ifelse(mydata$reason==9,1,0)
mydata$region2<-ifelse(mydata$region==2,1,0)
mydata$region3<-ifelse(mydata$region==3,1,0)
mydata$region4<-ifelse(mydata$region==4,1,0)
mydata$region5<-ifelse(mydata$region==5,1,0)

#----Separting Train and Test datasets----

set.seed(1)
ind<- sample(c(1:nrow(mydata)),size=floor(0.7*nrow(mydata)))
train<-mydata[ind,]
test<-mydata[-ind,]

#----Model Fitting-----------------
#Selecting categorical variables with less P value
sigcat<-anovadf[anovadf[,"ano"]<(5*10**(-3)),"nm"]
sigcat
fit<-lm(ln_spent~age+ed+income+lninc+debtinc+lncreddebt+
              reside+pets+carvalue+commutetime+tenure+
              lnlongmon+lnothdebt+spoused+wiremon+wireten+hourstv+
              lnlongten+tollmon+tollten+equipmon+equipten+cardmon+cardten+
              gender+agecat+edcat+jobcat+employ+empcat+retire+inccat+jobsat+
              homeown+address+addresscat+carown+carcatvalue+reason+vote+card+
              cardtenurecat+card2+card2tenurecat+wireless+internet+owntv+ownvcr+
              owndvd+owncd+ownpda+ownfax+response_03,data = train)

require(MASS)              
# fit2<- stepAIC(fit,direction="both")

fit3<-lm(formula = ln_spent ~ age + ed +lninc +
               gender + reason2+reason9 + hourstv+ lncreddebt+
               card  + card2 + region2+region4+region5+ internet3
               , data = train) 

#Using anova to compare 2 models by adding extra variables in fit4 to check 
#its significance

fit4<-lm(formula = ln_spent ~ age +lninc+ 
               gender+reason2+reason9+
               card+card2+region2+region4+region5+internet3+ cardfee
               , data = train) 

anova(fit3,fit4) #to check significance of the added predictor for ex. cardfee is somewhat significant
fit3 %>% summary
vif(fit4)

#----Handling influential observations and fitting the model again----

require(dplyr)
train$Cd<- cooks.distance(fit4)
train2<-subset(train, Cd< (4/3500))

fit5<-lm(formula = ln_spent ~ age +lninc+ 
               gender+reason2+reason9+
               card+card2+region2+region4+region5+internet3+ cardfee
         , data = train2) 
fit5 %>% summary

#----SCORING  USING PREDICTION-------------------

train2$pred_spend<- NULL
train2<-cbind(train2, pred_spend = exp(predict(fit3)))
# names(train)
t1<- transform(train2, APE = abs(pred_spend - total_spent)/total_spent)
mean(t1$APE)
# View(t1)
# 
test<-cbind(test, pred_spend=exp(predict(fit3,test)))
test<- transform(test, APE = abs(pred_spend - total_spent)/total_spent)
mean(test$APE)


#----Decile analysis of Development Daaset-----------------

require(sqldf)
# find the decile locations 
decLocations <- quantile(train2$pred_spend, probs = seq(0.1,0.9,by=0.1))

# use findInterval with -Inf and Inf as upper and lower bounds
train2$decile <- findInterval(train2$pred_spend,c(-Inf,decLocations, Inf))


train_D <- sqldf("select decile, count(decile) as count, avg(total_spent) as total_spent,   
                    avg(pred_spend) as pred_spend
                    from train2
                    group by decile
                    order by decile desc")

View(train_D)

#----DECILE ANALYSIS TESTING SAMPLE--------



# find the decile locations 
decLocations <- quantile(test$pred_spend, probs = seq(0.1,0.9,by=0.1))

# use findInterval with -Inf and Inf as upper and lower bounds
test$decile <- findInterval(test$pred_spend,c(-Inf,decLocations, Inf))

test_D <- sqldf("select decile, count(decile) as count, avg(total_spent) as total_spent,   
                 avg(pred_spend) as pred_spend
                 from test
                 group by decile
                 order by decile desc")

View(test_D)

#____END OF ANALYSIS_________####
