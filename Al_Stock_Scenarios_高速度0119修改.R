#Al Calculation
install.packages("openxlsx")
library(openxlsx)

setwd("D:/坚果云同步文件夹/RBOOK/AL-高速度")

#导入历史人均存量数据
HSPC= read.xlsx("Al Data_1950-2100.xlsx", sheet="HSPC")
IAI=HSPC[1:67,1:9]
WQC=HSPC[1:66,10:18]
PaS= read.xlsx("Al Data_1950-2100.xlsx", sheet="PaS")
JPN=PaS[7:74,1:9]
colnames(JPN)=JPN[1,]
JPN=JPN[-1,]
NAM=PaS[7:74,10:18]
colnames(NAM)=NAM[1,]
NAM=NAM[-1,]
EU=PaS[7:74,19:27]
colnames(EU)=EU[1,]
EU=EU[-1,]

#逻辑斯谛拟合并返回参数
#北美数据
##第1条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,2])
NAMf1 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf1)
a1=P$parameters[1,1]
xc1=P$parameters[2,1]
k1=1/(P$parameters[3,1])
##第2条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,3])
NAMf2 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf2)
a2=P$parameters[1,1]
xc2=P$parameters[2,1]
k2=1/(P$parameters[3,1])
##第3条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,4])
NAMf3 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf3)
a3=P$parameters[1,1]
xc3=P$parameters[2,1]
k3=1/(P$parameters[3,1])
##第4条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,5])
NAMf4 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf4)
a4=P$parameters[1,1]
xc4=P$parameters[2,1]
k4=1/(P$parameters[3,1])
##第5条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,6])
NAMf5 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf5)
a5=P$parameters[1,1]
xc5=P$parameters[2,1]
k5=1/(P$parameters[3,1])
##第6条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,7])
NAMf6 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf6)
a6=P$parameters[1,1]
xc6=P$parameters[2,1]
k6=1/(P$parameters[3,1])
##第7条
X=as.numeric(NAM[1:67,1])
Y=as.numeric(NAM[1:67,8])
NAMf7 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(NAMf7)
a7=P$parameters[1,1]
xc7=P$parameters[2,1]
k7=1/(P$parameters[3,1])

PM_NAM=data.frame(a=c(a1,a2,a3,a4,a5,a6,a7), xc=c(xc1,xc2,xc3,xc4,xc5,xc6,xc7), k=c(k1,k2,k3,k4,k5,k6,k7))
write.csv(PM_NAM,"OUT/Parameter_NAM.csv")

##日本数据
##第1条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,2])
JPNf1 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf1)
a1=P$parameters[1,1]
xc1=P$parameters[2,1]
k1=1/(P$parameters[3,1])
##第2条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,3])
JPNf2 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf2)
a2=P$parameters[1,1]
xc2=P$parameters[2,1]
k2=1/(P$parameters[3,1])
##第3条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,4])
JPNf3 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf3)
a3=P$parameters[1,1]
xc3=P$parameters[2,1]
k3=1/(P$parameters[3,1])
##第4条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,5])
JPNf4 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf4)
a4=P$parameters[1,1]
xc4=P$parameters[2,1]
k4=1/(P$parameters[3,1])
##第5条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,6])
JPNf5 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf5)
a5=P$parameters[1,1]
xc5=P$parameters[2,1]
k5=1/(P$parameters[3,1])
##第6条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,7])
JPNf6 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf6)
a6=P$parameters[1,1]
xc6=P$parameters[2,1]
k6=1/(P$parameters[3,1])
##第7条
X=as.numeric(JPN[1:67,1])
Y=as.numeric(JPN[1:67,8])
JPNf7 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(JPNf7)
a7=P$parameters[1,1]
xc7=P$parameters[2,1]
k7=1/(P$parameters[3,1])

PM_JPN=data.frame(a=c(a1,a2,a3,a4,a5,a6,a7), xc=c(xc1,xc2,xc3,xc4,xc5,xc6,xc7), k=c(k1,k2,k3,k4,k5,k6,k7))
write.csv(PM_JPN,"OUT/Parameter_JPN.csv")

##欧洲数据
##第1条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,2])
EUf1 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf1)
a1=P$parameters[1,1]
xc1=P$parameters[2,1]
k1=1/(P$parameters[3,1])
##第2条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,3])
EUf2 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf2)
a2=P$parameters[1,1]
xc2=P$parameters[2,1]
k2=1/(P$parameters[3,1])
##第3条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,4])
EUf3 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf3)
a3=P$parameters[1,1]
xc3=P$parameters[2,1]
k3=1/(P$parameters[3,1])
##第4条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,5])
EUf4 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf4)
a4=P$parameters[1,1]
xc4=P$parameters[2,1]
k4=1/(P$parameters[3,1])
##第5条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,6])
EUf5 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf5)
a5=P$parameters[1,1]
xc5=P$parameters[2,1]
k5=1/(P$parameters[3,1])
##第6条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,7])
EUf6 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf6)
a6=P$parameters[1,1]
xc6=P$parameters[2,1]
k6=1/(P$parameters[3,1])
##第7条
X=as.numeric(EU[1:67,1])
Y=as.numeric(EU[1:67,8])
EUf7 = nls(Y ~ SSlogis(X, a,xc,k))
P = summary(EUf7)
a7=P$parameters[1,1]
xc7=P$parameters[2,1]
k7=1/(P$parameters[3,1])

PM_EU=data.frame(a=c(a1,a2,a3,a4,a5,a6,a7), xc=c(xc1,xc2,xc3,xc4,xc5,xc6,xc7), k=c(k1,k2,k3,k4,k5,k6,k7))
write.csv(PM_EU,"OUT/Parameter_EU.csv")


#拟合未来人均存量曲线，12种情况
PaS_OUT=read.xlsx("Al Data_1950-2100.xlsx", sheet="PaS_OUT")
Peak=PaS_OUT[1:3,1:9]
Speed=PaS_OUT[1:3,10:18]
#PM12-1HPHS高高组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
HPHS1=nls(Y ~ (Peak[1,2]/(1+exp((xcc-X)/Speed[1,2]))),start=list(xcc=2000))
P=summary(HPHS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
HPHS2=nls(Y ~ (Peak[1,3]/(1+exp((xcc-X)/Speed[1,3]))),start=list(xcc=2000))
P=summary(HPHS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
HPHS3=nls(Y ~ (Peak[1,4]/(1+exp((xcc-X)/Speed[1,4]))),start=list(xcc=2000))
P=summary(HPHS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
HPHS4=nls(Y ~ (Peak[1,5]/(1+exp((xcc-X)/Speed[1,5]))),start=list(xcc=2000))
P=summary(HPHS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
HPHS5=nls(Y ~ (Peak[1,6]/(1+exp((xcc-X)/Speed[1,6]))),start=list(xcc=2000))
P=summary(HPHS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
HPHS6=nls(Y ~ (Peak[1,7]/(1+exp((xcc-X)/Speed[1,7]))),start=list(xcc=2000))
P=summary(HPHS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
HPHS7=nls(Y ~ (Peak[1,8]/(1+exp((xcc-X)/Speed[1,8]))),start=list(xcc=2000))
P=summary(HPHS7)
xcc7=P$parameters[1,1]
HPHS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-2HPMS高中组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
HPMS1=nls(Y ~ (Peak[1,2]/(1+exp((xcc-X)/Speed[2,2]))),start=list(xcc=2000))
P=summary(HPMS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
HPMS2=nls(Y ~ (Peak[1,3]/(1+exp((xcc-X)/Speed[2,3]))),start=list(xcc=2000))
P=summary(HPMS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
HPMS3=nls(Y ~ (Peak[1,4]/(1+exp((xcc-X)/Speed[2,4]))),start=list(xcc=2000))
P=summary(HPMS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
HPMS4=nls(Y ~ (Peak[1,5]/(1+exp((xcc-X)/Speed[2,5]))),start=list(xcc=2000))
P=summary(HPMS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
HPMS5=nls(Y ~ (Peak[1,6]/(1+exp((xcc-X)/Speed[2,6]))),start=list(xcc=2000))
P=summary(HPMS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
HPMS6=nls(Y ~ (Peak[1,7]/(1+exp((xcc-X)/Speed[2,7]))),start=list(xcc=2000))
P=summary(HPMS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
HPMS7=nls(Y ~ (Peak[1,8]/(1+exp((xcc-X)/Speed[2,8]))),start=list(xcc=2000))
P=summary(HPMS7)
xcc7=P$parameters[1,1]
HPMS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-3HPLS高低组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
HPLS1=nls(Y ~ (Peak[1,2]/(1+exp((xcc-X)/Speed[3,2]))),start=list(xcc=2000))
P=summary(HPLS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
HPLS2=nls(Y ~ (Peak[1,3]/(1+exp((xcc-X)/Speed[3,3]))),start=list(xcc=2000))
P=summary(HPLS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
HPLS3=nls(Y ~ (Peak[1,4]/(1+exp((xcc-X)/Speed[3,4]))),start=list(xcc=2000))
P=summary(HPLS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
HPLS4=nls(Y ~ (Peak[1,5]/(1+exp((xcc-X)/Speed[3,5]))),start=list(xcc=2000))
P=summary(HPLS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
HPLS5=nls(Y ~ (Peak[1,6]/(1+exp((xcc-X)/Speed[3,6]))),start=list(xcc=2000))
P=summary(HPLS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
HPLS6=nls(Y ~ (Peak[1,7]/(1+exp((xcc-X)/Speed[3,7]))),start=list(xcc=2000))
P=summary(HPLS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
HPLS7=nls(Y ~ (Peak[1,8]/(1+exp((xcc-X)/Speed[3,8]))),start=list(xcc=2000))
P=summary(HPLS7)
xcc7=P$parameters[1,1]
HPLS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-4HPNS高自然组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
HPNS1=nls(Y ~ (Peak[1,2]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS1)
xcc1=P$parameters[1,1]
k1=P$parameters[2,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
HPNS2=nls(Y ~ (Peak[1,3]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS2)
xcc2=P$parameters[1,1]
k2=P$parameters[2,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
HPNS3=nls(Y ~ (Peak[1,4]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS3)
xcc3=P$parameters[1,1]
k3=P$parameters[2,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
HPNS4=nls(Y ~ (Peak[1,5]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS4)
xcc4=P$parameters[1,1]
k4=P$parameters[2,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
HPNS5=nls(Y ~ (Peak[1,6]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS5)
xcc5=P$parameters[1,1]
k5=P$parameters[2,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
HPNS6=nls(Y ~ (Peak[1,7]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS6)
xcc6=P$parameters[1,1]
k6=P$parameters[2,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
HPNS7=nls(Y ~ (Peak[1,8]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(HPNS7)
xcc7=P$parameters[1,1]
k7=P$parameters[2,1]
HPNS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)
HPNS_k=c(k1,k2,k3,k4,k5,k6,k7)


#PM12-5MPHS中高组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
MPHS1=nls(Y ~ (Peak[2,2]/(1+exp((xcc-X)/Speed[1,2]))),start=list(xcc=2000))
P=summary(MPHS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
MPHS2=nls(Y ~ (Peak[2,3]/(1+exp((xcc-X)/Speed[1,3]))),start=list(xcc=2000))
P=summary(MPHS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
MPHS3=nls(Y ~ (Peak[2,4]/(1+exp((xcc-X)/Speed[1,4]))),start=list(xcc=2000))
P=summary(MPHS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
MPHS4=nls(Y ~ (Peak[2,5]/(1+exp((xcc-X)/Speed[1,5]))),start=list(xcc=2000))
P=summary(MPHS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
MPHS5=nls(Y ~ (Peak[2,6]/(1+exp((xcc-X)/Speed[1,6]))),start=list(xcc=2000))
P=summary(MPHS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
MPHS6=nls(Y ~ (Peak[2,7]/(1+exp((xcc-X)/Speed[1,7]))),start=list(xcc=2000))
P=summary(MPHS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
MPHS7=nls(Y ~ (Peak[2,8]/(1+exp((xcc-X)/Speed[1,8]))),start=list(xcc=2000))
P=summary(MPHS7)
xcc7=P$parameters[1,1]
MPHS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-6MPMS中中组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
MPMS1=nls(Y ~ (Peak[2,2]/(1+exp((xcc-X)/Speed[2,2]))),start=list(xcc=2000))
P=summary(MPMS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
MPMS2=nls(Y ~ (Peak[2,3]/(1+exp((xcc-X)/Speed[2,3]))),start=list(xcc=2000))
P=summary(MPMS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
MPMS3=nls(Y ~ (Peak[2,4]/(1+exp((xcc-X)/Speed[2,4]))),start=list(xcc=2000))
P=summary(MPMS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
MPMS4=nls(Y ~ (Peak[2,5]/(1+exp((xcc-X)/Speed[2,5]))),start=list(xcc=2000))
P=summary(MPMS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
MPMS5=nls(Y ~ (Peak[2,6]/(1+exp((xcc-X)/Speed[2,6]))),start=list(xcc=2000))
P=summary(MPMS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
MPMS6=nls(Y ~ (Peak[2,7]/(1+exp((xcc-X)/Speed[2,7]))),start=list(xcc=2000))
P=summary(MPMS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
MPMS7=nls(Y ~ (Peak[2,8]/(1+exp((xcc-X)/Speed[2,8]))),start=list(xcc=2000))
P=summary(MPMS7)
xcc7=P$parameters[1,1]
MPMS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-7MPLS中低组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
MPLS1=nls(Y ~ (Peak[2,2]/(1+exp((xcc-X)/Speed[3,2]))),start=list(xcc=2000))
P=summary(MPLS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
MPLS2=nls(Y ~ (Peak[2,3]/(1+exp((xcc-X)/Speed[3,3]))),start=list(xcc=2000))
P=summary(MPLS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
MPLS3=nls(Y ~ (Peak[2,4]/(1+exp((xcc-X)/Speed[3,4]))),start=list(xcc=2000))
P=summary(MPLS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
MPLS4=nls(Y ~ (Peak[2,5]/(1+exp((xcc-X)/Speed[3,5]))),start=list(xcc=2000))
P=summary(MPLS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
MPLS5=nls(Y ~ (Peak[2,6]/(1+exp((xcc-X)/Speed[3,6]))),start=list(xcc=2000))
P=summary(MPLS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
MPLS6=nls(Y ~ (Peak[2,7]/(1+exp((xcc-X)/Speed[3,7]))),start=list(xcc=2000))
P=summary(MPLS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
MPLS7=nls(Y ~ (Peak[2,8]/(1+exp((xcc-X)/Speed[3,8]))),start=list(xcc=2000))
P=summary(MPLS7)
xcc7=P$parameters[1,1]
MPLS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-8MPNS中自然组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
MPNS1=nls(Y ~ (Peak[2,2]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS1)
xcc1=P$parameters[1,1]
k1=P$parameters[2,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
MPNS2=nls(Y ~ (Peak[2,3]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS2)
xcc2=P$parameters[1,1]
k2=P$parameters[2,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
MPNS3=nls(Y ~ (Peak[2,4]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS3)
xcc3=P$parameters[1,1]
k3=P$parameters[2,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
MPNS4=nls(Y ~ (Peak[2,5]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS4)
xcc4=P$parameters[1,1]
k4=P$parameters[2,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
MPNS5=nls(Y ~ (Peak[2,6]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS5)
xcc5=P$parameters[1,1]
k5=P$parameters[2,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
MPNS6=nls(Y ~ (Peak[2,7]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS6)
xcc6=P$parameters[1,1]
k6=P$parameters[2,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
MPNS7=nls(Y ~ (Peak[2,8]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(MPNS7)
xcc7=P$parameters[1,1]
k7=P$parameters[2,1]
MPNS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)
MPNS_k=c(k1,k2,k3,k4,k5,k6,k7)

#PM12-9LPHS低高组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
LPHS1=nls(Y ~ (Peak[3,2]/(1+exp((xcc-X)/Speed[1,2]))),start=list(xcc=2000))
P=summary(LPHS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
LPHS2=nls(Y ~ (Peak[3,3]/(1+exp((xcc-X)/Speed[1,3]))),start=list(xcc=2000))
P=summary(LPHS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
LPHS3=nls(Y ~ (Peak[3,4]/(1+exp((xcc-X)/Speed[1,4]))),start=list(xcc=2000))
P=summary(LPHS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
LPHS4=nls(Y ~ (Peak[3,5]/(1+exp((xcc-X)/Speed[1,5]))),start=list(xcc=2000))
P=summary(LPHS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
LPHS5=nls(Y ~ (Peak[3,6]/(1+exp((xcc-X)/Speed[1,6]))),start=list(xcc=2000))
P=summary(LPHS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
LPHS6=nls(Y ~ (Peak[3,7]/(1+exp((xcc-X)/Speed[1,7]))),start=list(xcc=2000))
P=summary(LPHS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
LPHS7=nls(Y ~ (Peak[3,8]/(1+exp((xcc-X)/Speed[1,8]))),start=list(xcc=2000))
P=summary(LPHS7)
xcc7=P$parameters[1,1]
LPHS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-10LPMS低中组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
LPMS1=nls(Y ~ (Peak[3,2]/(1+exp((xcc-X)/Speed[2,2]))),start=list(xcc=2000))
P=summary(LPMS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
LPMS2=nls(Y ~ (Peak[3,3]/(1+exp((xcc-X)/Speed[2,3]))),start=list(xcc=2000))
P=summary(LPMS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
LPMS3=nls(Y ~ (Peak[3,4]/(1+exp((xcc-X)/Speed[2,4]))),start=list(xcc=2000))
P=summary(LPMS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
LPMS4=nls(Y ~ (Peak[3,5]/(1+exp((xcc-X)/Speed[2,5]))),start=list(xcc=2000))
P=summary(LPMS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
LPMS5=nls(Y ~ (Peak[3,6]/(1+exp((xcc-X)/Speed[2,6]))),start=list(xcc=2000))
P=summary(LPMS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
LPMS6=nls(Y ~ (Peak[3,7]/(1+exp((xcc-X)/Speed[2,7]))),start=list(xcc=2000))
P=summary(LPMS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
LPMS7=nls(Y ~ (Peak[3,8]/(1+exp((xcc-X)/Speed[2,8]))),start=list(xcc=2000))
P=summary(LPMS7)
xcc7=P$parameters[1,1]
LPMS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-11LPLS低低组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
LPLS1=nls(Y ~ (Peak[3,2]/(1+exp((xcc-X)/Speed[3,2]))),start=list(xcc=2000))
P=summary(LPLS1)
xcc1=P$parameters[1,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
LPLS2=nls(Y ~ (Peak[3,3]/(1+exp((xcc-X)/Speed[3,3]))),start=list(xcc=2000))
P=summary(LPLS2)
xcc2=P$parameters[1,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
LPLS3=nls(Y ~ (Peak[3,4]/(1+exp((xcc-X)/Speed[3,4]))),start=list(xcc=2000))
P=summary(LPLS3)
xcc3=P$parameters[1,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
LPLS4=nls(Y ~ (Peak[3,5]/(1+exp((xcc-X)/Speed[3,5]))),start=list(xcc=2000))
P=summary(LPLS4)
xcc4=P$parameters[1,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
LPLS5=nls(Y ~ (Peak[3,6]/(1+exp((xcc-X)/Speed[3,6]))),start=list(xcc=2000))
P=summary(LPLS5)
xcc5=P$parameters[1,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
LPLS6=nls(Y ~ (Peak[3,7]/(1+exp((xcc-X)/Speed[3,7]))),start=list(xcc=2000))
P=summary(LPLS6)
xcc6=P$parameters[1,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
LPLS7=nls(Y ~ (Peak[3,8]/(1+exp((xcc-X)/Speed[3,8]))),start=list(xcc=2000))
P=summary(LPLS7)
xcc7=P$parameters[1,1]
LPLS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)

#PM12-12LPNS低自然组合
#第1条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,2])
LPNS1=nls(Y ~ (Peak[3,2]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS1)
xcc1=P$parameters[1,1]
k1=P$parameters[2,1]
#第2条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,3])
LPNS2=nls(Y ~ (Peak[3,3]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS2)
xcc2=P$parameters[1,1]
k2=P$parameters[2,1]
#第3条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,4])
LPNS3=nls(Y ~ (Peak[3,4]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS3)
xcc3=P$parameters[1,1]
k3=P$parameters[2,1]
#第4条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,5])
LPNS4=nls(Y ~ (Peak[3,5]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS4)
xcc4=P$parameters[1,1]
k4=P$parameters[2,1]
#第5条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,6])
LPNS5=nls(Y ~ (Peak[3,6]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS5)
xcc5=P$parameters[1,1]
k5=P$parameters[2,1]
#第6条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,7])
LPNS6=nls(Y ~ (Peak[3,7]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS6)
xcc6=P$parameters[1,1]
k6=P$parameters[2,1]
#第7条
X=as.numeric(IAI[1:67,1])
Y=as.numeric(IAI[1:67,8])
LPNS7=nls(Y ~ (Peak[3,8]/(1+exp((xcc-X)/k))),start=list(xcc=2000,k=10))
P=summary(LPNS7)
xcc7=P$parameters[1,1]
k7=P$parameters[2,1]
LPNS_xcc=c(xcc1,xcc2,xcc3,xcc4,xcc5,xcc6,xcc7)
LPNS_k=c(k1,k2,k3,k4,k5,k6,k7)

#将12种情况的参数输出到一个csv文件
FSPC1_PM=rbind(Peak[1,2:8],HPHS_xcc,Speed[1,2:8],
               Peak[1,2:8],HPMS_xcc,Speed[2,2:8],
               Peak[1,2:8],HPLS_xcc,Speed[3,2:8],
               Peak[1,2:8],HPNS_xcc,HPNS_k,
               Peak[2,2:8],MPHS_xcc,Speed[1,2:8],
               Peak[2,2:8],MPMS_xcc,Speed[2,2:8],
               Peak[2,2:8],MPLS_xcc,Speed[3,2:8],
               Peak[2,2:8],MPNS_xcc,MPNS_k,
               Peak[3,2:8],LPHS_xcc,Speed[1,2:8],
               Peak[3,2:8],LPMS_xcc,Speed[2,2:8],
               Peak[3,2:8],LPLS_xcc,Speed[3,2:8],
               Peak[3,2:8],LPNS_xcc,LPNS_k)
write.csv(FSPC1_PM,"OUT/Parameter_FSPC1.csv")

#开始计算未来人均存量曲线
#定义年份
FSPC_T=c(1950:2100)
#开始计算12种情况的人均存量曲线
#C12-1
HPHS_C1=FSPC1_PM[1,1]/(1+exp((FSPC1_PM[2,1]-FSPC_T)/FSPC1_PM[3,1]))
HPHS_C2=FSPC1_PM[1,2]/(1+exp((FSPC1_PM[2,2]-FSPC_T)/FSPC1_PM[3,2]))
HPHS_C3=FSPC1_PM[1,3]/(1+exp((FSPC1_PM[2,3]-FSPC_T)/FSPC1_PM[3,3]))
HPHS_C4=FSPC1_PM[1,4]/(1+exp((FSPC1_PM[2,4]-FSPC_T)/FSPC1_PM[3,4]))
HPHS_C5=FSPC1_PM[1,5]/(1+exp((FSPC1_PM[2,5]-FSPC_T)/FSPC1_PM[3,5]))
HPHS_C6=FSPC1_PM[1,6]/(1+exp((FSPC1_PM[2,6]-FSPC_T)/FSPC1_PM[3,6]))
HPHS_C7=FSPC1_PM[1,7]/(1+exp((FSPC1_PM[2,7]-FSPC_T)/FSPC1_PM[3,7]))
HPHS_C=HPHS_C1+HPHS_C2+HPHS_C3+HPHS_C4+HPHS_C5+HPHS_C6+HPHS_C7

#C12-2
HPMS_C1=FSPC1_PM[4,1]/(1+exp((FSPC1_PM[5,1]-FSPC_T)/FSPC1_PM[6,1]))
HPMS_C2=FSPC1_PM[4,2]/(1+exp((FSPC1_PM[5,2]-FSPC_T)/FSPC1_PM[6,2]))
HPMS_C3=FSPC1_PM[4,3]/(1+exp((FSPC1_PM[5,3]-FSPC_T)/FSPC1_PM[6,3]))
HPMS_C4=FSPC1_PM[4,4]/(1+exp((FSPC1_PM[5,4]-FSPC_T)/FSPC1_PM[6,4]))
HPMS_C5=FSPC1_PM[4,5]/(1+exp((FSPC1_PM[5,5]-FSPC_T)/FSPC1_PM[6,5]))
HPMS_C6=FSPC1_PM[4,6]/(1+exp((FSPC1_PM[5,6]-FSPC_T)/FSPC1_PM[6,6]))
HPMS_C7=FSPC1_PM[4,7]/(1+exp((FSPC1_PM[5,7]-FSPC_T)/FSPC1_PM[6,7]))
HPMS_C=HPMS_C1+HPMS_C2+HPMS_C3+HPMS_C4+HPMS_C5+HPMS_C6+HPMS_C7

#C12-3
HPLS_C1=FSPC1_PM[7,1]/(1+exp((FSPC1_PM[8,1]-FSPC_T)/FSPC1_PM[9,1]))
HPLS_C2=FSPC1_PM[7,2]/(1+exp((FSPC1_PM[8,2]-FSPC_T)/FSPC1_PM[9,2]))
HPLS_C3=FSPC1_PM[7,3]/(1+exp((FSPC1_PM[8,3]-FSPC_T)/FSPC1_PM[9,3]))
HPLS_C4=FSPC1_PM[7,4]/(1+exp((FSPC1_PM[8,4]-FSPC_T)/FSPC1_PM[9,4]))
HPLS_C5=FSPC1_PM[7,5]/(1+exp((FSPC1_PM[8,5]-FSPC_T)/FSPC1_PM[9,5]))
HPLS_C6=FSPC1_PM[7,6]/(1+exp((FSPC1_PM[8,6]-FSPC_T)/FSPC1_PM[9,6]))
HPLS_C7=FSPC1_PM[7,7]/(1+exp((FSPC1_PM[8,7]-FSPC_T)/FSPC1_PM[9,7]))
HPLS_C=HPLS_C1+HPLS_C2+HPLS_C3+HPLS_C4+HPLS_C5+HPLS_C6+HPLS_C7

#C12-4
HPNS_C1=FSPC1_PM[10,1]/(1+exp((FSPC1_PM[11,1]-FSPC_T)/FSPC1_PM[12,1]))
HPNS_C2=FSPC1_PM[10,2]/(1+exp((FSPC1_PM[11,2]-FSPC_T)/FSPC1_PM[12,2]))
HPNS_C3=FSPC1_PM[10,3]/(1+exp((FSPC1_PM[11,3]-FSPC_T)/FSPC1_PM[12,3]))
HPNS_C4=FSPC1_PM[10,4]/(1+exp((FSPC1_PM[11,4]-FSPC_T)/FSPC1_PM[12,4]))
HPNS_C5=FSPC1_PM[10,5]/(1+exp((FSPC1_PM[11,5]-FSPC_T)/FSPC1_PM[12,5]))
HPNS_C6=FSPC1_PM[10,6]/(1+exp((FSPC1_PM[11,6]-FSPC_T)/FSPC1_PM[12,6]))
HPNS_C7=FSPC1_PM[10,7]/(1+exp((FSPC1_PM[11,7]-FSPC_T)/FSPC1_PM[12,7]))
HPNS_C=HPNS_C1+HPNS_C2+HPNS_C3+HPNS_C4+HPNS_C5+HPNS_C6+HPNS_C7

#C12-5
MPHS_C1=FSPC1_PM[13,1]/(1+exp((FSPC1_PM[14,1]-FSPC_T)/FSPC1_PM[15,1]))
MPHS_C2=FSPC1_PM[13,2]/(1+exp((FSPC1_PM[14,2]-FSPC_T)/FSPC1_PM[15,2]))
MPHS_C3=FSPC1_PM[13,3]/(1+exp((FSPC1_PM[14,3]-FSPC_T)/FSPC1_PM[15,3]))
MPHS_C4=FSPC1_PM[13,4]/(1+exp((FSPC1_PM[14,4]-FSPC_T)/FSPC1_PM[15,4]))
MPHS_C5=FSPC1_PM[13,5]/(1+exp((FSPC1_PM[14,5]-FSPC_T)/FSPC1_PM[15,5]))
MPHS_C6=FSPC1_PM[13,6]/(1+exp((FSPC1_PM[14,6]-FSPC_T)/FSPC1_PM[15,6]))
MPHS_C7=FSPC1_PM[13,7]/(1+exp((FSPC1_PM[14,7]-FSPC_T)/FSPC1_PM[15,7]))
MPHS_C=MPHS_C1+MPHS_C2+MPHS_C3+MPHS_C4+MPHS_C5+MPHS_C6+MPHS_C7

#C12-6
MPMS_C1=FSPC1_PM[16,1]/(1+exp((FSPC1_PM[17,1]-FSPC_T)/FSPC1_PM[18,1]))
MPMS_C2=FSPC1_PM[16,2]/(1+exp((FSPC1_PM[17,2]-FSPC_T)/FSPC1_PM[18,2]))
MPMS_C3=FSPC1_PM[16,3]/(1+exp((FSPC1_PM[17,3]-FSPC_T)/FSPC1_PM[18,3]))
MPMS_C4=FSPC1_PM[16,4]/(1+exp((FSPC1_PM[17,4]-FSPC_T)/FSPC1_PM[18,4]))
MPMS_C5=FSPC1_PM[16,5]/(1+exp((FSPC1_PM[17,5]-FSPC_T)/FSPC1_PM[18,5]))
MPMS_C6=FSPC1_PM[16,6]/(1+exp((FSPC1_PM[17,6]-FSPC_T)/FSPC1_PM[18,6]))
MPMS_C7=FSPC1_PM[16,7]/(1+exp((FSPC1_PM[17,7]-FSPC_T)/FSPC1_PM[18,7]))
MPMS_C=MPMS_C1+MPMS_C2+MPMS_C3+MPMS_C4+MPMS_C5+MPMS_C6+MPMS_C7

#C12-7
MPLS_C1=FSPC1_PM[19,1]/(1+exp((FSPC1_PM[20,1]-FSPC_T)/FSPC1_PM[21,1]))
MPLS_C2=FSPC1_PM[19,2]/(1+exp((FSPC1_PM[20,2]-FSPC_T)/FSPC1_PM[21,2]))
MPLS_C3=FSPC1_PM[19,3]/(1+exp((FSPC1_PM[20,3]-FSPC_T)/FSPC1_PM[21,3]))
MPLS_C4=FSPC1_PM[19,4]/(1+exp((FSPC1_PM[20,4]-FSPC_T)/FSPC1_PM[21,4]))
MPLS_C5=FSPC1_PM[19,5]/(1+exp((FSPC1_PM[20,5]-FSPC_T)/FSPC1_PM[21,5]))
MPLS_C6=FSPC1_PM[19,6]/(1+exp((FSPC1_PM[20,6]-FSPC_T)/FSPC1_PM[21,6]))
MPLS_C7=FSPC1_PM[19,7]/(1+exp((FSPC1_PM[20,7]-FSPC_T)/FSPC1_PM[21,7]))
MPLS_C=MPLS_C1+MPLS_C2+MPLS_C3+MPLS_C4+MPLS_C5+MPLS_C6+MPLS_C7

#C12-8
MPNS_C1=FSPC1_PM[22,1]/(1+exp((FSPC1_PM[23,1]-FSPC_T)/FSPC1_PM[24,1]))
MPNS_C2=FSPC1_PM[22,2]/(1+exp((FSPC1_PM[23,2]-FSPC_T)/FSPC1_PM[24,2]))
MPNS_C3=FSPC1_PM[22,3]/(1+exp((FSPC1_PM[23,3]-FSPC_T)/FSPC1_PM[24,3]))
MPNS_C4=FSPC1_PM[22,4]/(1+exp((FSPC1_PM[23,4]-FSPC_T)/FSPC1_PM[24,4]))
MPNS_C5=FSPC1_PM[22,5]/(1+exp((FSPC1_PM[23,5]-FSPC_T)/FSPC1_PM[24,5]))
MPNS_C6=FSPC1_PM[22,6]/(1+exp((FSPC1_PM[23,6]-FSPC_T)/FSPC1_PM[24,6]))
MPNS_C7=FSPC1_PM[22,7]/(1+exp((FSPC1_PM[23,7]-FSPC_T)/FSPC1_PM[24,7]))
MPNS_C=MPNS_C1+MPNS_C2+MPNS_C3+MPNS_C4+MPNS_C5+MPNS_C6+MPNS_C7

#C12-9
LPHS_C1=FSPC1_PM[25,1]/(1+exp((FSPC1_PM[26,1]-FSPC_T)/FSPC1_PM[27,1]))
LPHS_C2=FSPC1_PM[25,2]/(1+exp((FSPC1_PM[26,2]-FSPC_T)/FSPC1_PM[27,2]))
LPHS_C3=FSPC1_PM[25,3]/(1+exp((FSPC1_PM[26,3]-FSPC_T)/FSPC1_PM[27,3]))
LPHS_C4=FSPC1_PM[25,4]/(1+exp((FSPC1_PM[26,4]-FSPC_T)/FSPC1_PM[27,4]))
LPHS_C5=FSPC1_PM[25,5]/(1+exp((FSPC1_PM[26,5]-FSPC_T)/FSPC1_PM[27,5]))
LPHS_C6=FSPC1_PM[25,6]/(1+exp((FSPC1_PM[26,6]-FSPC_T)/FSPC1_PM[27,6]))
LPHS_C7=FSPC1_PM[25,7]/(1+exp((FSPC1_PM[26,7]-FSPC_T)/FSPC1_PM[27,7]))
LPHS_C=LPHS_C1+LPHS_C2+LPHS_C3+LPHS_C4+LPHS_C5+LPHS_C6+LPHS_C7

#C12-10
LPMS_C1=FSPC1_PM[28,1]/(1+exp((FSPC1_PM[29,1]-FSPC_T)/FSPC1_PM[30,1]))
LPMS_C2=FSPC1_PM[28,2]/(1+exp((FSPC1_PM[29,2]-FSPC_T)/FSPC1_PM[30,2]))
LPMS_C3=FSPC1_PM[28,3]/(1+exp((FSPC1_PM[29,3]-FSPC_T)/FSPC1_PM[30,3]))
LPMS_C4=FSPC1_PM[28,4]/(1+exp((FSPC1_PM[29,4]-FSPC_T)/FSPC1_PM[30,4]))
LPMS_C5=FSPC1_PM[28,5]/(1+exp((FSPC1_PM[29,5]-FSPC_T)/FSPC1_PM[30,5]))
LPMS_C6=FSPC1_PM[28,6]/(1+exp((FSPC1_PM[29,6]-FSPC_T)/FSPC1_PM[30,6]))
LPMS_C7=FSPC1_PM[28,7]/(1+exp((FSPC1_PM[29,7]-FSPC_T)/FSPC1_PM[30,7]))
LPMS_C=LPMS_C1+LPMS_C2+LPMS_C3+LPMS_C4+LPMS_C5+LPMS_C6+LPMS_C7

#C12-11
LPLS_C1=FSPC1_PM[31,1]/(1+exp((FSPC1_PM[32,1]-FSPC_T)/FSPC1_PM[33,1]))
LPLS_C2=FSPC1_PM[31,2]/(1+exp((FSPC1_PM[32,2]-FSPC_T)/FSPC1_PM[33,2]))
LPLS_C3=FSPC1_PM[31,3]/(1+exp((FSPC1_PM[32,3]-FSPC_T)/FSPC1_PM[33,3]))
LPLS_C4=FSPC1_PM[31,4]/(1+exp((FSPC1_PM[32,4]-FSPC_T)/FSPC1_PM[33,4]))
LPLS_C5=FSPC1_PM[31,5]/(1+exp((FSPC1_PM[32,5]-FSPC_T)/FSPC1_PM[33,5]))
LPLS_C6=FSPC1_PM[31,6]/(1+exp((FSPC1_PM[32,6]-FSPC_T)/FSPC1_PM[33,6]))
LPLS_C7=FSPC1_PM[31,7]/(1+exp((FSPC1_PM[32,7]-FSPC_T)/FSPC1_PM[33,7]))
LPLS_C=LPLS_C1+LPLS_C2+LPLS_C3+LPLS_C4+LPLS_C5+LPLS_C6+LPLS_C7

#C12-12
LPNS_C1=FSPC1_PM[34,1]/(1+exp((FSPC1_PM[35,1]-FSPC_T)/FSPC1_PM[36,1]))
LPNS_C2=FSPC1_PM[34,2]/(1+exp((FSPC1_PM[35,2]-FSPC_T)/FSPC1_PM[36,2]))
LPNS_C3=FSPC1_PM[34,3]/(1+exp((FSPC1_PM[35,3]-FSPC_T)/FSPC1_PM[36,3]))
LPNS_C4=FSPC1_PM[34,4]/(1+exp((FSPC1_PM[35,4]-FSPC_T)/FSPC1_PM[36,4]))
LPNS_C5=FSPC1_PM[34,5]/(1+exp((FSPC1_PM[35,5]-FSPC_T)/FSPC1_PM[36,5]))
LPNS_C6=FSPC1_PM[34,6]/(1+exp((FSPC1_PM[35,6]-FSPC_T)/FSPC1_PM[36,6]))
LPNS_C7=FSPC1_PM[34,7]/(1+exp((FSPC1_PM[35,7]-FSPC_T)/FSPC1_PM[36,7]))
LPNS_C=LPNS_C1+LPNS_C2+LPNS_C3+LPNS_C4+LPNS_C5+LPNS_C6+LPNS_C7

FSPC1_C=rbind(HPHS_C1,HPHS_C2,HPHS_C3,HPHS_C4,HPHS_C5,HPHS_C6,HPHS_C7,HPHS_C,
              HPMS_C1,HPMS_C2,HPMS_C3,HPMS_C4,HPMS_C5,HPMS_C6,HPMS_C7,HPMS_C,
              HPLS_C1,HPLS_C2,HPLS_C3,HPLS_C4,HPLS_C5,HPLS_C6,HPLS_C7,HPLS_C,
              HPNS_C1,HPNS_C2,HPNS_C3,HPNS_C4,HPNS_C5,HPNS_C6,HPNS_C7,HPNS_C,
              MPHS_C1,MPHS_C2,MPHS_C3,MPHS_C4,MPHS_C5,MPHS_C6,MPHS_C7,MPHS_C,
              MPMS_C1,MPMS_C2,MPMS_C3,MPMS_C4,MPMS_C5,MPMS_C6,MPMS_C7,MPMS_C,
              MPLS_C1,MPLS_C2,MPLS_C3,MPLS_C4,MPLS_C5,MPLS_C6,MPLS_C7,MPLS_C,
              MPNS_C1,MPNS_C2,MPNS_C3,MPNS_C4,MPNS_C5,MPNS_C6,MPNS_C7,MPNS_C,
              LPHS_C1,LPHS_C2,LPHS_C3,LPHS_C4,LPHS_C5,LPHS_C6,LPHS_C7,LPHS_C,
              LPMS_C1,LPMS_C2,LPMS_C3,LPMS_C4,LPMS_C5,LPMS_C6,LPMS_C7,LPMS_C,
              LPLS_C1,LPLS_C2,LPLS_C3,LPLS_C4,LPLS_C5,LPLS_C6,LPLS_C7,LPLS_C,
              LPNS_C1,LPNS_C2,LPNS_C3,LPNS_C4,LPNS_C5,LPNS_C6,LPNS_C7,LPNS_C)
FSPC1_C=t(FSPC1_C)
write.csv(FSPC1_C,"OUT/CalResults_FSPC1.csv")

FSPC1_C=read.csv("OUT/CalResults_FSPC1.csv")
POP=read.xlsx("Al Data_1950-2100.xlsx", sheet="POP")
TS1_HPOP=FSPC1_C[,2:97]*POP[1:151,5]
TS1_MPOP=FSPC1_C[,2:97]*POP[1:151,4]
TS1_LPOP=FSPC1_C[,2:97]*POP[1:151,3]
TS1=cbind(TS1_HPOP,TS1_MPOP,TS1_LPOP)
write.csv(TS1,"OUT/CalResults_TS1横版.csv")
TS1=rbind(TS1_HPOP,TS1_MPOP,TS1_LPOP)
write.csv(TS1,"OUT/CalResults_TS1竖版.csv")

#报废年限设置及模拟
setwd("E:/RBOOK/AL-高速度")

TS1=read.csv("OUT/CalResults_TS1竖版.CSV")
DS1=TS1[2:455,2:97]-TS1[1:454,2:97]
write.csv(DS1,"OUT/CalResults_DS1.csv")
#这里的DS1文件需要些许的加工，在excel中进行，加工包括：
#1.去掉total列
#2.把1950年数据添加成0
DS1=read.csv("DS1.csv")
IN_TT=read.csv("IN.csv")
OU_TT=read.csv("OU.csv")
DP=read.csv("DP.csv",header = FALSE)
#36_1 HPHSHO 
DS=DS1[1:151,2:8]
IN=IN_TT[1:151,2:8]
OU=OU_TT[1:151,2:8]
IN36_1=data.frame(matrix(0,151,7))
OU36_1=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_1[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_1[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_1=cbind(IN36_1,OU36_1,DS)
write.csv(INOU36_1,"OUT/IN&OUT36_1.csv")

#36_2 HPMSHO 
DS=DS1[1:151,9:15]
IN=IN_TT[1:151,9:15]
OU=OU_TT[1:151,9:15]
IN36_2=data.frame(matrix(0,151,7))
OU36_2=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_2[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_2[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_2=cbind(IN36_2,OU36_2,DS)
write.csv(INOU36_2,"OUT/IN&OUT36_2.csv")

#36_3 HPLSHO 
DS=DS1[1:151,16:22]
IN=IN_TT[1:151,16:22]
OU=OU_TT[1:151,16:22]
IN36_3=data.frame(matrix(0,151,7))
OU36_3=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_3[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_3[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_3=cbind(IN36_3,OU36_3,DS)
write.csv(INOU36_3,"OUT/IN&OUT36_3.csv")

#36_4 HPNSHO 
DS=DS1[1:151,23:29]
IN=IN_TT[1:151,23:29]
OU=OU_TT[1:151,23:29]
IN36_4=data.frame(matrix(0,151,7))
OU36_4=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_4[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_4[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_4=cbind(IN36_4,OU36_4,DS)
write.csv(INOU36_4,"OUT/IN&OUT36_4.csv")

#36_5 MPHSHO 
DS=DS1[1:151,30:36]
IN=IN_TT[1:151,30:36]
OU=OU_TT[1:151,30:36]
IN36_5=data.frame(matrix(0,151,7))
OU36_5=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_5[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_5[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_5=cbind(IN36_5,OU36_5,DS)
write.csv(INOU36_5,"OUT/IN&OUT36_5.csv")

#36_6 MPMSHO 
DS=DS1[1:151,37:43]
IN=IN_TT[1:151,37:43]
OU=OU_TT[1:151,37:43]
IN36_6=data.frame(matrix(0,151,7))
OU36_6=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_6[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_6[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_6=cbind(IN36_6,OU36_6,DS)
write.csv(INOU36_6,"OUT/IN&OUT36_6.csv")

#36_7 MPLSHO 
DS=DS1[1:151,44:50]
IN=IN_TT[1:151,44:50]
OU=OU_TT[1:151,44:50]
IN36_7=data.frame(matrix(0,151,7))
OU36_7=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_7[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_7[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_7=cbind(IN36_7,OU36_7,DS)
write.csv(INOU36_7,"OUT/IN&OUT36_7.csv")

#36_8 MPNSHO 
DS=DS1[1:151,51:57]
IN=IN_TT[1:151,51:57]
OU=OU_TT[1:151,51:57]
IN36_8=data.frame(matrix(0,151,7))
OU36_8=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_8[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_8[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_8=cbind(IN36_8,OU36_8,DS)
write.csv(INOU36_8,"OUT/IN&OUT36_8.csv")

#36_9 LPHSHO 
DS=DS1[1:151,58:64]
IN=IN_TT[1:151,58:64]
OU=OU_TT[1:151,58:64]
IN36_9=data.frame(matrix(0,151,7))
OU36_9=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_9[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_9[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_9=cbind(IN36_9,OU36_9,DS)
write.csv(INOU36_9,"OUT/IN&OUT36_9.csv")

#36_10 LPMSHO 
DS=DS1[1:151,65:71]
IN=IN_TT[1:151,65:71]
OU=OU_TT[1:151,65:71]
IN36_10=data.frame(matrix(0,151,7))
OU36_10=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_10[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_10[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_10=cbind(IN36_10,OU36_10,DS)
write.csv(INOU36_10,"OUT/IN&OUT36_10.csv")

#36_11 LPLSHO 
DS=DS1[1:151,72:78]
IN=IN_TT[1:151,72:78]
OU=OU_TT[1:151,72:78]
IN36_11=data.frame(matrix(0,151,7))
OU36_11=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_11[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_11[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_11=cbind(IN36_11,OU36_11,DS)
write.csv(INOU36_11,"OUT/IN&OUT36_11.csv")

#36_12 LPNSMO 
DS=DS1[1:151,79:85]
IN=IN_TT[1:151,79:85]
OU=OU_TT[1:151,79:85]
IN36_12=data.frame(matrix(0,151,7))
OU36_12=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_12[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_12[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_12=cbind(IN36_12,OU36_12,DS)
write.csv(INOU36_12,"OUT/IN&OUT36_12.csv")

#36_13 HPHSMO 
DS=DS1[152:302,2:8]
IN=IN_TT[152:302,2:8]
OU=OU_TT[152:302,2:8]
IN36_13=data.frame(matrix(0,151,7))
OU36_13=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_13[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_13[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_13=cbind(IN36_13,OU36_13,DS)
write.csv(INOU36_13,"OUT/IN&OUT36_13.csv")

#36_14 HPMSMO 
DS=DS1[152:302,9:15]
IN=IN_TT[152:302,9:15]
OU=OU_TT[152:302,9:15]
IN36_14=data.frame(matrix(0,151,7))
OU36_14=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_14[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_14[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_14=cbind(IN36_14,OU36_14,DS)
write.csv(INOU36_14,"OUT/IN&OUT36_14.csv")

#36_15 HPLSMO 
DS=DS1[152:302,16:22]
IN=IN_TT[152:302,16:22]
OU=OU_TT[152:302,16:22]
IN36_15=data.frame(matrix(0,151,7))
OU36_15=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_15[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_15[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_15=cbind(IN36_15,OU36_15,DS)
write.csv(INOU36_15,"OUT/IN&OUT36_15.csv")

#36_16 HPNSMO 
DS=DS1[152:302,23:29]
IN=IN_TT[152:302,23:29]
OU=OU_TT[152:302,23:29]
IN36_16=data.frame(matrix(0,151,7))
OU36_16=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_16[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_16[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_16=cbind(IN36_16,OU36_16,DS)
write.csv(INOU36_16,"OUT/IN&OUT36_16.csv")

#36_17 MPHSMO 
DS=DS1[152:302,30:36]
IN=IN_TT[152:302,30:36]
OU=OU_TT[152:302,30:36]
IN36_17=data.frame(matrix(0,151,7))
OU36_17=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_17[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_17[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_17=cbind(IN36_17,OU36_17,DS)
write.csv(INOU36_17,"OUT/IN&OUT36_17.csv")

#36_18 MPMSMO 
DS=DS1[152:302,37:43]
IN=IN_TT[152:302,37:43]
OU=OU_TT[152:302,37:43]
IN36_18=data.frame(matrix(0,151,7))
OU36_18=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_18[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_18[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_18=cbind(IN36_18,OU36_18,DS)
write.csv(INOU36_18,"OUT/IN&OUT36_18.csv")

#36_19 MPLSMO 
DS=DS1[152:302,44:50]
IN=IN_TT[152:302,44:50]
OU=OU_TT[152:302,44:50]
IN36_19=data.frame(matrix(0,151,7))
OU36_19=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_19[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_19[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_19=cbind(IN36_19,OU36_19,DS)
write.csv(INOU36_19,"OUT/IN&OUT36_19.csv")

#36_20 MPNSMO 
DS=DS1[152:302,51:57]
IN=IN_TT[152:302,51:57]
OU=OU_TT[152:302,51:57]
IN36_20=data.frame(matrix(0,151,7))
OU36_20=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_20[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_20[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_20=cbind(IN36_20,OU36_20,DS)
write.csv(INOU36_20,"OUT/IN&OUT36_20.csv")


#36_21 LPHSMO 
DS=DS1[152:302,58:64]
IN=IN_TT[152:302,58:64]
OU=OU_TT[152:302,58:64]
IN36_21=data.frame(matrix(0,151,7))
OU36_21=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_21[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_21[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_21=cbind(IN36_21,OU36_21,DS)
write.csv(INOU36_21,"OUT/IN&OUT36_21.csv")

#36_22 LPMSMO 
DS=DS1[152:302,65:71]
IN=IN_TT[152:302,65:71]
OU=OU_TT[152:302,65:71]
IN36_22=data.frame(matrix(0,151,7))
OU36_22=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_22[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_22[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_22=cbind(IN36_22,OU36_22,DS)
write.csv(INOU36_22,"OUT/IN&OUT36_22.csv")

#36_23 LPLSMO 
DS=DS1[152:302,72:78]
IN=IN_TT[152:302,72:78]
OU=OU_TT[152:302,72:78]
IN36_23=data.frame(matrix(0,151,7))
OU36_23=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_23[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_23[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_23=cbind(IN36_23,OU36_23,DS)
write.csv(INOU36_23,"OUT/IN&OUT36_23.csv")

#36_24 LPNSMO 
DS=DS1[152:302,79:85]
IN=IN_TT[152:302,79:85]
OU=OU_TT[152:302,79:85]
IN36_24=data.frame(matrix(0,151,7))
OU36_24=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_24[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_24[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_24=cbind(IN36_24,OU36_24,DS)
write.csv(INOU36_24,"OUT/IN&OUT36_24.csv")

#36_25 HPHSLO 
DS=DS1[303:453,2:8]
IN=IN_TT[303:453,2:8]
OU=OU_TT[303:453,2:8]
IN36_25=data.frame(matrix(0,151,7))
OU36_25=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_25[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_25[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_25=cbind(IN36_25,OU36_25,DS)
write.csv(INOU36_25,"OUT/IN&OUT36_25.csv")

#36_26 HPMSLO 
DS=DS1[303:453,9:15]
IN=IN_TT[303:453,9:15]
OU=OU_TT[303:453,9:15]
IN36_26=data.frame(matrix(0,151,7))
OU36_26=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_26[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_26[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_26=cbind(IN36_26,OU36_26,DS)
write.csv(INOU36_26,"OUT/IN&OUT36_26.csv")

#36_27 HPLSLO 
DS=DS1[303:453,16:22]
IN=IN_TT[303:453,16:22]
OU=OU_TT[303:453,16:22]
IN36_27=data.frame(matrix(0,151,7))
OU36_27=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_27[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_27[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_27=cbind(IN36_27,OU36_27,DS)
write.csv(INOU36_27,"OUT/IN&OUT36_27.csv")

#36_28 HPNSLO 
DS=DS1[303:453,23:29]
IN=IN_TT[303:453,23:29]
OU=OU_TT[303:453,23:29]
IN36_28=data.frame(matrix(0,151,7))
OU36_28=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_28[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_28[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_28=cbind(IN36_28,OU36_28,DS)
write.csv(INOU36_28,"OUT/IN&OUT36_28.csv")

#36_29 MPHSLO 
DS=DS1[303:453,30:36]
IN=IN_TT[303:453,30:36]
OU=OU_TT[303:453,30:36]
IN36_29=data.frame(matrix(0,151,7))
OU36_29=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_29[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_29[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_29=cbind(IN36_29,OU36_29,DS)
write.csv(INOU36_29,"OUT/IN&OUT36_29.csv")

#36_30 MPMSLO 
DS=DS1[303:453,37:43]
IN=IN_TT[303:453,37:43]
OU=OU_TT[303:453,37:43]
IN36_30=data.frame(matrix(0,151,7))
OU36_30=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_30[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_30[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_30=cbind(IN36_30,OU36_30,DS)
write.csv(INOU36_30,"OUT/IN&OUT36_30.csv")

#36_31 MPLSLO 
DS=DS1[303:453,44:50]
IN=IN_TT[303:453,44:50]
OU=OU_TT[303:453,44:50]
IN36_31=data.frame(matrix(0,151,7))
OU36_31=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_31[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_31[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_31=cbind(IN36_31,OU36_31,DS)
write.csv(INOU36_31,"OUT/IN&OUT36_31.csv")

#36_32 MPNSLO 
DS=DS1[303:453,51:57]
IN=IN_TT[303:453,51:57]
OU=OU_TT[303:453,51:57]
IN36_32=data.frame(matrix(0,151,7))
OU36_32=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_32[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_32[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_32=cbind(IN36_32,OU36_32,DS)
write.csv(INOU36_32,"OUT/IN&OUT36_32.csv")

#36_33 LPHSLO 
DS=DS1[303:453,58:64]
IN=IN_TT[303:453,58:64]
OU=OU_TT[303:453,58:64]
IN36_33=data.frame(matrix(0,151,7))
OU36_33=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_33[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_33[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_33=cbind(IN36_33,OU36_33,DS)
write.csv(INOU36_33,"OUT/IN&OUT36_33.csv")

#36_34 LPMSLO 
DS=DS1[303:453,65:71]
IN=IN_TT[303:453,65:71]
OU=OU_TT[303:453,65:71]
IN36_34=data.frame(matrix(0,151,7))
OU36_34=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_34[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_34[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_34=cbind(IN36_34,OU36_34,DS)
write.csv(INOU36_34,"OUT/IN&OUT36_34.csv")

#36_35 LPLSLO 
DS=DS1[303:453,72:78]
IN=IN_TT[303:453,72:78]
OU=OU_TT[303:453,72:78]
IN36_35=data.frame(matrix(0,151,7))
OU36_35=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_35[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_35[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_35=cbind(IN36_35,OU36_35,DS)
write.csv(INOU36_35,"OUT/IN&OUT36_35.csv")

#36_36 LPNSLO 
DS=DS1[303:453,79:85]
IN=IN_TT[303:453,79:85]
OU=OU_TT[303:453,79:85]
IN36_36=data.frame(matrix(0,151,7))
OU36_36=data.frame(matrix(0,151,7))
t=seq(1950,2100,1)
for (m in 1:3){
  for (k in 1:7){
    for (i in 2:151) {
      j=1:i
      OU[i,k]=sum(IN[t[j]-1949,k]*dnorm(t[i]-t[j],DP[k,m],DP[k,m+3]))
      IN[i,k]=OU[i,k]+DS[i,k]
    }
  }
  OU36_36[,((m-1)*7+1):((m-1)*7+7)]=OU[,1:7]
  IN36_36[,((m-1)*7+1):((m-1)*7+7)]=IN[,1:7]
}
DS=cbind(DS,DS,DS)
INOU36_36=cbind(IN36_36,OU36_36,DS)
write.csv(INOU36_36,"OUT/IN&OUT36_36.csv")

#总量
#该步骤在excel中进行
#1.将以上36个文件夹合并在一个工作簿中
#2.选中所有sheet，计算sector6的所有数值
#3.选中所有sheet，添加总量一列，进行求和计算

#提取108种情况（峰值3*速度4*人口3*寿命3）中的7个sector总和
library(openxlsx)
setwd("E:/RBOOK/AL-高速度")
INOU36_1R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_1")
INOU36_2R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_2")
INOU36_3R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_3")
INOU36_4R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_4")
INOU36_5R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_5")
INOU36_6R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_6")
INOU36_7R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_7")
INOU36_8R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_8")
INOU36_9R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_9")
INOU36_10R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_10")
INOU36_11R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_11")
INOU36_12R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_12")
INOU36_13R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_13")
INOU36_14R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_14")
INOU36_15R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_15")
INOU36_16R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_16")
INOU36_17R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_17")
INOU36_18R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_18")
INOU36_19R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_19")
INOU36_20R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_20")
INOU36_21R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_21")
INOU36_22R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_22")
INOU36_23R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_23")
INOU36_24R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_24")
INOU36_25R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_25")
INOU36_26R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_26")
INOU36_27R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_27")
INOU36_28R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_28")
INOU36_29R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_29")
INOU36_30R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_31")
INOU36_31R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_32")
INOU36_32R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_33")
INOU36_33R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_30")
INOU36_34R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_34")
INOU36_35R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_35")
INOU36_36R=read.xlsx("IN&OUT.xlsx",sheet = "IN&OUT36_36")

IN36_1R_LLS=INOU36_1R[3:153,9]
IN36_1R_MLS=INOU36_1R[3:153,17]
IN36_1R_HLS=INOU36_1R[3:153,25]
OU36_1R_LLS=INOU36_1R[3:153,33]
OU36_1R_MLS=INOU36_1R[3:153,41]
OU36_1R_HLS=INOU36_1R[3:153,49]

IN36_2R_LLS=INOU36_2R[3:153,9]
IN36_2R_MLS=INOU36_2R[3:153,17]
IN36_2R_HLS=INOU36_2R[3:153,25]
OU36_2R_LLS=INOU36_2R[3:153,33]
OU36_2R_MLS=INOU36_2R[3:153,41]
OU36_2R_HLS=INOU36_2R[3:153,49]

IN36_3R_LLS=INOU36_3R[3:153,9]
IN36_3R_MLS=INOU36_3R[3:153,17]
IN36_3R_HLS=INOU36_3R[3:153,25]
OU36_3R_LLS=INOU36_3R[3:153,33]
OU36_3R_MLS=INOU36_3R[3:153,41]
OU36_3R_HLS=INOU36_3R[3:153,49]

IN36_4R_LLS=INOU36_4R[3:153,9]
IN36_4R_MLS=INOU36_4R[3:153,17]
IN36_4R_HLS=INOU36_4R[3:153,25]
OU36_4R_LLS=INOU36_4R[3:153,33]
OU36_4R_MLS=INOU36_4R[3:153,41]
OU36_4R_HLS=INOU36_4R[3:153,49]

IN36_5R_LLS=INOU36_5R[3:153,9]
IN36_5R_MLS=INOU36_5R[3:153,17]
IN36_5R_HLS=INOU36_5R[3:153,25]
OU36_5R_LLS=INOU36_5R[3:153,33]
OU36_5R_MLS=INOU36_5R[3:153,41]
OU36_5R_HLS=INOU36_5R[3:153,49]

IN36_6R_LLS=INOU36_6R[3:153,9]
IN36_6R_MLS=INOU36_6R[3:153,17]
IN36_6R_HLS=INOU36_6R[3:153,25]
OU36_6R_LLS=INOU36_6R[3:153,33]
OU36_6R_MLS=INOU36_6R[3:153,41]
OU36_6R_HLS=INOU36_6R[3:153,49]

IN36_7R_LLS=INOU36_7R[3:153,9]
IN36_7R_MLS=INOU36_7R[3:153,17]
IN36_7R_HLS=INOU36_7R[3:153,25]
OU36_7R_LLS=INOU36_7R[3:153,33]
OU36_7R_MLS=INOU36_7R[3:153,41]
OU36_7R_HLS=INOU36_7R[3:153,49]

IN36_8R_LLS=INOU36_8R[3:153,9]
IN36_8R_MLS=INOU36_8R[3:153,17]
IN36_8R_HLS=INOU36_8R[3:153,25]
OU36_8R_LLS=INOU36_8R[3:153,33]
OU36_8R_MLS=INOU36_8R[3:153,41]
OU36_8R_HLS=INOU36_8R[3:153,49]

IN36_9R_LLS=INOU36_9R[3:153,9]
IN36_9R_MLS=INOU36_9R[3:153,17]
IN36_9R_HLS=INOU36_9R[3:153,25]
OU36_9R_LLS=INOU36_9R[3:153,33]
OU36_9R_MLS=INOU36_9R[3:153,41]
OU36_9R_HLS=INOU36_9R[3:153,49]

IN36_10R_LLS=INOU36_10R[3:153,9]
IN36_10R_MLS=INOU36_10R[3:153,17]
IN36_10R_HLS=INOU36_10R[3:153,25]
OU36_10R_LLS=INOU36_10R[3:153,33]
OU36_10R_MLS=INOU36_10R[3:153,41]
OU36_10R_HLS=INOU36_10R[3:153,49]

IN36_11R_LLS=INOU36_11R[3:153,9]
IN36_11R_MLS=INOU36_11R[3:153,17]
IN36_11R_HLS=INOU36_11R[3:153,25]
OU36_11R_LLS=INOU36_11R[3:153,33]
OU36_11R_MLS=INOU36_11R[3:153,41]
OU36_11R_HLS=INOU36_11R[3:153,49]

IN36_12R_LLS=INOU36_12R[3:153,9]
IN36_12R_MLS=INOU36_12R[3:153,17]
IN36_12R_HLS=INOU36_12R[3:153,25]
OU36_12R_LLS=INOU36_12R[3:153,33]
OU36_12R_MLS=INOU36_12R[3:153,41]
OU36_12R_HLS=INOU36_12R[3:153,49]

IN36_13R_LLS=INOU36_13R[3:153,9]
IN36_13R_MLS=INOU36_13R[3:153,17]
IN36_13R_HLS=INOU36_13R[3:153,25]
OU36_13R_LLS=INOU36_13R[3:153,33]
OU36_13R_MLS=INOU36_13R[3:153,41]
OU36_13R_HLS=INOU36_13R[3:153,49]

IN36_14R_LLS=INOU36_14R[3:153,9]
IN36_14R_MLS=INOU36_14R[3:153,17]
IN36_14R_HLS=INOU36_14R[3:153,25]
OU36_14R_LLS=INOU36_14R[3:153,33]
OU36_14R_MLS=INOU36_14R[3:153,41]
OU36_14R_HLS=INOU36_14R[3:153,49]

IN36_15R_LLS=INOU36_15R[3:153,9]
IN36_15R_MLS=INOU36_15R[3:153,17]
IN36_15R_HLS=INOU36_15R[3:153,25]
OU36_15R_LLS=INOU36_15R[3:153,33]
OU36_15R_MLS=INOU36_15R[3:153,41]
OU36_15R_HLS=INOU36_15R[3:153,49]

IN36_16R_LLS=INOU36_16R[3:153,9]
IN36_16R_MLS=INOU36_16R[3:153,17]
IN36_16R_HLS=INOU36_16R[3:153,25]
OU36_16R_LLS=INOU36_16R[3:153,33]
OU36_16R_MLS=INOU36_16R[3:153,41]
OU36_16R_HLS=INOU36_16R[3:153,49]

IN36_17R_LLS=INOU36_17R[3:153,9]
IN36_17R_MLS=INOU36_17R[3:153,17]
IN36_17R_HLS=INOU36_17R[3:153,25]
OU36_17R_LLS=INOU36_17R[3:153,33]
OU36_17R_MLS=INOU36_17R[3:153,41]
OU36_17R_HLS=INOU36_17R[3:153,49]

IN36_18R_LLS=INOU36_18R[3:153,9]
IN36_18R_MLS=INOU36_18R[3:153,17]
IN36_18R_HLS=INOU36_18R[3:153,25]
OU36_18R_LLS=INOU36_18R[3:153,33]
OU36_18R_MLS=INOU36_18R[3:153,41]
OU36_18R_HLS=INOU36_18R[3:153,49]

IN36_19R_LLS=INOU36_19R[3:153,9]
IN36_19R_MLS=INOU36_19R[3:153,17]
IN36_19R_HLS=INOU36_19R[3:153,25]
OU36_19R_LLS=INOU36_19R[3:153,33]
OU36_19R_MLS=INOU36_19R[3:153,41]
OU36_19R_HLS=INOU36_19R[3:153,49]

IN36_20R_LLS=INOU36_20R[3:153,9]
IN36_20R_MLS=INOU36_20R[3:153,17]
IN36_20R_HLS=INOU36_20R[3:153,25]
OU36_20R_LLS=INOU36_20R[3:153,33]
OU36_20R_MLS=INOU36_20R[3:153,41]
OU36_20R_HLS=INOU36_20R[3:153,49]

IN36_21R_LLS=INOU36_21R[3:153,9]
IN36_21R_MLS=INOU36_21R[3:153,17]
IN36_21R_HLS=INOU36_21R[3:153,25]
OU36_21R_LLS=INOU36_21R[3:153,33]
OU36_21R_MLS=INOU36_21R[3:153,41]
OU36_21R_HLS=INOU36_21R[3:153,49]

IN36_22R_LLS=INOU36_22R[3:153,9]
IN36_22R_MLS=INOU36_22R[3:153,17]
IN36_22R_HLS=INOU36_22R[3:153,25]
OU36_22R_LLS=INOU36_22R[3:153,33]
OU36_22R_MLS=INOU36_22R[3:153,41]
OU36_22R_HLS=INOU36_22R[3:153,49]

IN36_23R_LLS=INOU36_23R[3:153,9]
IN36_23R_MLS=INOU36_23R[3:153,17]
IN36_23R_HLS=INOU36_23R[3:153,25]
OU36_23R_LLS=INOU36_23R[3:153,33]
OU36_23R_MLS=INOU36_23R[3:153,41]
OU36_23R_HLS=INOU36_23R[3:153,49]

IN36_24R_LLS=INOU36_24R[3:153,9]
IN36_24R_MLS=INOU36_24R[3:153,17]
IN36_24R_HLS=INOU36_24R[3:153,25]
OU36_24R_LLS=INOU36_24R[3:153,33]
OU36_24R_MLS=INOU36_24R[3:153,41]
OU36_24R_HLS=INOU36_24R[3:153,49]

IN36_25R_LLS=INOU36_25R[3:153,9]
IN36_25R_MLS=INOU36_25R[3:153,17]
IN36_25R_HLS=INOU36_25R[3:153,25]
OU36_25R_LLS=INOU36_25R[3:153,33]
OU36_25R_MLS=INOU36_25R[3:153,41]
OU36_25R_HLS=INOU36_25R[3:153,49]

IN36_26R_LLS=INOU36_26R[3:153,9]
IN36_26R_MLS=INOU36_26R[3:153,17]
IN36_26R_HLS=INOU36_26R[3:153,25]
OU36_26R_LLS=INOU36_26R[3:153,33]
OU36_26R_MLS=INOU36_26R[3:153,41]
OU36_26R_HLS=INOU36_26R[3:153,49]

IN36_27R_LLS=INOU36_27R[3:153,9]
IN36_27R_MLS=INOU36_27R[3:153,17]
IN36_27R_HLS=INOU36_27R[3:153,25]
OU36_27R_LLS=INOU36_27R[3:153,33]
OU36_27R_MLS=INOU36_27R[3:153,41]
OU36_27R_HLS=INOU36_27R[3:153,49]

IN36_28R_LLS=INOU36_28R[3:153,9]
IN36_28R_MLS=INOU36_28R[3:153,17]
IN36_28R_HLS=INOU36_28R[3:153,25]
OU36_28R_LLS=INOU36_28R[3:153,33]
OU36_28R_MLS=INOU36_28R[3:153,41]
OU36_28R_HLS=INOU36_28R[3:153,49]

IN36_29R_LLS=INOU36_29R[3:153,9]
IN36_29R_MLS=INOU36_29R[3:153,17]
IN36_29R_HLS=INOU36_29R[3:153,25]
OU36_29R_LLS=INOU36_29R[3:153,33]
OU36_29R_MLS=INOU36_29R[3:153,41]
OU36_29R_HLS=INOU36_29R[3:153,49]

IN36_30R_LLS=INOU36_30R[3:153,9]
IN36_30R_MLS=INOU36_30R[3:153,17]
IN36_30R_HLS=INOU36_30R[3:153,25]
OU36_30R_LLS=INOU36_30R[3:153,33]
OU36_30R_MLS=INOU36_30R[3:153,41]
OU36_30R_HLS=INOU36_30R[3:153,49]

IN36_31R_LLS=INOU36_31R[3:153,9]
IN36_31R_MLS=INOU36_31R[3:153,17]
IN36_31R_HLS=INOU36_31R[3:153,25]
OU36_31R_LLS=INOU36_31R[3:153,33]
OU36_31R_MLS=INOU36_31R[3:153,41]
OU36_31R_HLS=INOU36_31R[3:153,49]

IN36_32R_LLS=INOU36_32R[3:153,9]
IN36_32R_MLS=INOU36_32R[3:153,17]
IN36_32R_HLS=INOU36_32R[3:153,25]
OU36_32R_LLS=INOU36_32R[3:153,33]
OU36_32R_MLS=INOU36_32R[3:153,41]
OU36_32R_HLS=INOU36_32R[3:153,49]

IN36_33R_LLS=INOU36_33R[3:153,9]
IN36_33R_MLS=INOU36_33R[3:153,17]
IN36_33R_HLS=INOU36_33R[3:153,25]
OU36_33R_LLS=INOU36_33R[3:153,33]
OU36_33R_MLS=INOU36_33R[3:153,41]
OU36_33R_HLS=INOU36_33R[3:153,49]

IN36_34R_LLS=INOU36_34R[3:153,9]
IN36_34R_MLS=INOU36_34R[3:153,17]
IN36_34R_HLS=INOU36_34R[3:153,25]
OU36_34R_LLS=INOU36_34R[3:153,33]
OU36_34R_MLS=INOU36_34R[3:153,41]
OU36_34R_HLS=INOU36_34R[3:153,49]

IN36_35R_LLS=INOU36_35R[3:153,9]
IN36_35R_MLS=INOU36_35R[3:153,17]
IN36_35R_HLS=INOU36_35R[3:153,25]
OU36_35R_LLS=INOU36_35R[3:153,33]
OU36_35R_MLS=INOU36_35R[3:153,41]
OU36_35R_HLS=INOU36_35R[3:153,49]

IN36_36R_LLS=INOU36_36R[3:153,9]
IN36_36R_MLS=INOU36_36R[3:153,17]
IN36_36R_HLS=INOU36_36R[3:153,25]
OU36_36R_LLS=INOU36_36R[3:153,33]
OU36_36R_MLS=INOU36_36R[3:153,41]
OU36_36R_HLS=INOU36_36R[3:153,49]

IN108=cbind(IN36_1R_LLS,IN36_1R_MLS,IN36_1R_HLS,
            IN36_2R_LLS,IN36_2R_MLS,IN36_2R_HLS,
            IN36_3R_LLS,IN36_3R_MLS,IN36_3R_HLS,
            IN36_4R_LLS,IN36_4R_MLS,IN36_4R_HLS,
            IN36_5R_LLS,IN36_5R_MLS,IN36_5R_HLS,
            IN36_6R_LLS,IN36_6R_MLS,IN36_6R_HLS,
            IN36_7R_LLS,IN36_7R_MLS,IN36_7R_HLS,
            IN36_8R_LLS,IN36_8R_MLS,IN36_8R_HLS,
            IN36_9R_LLS,IN36_9R_MLS,IN36_9R_HLS,
            IN36_10R_LLS,IN36_10R_MLS,IN36_10R_HLS,
            IN36_11R_LLS,IN36_11R_MLS,IN36_11R_HLS,
            IN36_12R_LLS,IN36_12R_MLS,IN36_12R_HLS,
            IN36_13R_LLS,IN36_13R_MLS,IN36_13R_HLS,
            IN36_14R_LLS,IN36_14R_MLS,IN36_14R_HLS,
            IN36_15R_LLS,IN36_15R_MLS,IN36_15R_HLS,
            IN36_16R_LLS,IN36_16R_MLS,IN36_16R_HLS,
            IN36_17R_LLS,IN36_17R_MLS,IN36_17R_HLS,
            IN36_18R_LLS,IN36_18R_MLS,IN36_18R_HLS,
            IN36_19R_LLS,IN36_19R_MLS,IN36_19R_HLS,
            IN36_20R_LLS,IN36_20R_MLS,IN36_20R_HLS,
            IN36_21R_LLS,IN36_21R_MLS,IN36_21R_HLS,
            IN36_22R_LLS,IN36_22R_MLS,IN36_22R_HLS,
            IN36_23R_LLS,IN36_23R_MLS,IN36_23R_HLS,
            IN36_24R_LLS,IN36_24R_MLS,IN36_24R_HLS,
            IN36_25R_LLS,IN36_25R_MLS,IN36_25R_HLS,
            IN36_26R_LLS,IN36_26R_MLS,IN36_26R_HLS,
            IN36_27R_LLS,IN36_27R_MLS,IN36_27R_HLS,
            IN36_28R_LLS,IN36_28R_MLS,IN36_28R_HLS,
            IN36_29R_LLS,IN36_29R_MLS,IN36_29R_HLS,
            IN36_30R_LLS,IN36_30R_MLS,IN36_30R_HLS,
            IN36_31R_LLS,IN36_31R_MLS,IN36_31R_HLS,
            IN36_32R_LLS,IN36_32R_MLS,IN36_32R_HLS,
            IN36_33R_LLS,IN36_33R_MLS,IN36_33R_HLS,
            IN36_34R_LLS,IN36_34R_MLS,IN36_34R_HLS,
            IN36_35R_LLS,IN36_35R_MLS,IN36_35R_HLS,
            IN36_36R_LLS,IN36_36R_MLS,IN36_36R_HLS)

OU108=cbind(OU36_1R_LLS,OU36_1R_MLS,OU36_1R_HLS,
            OU36_2R_LLS,OU36_2R_MLS,OU36_2R_HLS,
            OU36_3R_LLS,OU36_3R_MLS,OU36_3R_HLS,
            OU36_4R_LLS,OU36_4R_MLS,OU36_4R_HLS,
            OU36_5R_LLS,OU36_5R_MLS,OU36_5R_HLS,
            OU36_6R_LLS,OU36_6R_MLS,OU36_6R_HLS,
            OU36_7R_LLS,OU36_7R_MLS,OU36_7R_HLS,
            OU36_8R_LLS,OU36_8R_MLS,OU36_8R_HLS,
            OU36_9R_LLS,OU36_9R_MLS,OU36_9R_HLS,
            OU36_10R_LLS,OU36_10R_MLS,OU36_10R_HLS,
            OU36_11R_LLS,OU36_11R_MLS,OU36_11R_HLS,
            OU36_12R_LLS,OU36_12R_MLS,OU36_12R_HLS,
            OU36_13R_LLS,OU36_13R_MLS,OU36_13R_HLS,
            OU36_14R_LLS,OU36_14R_MLS,OU36_14R_HLS,
            OU36_15R_LLS,OU36_15R_MLS,OU36_15R_HLS,
            OU36_16R_LLS,OU36_16R_MLS,OU36_16R_HLS,
            OU36_17R_LLS,OU36_17R_MLS,OU36_17R_HLS,
            OU36_18R_LLS,OU36_18R_MLS,OU36_18R_HLS,
            OU36_19R_LLS,OU36_19R_MLS,OU36_19R_HLS,
            OU36_20R_LLS,OU36_20R_MLS,OU36_20R_HLS,
            OU36_21R_LLS,OU36_21R_MLS,OU36_21R_HLS,
            OU36_22R_LLS,OU36_22R_MLS,OU36_22R_HLS,
            OU36_23R_LLS,OU36_23R_MLS,OU36_23R_HLS,
            OU36_24R_LLS,OU36_24R_MLS,OU36_24R_HLS,
            OU36_25R_LLS,OU36_25R_MLS,OU36_25R_HLS,
            OU36_26R_LLS,OU36_26R_MLS,OU36_26R_HLS,
            OU36_27R_LLS,OU36_27R_MLS,OU36_27R_HLS,
            OU36_28R_LLS,OU36_28R_MLS,OU36_28R_HLS,
            OU36_29R_LLS,OU36_29R_MLS,OU36_29R_HLS,
            OU36_30R_LLS,OU36_30R_MLS,OU36_30R_HLS,
            OU36_31R_LLS,OU36_31R_MLS,OU36_31R_HLS,
            OU36_32R_LLS,OU36_32R_MLS,OU36_32R_HLS,
            OU36_33R_LLS,OU36_33R_MLS,OU36_33R_HLS,
            OU36_34R_LLS,OU36_34R_MLS,OU36_34R_HLS,
            OU36_35R_LLS,OU36_35R_MLS,OU36_35R_HLS,
            OU36_36R_LLS,OU36_36R_MLS,OU36_36R_HLS)

write.csv(IN108,"OUT/IN108.csv")
write.csv(OU108,"OUT/OU108.csv")

#构造SPC矩阵
library(openxlsx)

setwd("E:/RBOOK/AL-高速度")
SPC=read.xlsx("CalResults_FSPC1%.xlsx",sheet="CalResults_FSPC1")
SPCR=SPC[,c(9,9,9,17,17,17,25,25,25,33,33,33,41,41,41, 49,49,49,57,57,57,65,65,65,73,73,73,81,81,81,89,89,89,97,97,97)]
SPCR=cbind(SPCR,SPCR,SPCR)
write.csv(SPCR,"OUT/SPCR108Matrix.csv")

#构造TS矩阵
TS=read.xlsx("CalResults_TS1横版%.xlsx",sheet = "CalResults_TS1横版")
TSR=data.frame(matrix(0,151,108))
for (i in 1:36){
  TSR[,(3*i-2):(3*i)]=cbind(TS[,i*8+1],TS[,i*8+1],TS[,i*8+1])
}
write.csv(TSR,"OUT/TSR108Matrix.csv")

#构造DS矩阵
#把DS1弄成横版
DS=read.csv("DS1%.csv")
DSR=data.frame(matrix(0,151,108))
for (i in 1:36){
  DSR[,(3*i-2):(3*i)]=cbind(rowSums(cbind(DS[,(7*i-5)],DS[,(7*i-4)],DS[,(7*i-3)],DS[,(7*i-2)],DS[,(7*i-1)],DS[,(7*i)],DS[,7*i+1])),
                            rowSums(cbind(DS[,(7*i-5)],DS[,(7*i-4)],DS[,(7*i-3)],DS[,(7*i-2)],DS[,(7*i-1)],DS[,(7*i)],DS[,7*i+1])),
                            rowSums(cbind(DS[,(7*i-5)],DS[,(7*i-4)],DS[,(7*i-3)],DS[,(7*i-2)],DS[,(7*i-1)],DS[,(7*i)],DS[,7*i+1])))
}
write.csv(DSR,"OUT/DSR108Matrix.csv")


#导入Loss Rate数
LR=read.xlsx("LR&NE.xlsx",sheet="LR")
OUR=read.xlsx("IN&OUT_SEC108%.xlsx",sheet="OU108")
#1_1构造加工损失率矩阵
LR_se_H=cbind(LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,
              LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,
              LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H,LR$semi_H)
LR_se_H=cbind(LR_se_H,LR_se_H,LR_se_H,LR_se_H,LR_se_H,LR_se_H)
LR_se_H=data.frame(LR_se_H)
#1_2构造使用损失率矩阵
LR_use_H=cbind(LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H,
              LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H,
              LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H,LR$use_H)
LR_use_H=cbind(LR_use_H,LR_use_H,LR_use_H,LR_use_H,LR_use_H,LR_use_H)
LR_use_H=data.frame(LR_use_H)
#1_3构造回收损失率矩阵
LR_re_H=cbind(LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,
              LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,
              LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H,LR$recycle_H)
LR_re_H=cbind(LR_re_H,LR_re_H,LR_re_H,LR_re_H,LR_re_H,LR_re_H)
LR_re_H=data.frame(LR_re_H)
#1_4构造预处理损失率矩阵
LR_prc_H=cbind(LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,
               LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,
               LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H,LR$prcs_H)
LR_prc_H=cbind(LR_prc_H,LR_prc_H,LR_prc_H,LR_prc_H,LR_prc_H,LR_prc_H)
LR_prc_H=data.frame(LR_prc_H)
#1_5构造重熔融损失率矩阵
LR_rlt_H=cbind(LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,
               LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,
               LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H,LR$rmlt_H)
LR_rlt_H=cbind(LR_rlt_H,LR_rlt_H,LR_rlt_H,LR_rlt_H,LR_rlt_H,LR_rlt_H)
LR_rlt_H=data.frame(LR_rlt_H)
#2_1构造加工损失率矩阵
LR_se_M=cbind(LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,
              LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,
              LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M,LR$semi_M)
LR_se_M=cbind(LR_se_M,LR_se_M,LR_se_M,LR_se_M,LR_se_M,LR_se_M)
LR_se_M=data.frame(LR_se_M)
#2_2构造使用损失率矩阵
LR_use_M=cbind(LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M,
               LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M,
               LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M,LR$use_M)
LR_use_M=cbind(LR_use_M,LR_use_M,LR_use_M,LR_use_M,LR_use_M,LR_use_M)
LR_use_M=data.frame(LR_use_M)
#构造回收损失率矩阵
LR_re_M=cbind(LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,
              LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,
              LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M,LR$recycle_M)
LR_re_M=cbind(LR_re_M,LR_re_M,LR_re_M,LR_re_M,LR_re_M,LR_re_M)
LR_re_M=data.frame(LR_re_M)
#构造预处理损失率矩阵
LR_prc_M=cbind(LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,
               LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,
               LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M,LR$prcs_M)
LR_prc_M=cbind(LR_prc_M,LR_prc_M,LR_prc_M,LR_prc_M,LR_prc_M,LR_prc_M)
LR_prc_M=data.frame(LR_prc_M)
#构造重熔融损失率矩阵
LR_rlt_M=cbind(LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,
               LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,
               LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M,LR$rmlt_M)
LR_rlt_M=cbind(LR_rlt_M,LR_rlt_M,LR_rlt_M,LR_rlt_M,LR_rlt_M,LR_rlt_M)
LR_rlt_M=data.frame(LR_rlt_M)
#3_1构造加工损失率矩阵
LR_se_L=cbind(LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,
              LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,
              LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L,LR$semi_L)
LR_se_L=cbind(LR_se_L,LR_se_L,LR_se_L,LR_se_L,LR_se_L,LR_se_L)
LR_se_L=data.frame(LR_se_L)
#3_2构造使用损失率矩阵
LR_use_L=cbind(LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L,
               LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L,
               LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L,LR$use_L)
LR_use_L=cbind(LR_use_L,LR_use_L,LR_use_L,LR_use_L,LR_use_L,LR_use_L)
LR_use_L=data.frame(LR_use_L)
#构造回收损失率矩阵
LR_re_L=cbind(LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,
              LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,
              LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L,LR$recycle_L)
LR_re_L=cbind(LR_re_L,LR_re_L,LR_re_L,LR_re_L,LR_re_L,LR_re_L)
LR_re_L=data.frame(LR_re_L)
#构造预处理损失率矩阵
LR_prc_L=cbind(LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,
               LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,
               LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L,LR$prcs_L)
LR_prc_L=cbind(LR_prc_L,LR_prc_L,LR_prc_L,LR_prc_L,LR_prc_L,LR_prc_L)
LR_prc_L=data.frame(LR_prc_L)
#构造重熔融损失率矩阵
LR_rlt_L=cbind(LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,
               LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,
               LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L,LR$rmlt_L)
LR_rlt_L=cbind(LR_rlt_L,LR_rlt_L,LR_rlt_L,LR_rlt_L,LR_rlt_L,LR_rlt_L)
LR_rlt_L=data.frame(LR_rlt_L)

NE4=read.xlsx("LR&NE.xlsx",sheet="NE4")
NE4_H=NE4$NE4_H
NE4_H=cbind(NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,
           NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,
           NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,
           NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H,NE4_H)
NE4_M=NE4$NE4_M
NE4_M=cbind(NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,
           NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,
           NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,
           NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M,NE4_M)
NE4_L=NE4$NE4_L
NE4_L=cbind(NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,
           NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,
           NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,
           NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L,NE4_L)

#高、中、低三种损失率下的再生铝生产潜力矩阵:全部年份为模拟
ONE=data.frame(matrix(1,151,108))
SAP_HH=(OUR-NE4_H)*(ONE-LR_re_H)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_HM=(OUR-NE4_M)*(ONE-LR_re_H)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_HL=(OUR-NE4_L)*(ONE-LR_re_H)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_MH=(OUR-NE4_H)*(ONE-LR_re_M)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_MM=(OUR-NE4_M)*(ONE-LR_re_M)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_ML=(OUR-NE4_L)*(ONE-LR_re_M)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_LH=(OUR-NE4_H)*(ONE-LR_re_L)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
SAP_LM=(OUR-NE4_M)*(ONE-LR_re_L)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
SAP_LL=(OUR-NE4_L)*(ONE-LR_re_L)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
write.csv(SAP_HH,"OUT/SAP_HH.csv")
write.csv(SAP_HM,"OUT/SAP_HM.csv")
write.csv(SAP_HL,"OUT/SAP_HL.csv")
write.csv(SAP_MH,"OUT/SAP_MH.csv")
write.csv(SAP_MM,"OUT/SAP_MM.csv")
write.csv(SAP_ML,"OUT/SAP_ML.csv")
write.csv(SAP_HL,"OUT/SAP_HL.csv")
write.csv(SAP_ML,"OUT/SAP_ML.csv")
write.csv(SAP_LL,"OUT/SAP_LL.csv")
#高、中、低三种损失率下的再生铝生产潜力矩阵:历史年份为历史值
OUR_hst=read.xlsx("IN&OUT_SEC108%.xlsx",sheet="OU+History")
INR_hst=read.xlsx("IN&OUT_SEC108%.xlsx",sheet="IN+History")
ONE=data.frame(matrix(1,151,108))
SAP_hst_HH=(OUR_hst*(ONE-LR_re_H)-NE4_H)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_hst_HM=(OUR_hst*(ONE-LR_re_H)-NE4_M)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_hst_HL=(OUR_hst*(ONE-LR_re_H)-NE4_L)*(ONE-LR_prc_H)*(ONE-LR_rlt_H)
SAP_hst_MH=(OUR_hst*(ONE-LR_re_M)-NE4_H)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_hst_MM=(OUR_hst*(ONE-LR_re_M)-NE4_M)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_hst_ML=(OUR_hst*(ONE-LR_re_M)-NE4_L)*(ONE-LR_prc_M)*(ONE-LR_rlt_M)
SAP_hst_LH=(OUR_hst*(ONE-LR_re_L)-NE4_H)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
SAP_hst_LM=(OUR_hst*(ONE-LR_re_L)-NE4_M)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
SAP_hst_LL=(OUR_hst*(ONE-LR_re_L)-NE4_L)*(ONE-LR_prc_L)*(ONE-LR_rlt_L)
write.csv(SAP_hst_HH,"OUT/SAP_hst_HH.csv")
write.csv(SAP_hst_HM,"OUT/SAP_hst_HM.csv")
write.csv(SAP_hst_HL,"OUT/SAP_hst_HL.csv")
write.csv(SAP_hst_MH,"OUT/SAP_hst_MH.csv")
write.csv(SAP_hst_MM,"OUT/SAP_hst_MM.csv")
write.csv(SAP_hst_ML,"OUT/SAP_hst_ML.csv")
write.csv(SAP_hst_HL,"OUT/SAP_hst_LH.csv")
write.csv(SAP_hst_ML,"OUT/SAP_hst_LM.csv")
write.csv(SAP_hst_LL,"OUT/SAP_hst_LL.csv")

#导入净出口数据
NE=read.xlsx("LR&NE.xlsx",sheet="NE")
NE_H=NE$NE_H
NE_H=cbind(NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,
           NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,
           NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,
           NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H,NE_H)
NE_M=NE$NE_M
NE_M=cbind(NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,
           NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,
           NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,
           NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M,NE_M)
NE_L=NE$NE_L
NE_L=cbind(NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,
           NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,
           NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,
           NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L,NE_L)


NE3=read.xlsx("LR&NE.xlsx",sheet="NE3")
NE3_H=NE3$NE3_H
NE3_H=cbind(NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,
            NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,
            NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,
            NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H,NE3_H)
NE3_M=NE3$NE3_M
NE3_M=cbind(NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,
            NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,
            NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,
            NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M,NE3_M)
NE3_L=NE3$NE3_L
NE3_L=cbind(NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,
            NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,
            NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,
            NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L,NE3_L)

NE2=read.xlsx("LR&NE.xlsx",sheet="NE2")
NE2_H=NE2$NE2_H
NE2_H=cbind(NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,
            NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,
            NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,
            NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H,NE2_H)
NE2_M=NE2$NE2_M
NE2_M=cbind(NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,
            NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,
            NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,
            NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M,NE2_M)
NE2_L=NE2$NE2_L
NE2_L=cbind(NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,
            NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,
            NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,
            NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L,NE2_L)

NE1=read.xlsx("LR&NE.xlsx",sheet="NE1")
NE1_H=NE1$NE1_H
NE1_H=cbind(NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,
            NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,
            NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,
            NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H,NE1_H)
NE1_M=NE1$NE1_M
NE1_M=cbind(NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,
            NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,
            NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,
            NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M,NE1_M)
NE1_L=NE1$NE1_L
NE1_L=cbind(NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,
            NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,
            NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,
            NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L,NE1_L)

#计算总损失量
Tloss_HH=(INR_hst+NE3_H+NE2_H)*(LR_se_H/(1-LR_se_H))+LR_use_H+OUR_hst*LR_re_H+(OUR_hst*LR_re_H - NE4_H)*(1-LR_re_H)*LR_prc_H+(OUR_hst*LR_re_H - NE4_H)*(1-LR_re_H)*(1-LR_prc_H)*LR_rlt_H
Tloss_MH=(INR_hst+NE3_H+NE2_H)*(LR_se_M/(1-LR_se_M))+LR_use_M+OUR_hst*LR_re_M+(OUR_hst*LR_re_M - NE4_H)*(1-LR_re_M)*LR_prc_M+(OUR_hst*LR_re_M - NE4_H)*(1-LR_re_M)*(1-LR_prc_M)*LR_rlt_M
Tloss_LH=(INR_hst+NE3_H+NE2_H)*(LR_se_L/(1-LR_se_L))+LR_use_L+OUR_hst*LR_re_L+(OUR_hst*LR_re_L - NE4_H)*(1-LR_re_L)*LR_prc_L+(OUR_hst*LR_re_L - NE4_H)*(1-LR_re_L)*(1-LR_prc_L)*LR_rlt_L
Tloss_HM=(INR_hst+NE3_M+NE2_M)*(LR_se_H/(1-LR_se_H))+LR_use_H+OUR_hst*LR_re_H+(OUR_hst*LR_re_H - NE4_M)*(1-LR_re_H)*LR_prc_H+(OUR_hst*LR_re_H - NE4_M)*(1-LR_re_H)*(1-LR_prc_H)*LR_rlt_H
Tloss_MM=(INR_hst+NE3_M+NE2_M)*(LR_se_M/(1-LR_se_M))+LR_use_M+OUR_hst*LR_re_M+(OUR_hst*LR_re_M - NE4_M)*(1-LR_re_M)*LR_prc_M+(OUR_hst*LR_re_M - NE4_M)*(1-LR_re_M)*(1-LR_prc_M)*LR_rlt_M
Tloss_LM=(INR_hst+NE3_M+NE2_M)*(LR_se_L/(1-LR_se_L))+LR_use_L+OUR_hst*LR_re_L+(OUR_hst*LR_re_L - NE4_M)*(1-LR_re_L)*LR_prc_L+(OUR_hst*LR_re_L - NE4_M)*(1-LR_re_L)*(1-LR_prc_L)*LR_rlt_L
Tloss_HL=(INR_hst+NE3_L+NE2_L)*(LR_se_H/(1-LR_se_H))+LR_use_H+OUR_hst*LR_re_H+(OUR_hst*LR_re_H - NE4_L)*(1-LR_re_H)*LR_prc_H+(OUR_hst*LR_re_H - NE4_L)*(1-LR_re_H)*(1-LR_prc_H)*LR_rlt_H
Tloss_ML=(INR_hst+NE3_L+NE2_L)*(LR_se_M/(1-LR_se_M))+LR_use_M+OUR_hst*LR_re_M+(OUR_hst*LR_re_M - NE4_L)*(1-LR_re_M)*LR_prc_M+(OUR_hst*LR_re_M - NE4_L)*(1-LR_re_M)*(1-LR_prc_M)*LR_rlt_M
Tloss_LL=(INR_hst+NE3_L+NE2_L)*(LR_se_L/(1-LR_se_L))+LR_use_L+OUR_hst*LR_re_L+(OUR_hst*LR_re_L - NE4_L)*(1-LR_re_L)*LR_prc_L+(OUR_hst*LR_re_L - NE4_L)*(1-LR_re_L)*(1-LR_prc_L)*LR_rlt_L

#导入净增量
DSR_hst=read.xlsx("DSR108Matrix%.xlsx",sheet = "DS+History")
DSR_hst=DSR_hst[,2:109]
#计算总铝需求量
INR_hst=read.xlsx("IN&OUT_SEC108%.xlsx",sheet="IN+History")
TAP_HH=(INR_hst+NE3_H+NE2_H)/(1-LR_se_H)+NE1_H
TAP_HM=(INR_hst+NE3_M+NE2_M)/(1-LR_se_H)+NE1_M
TAP_HL=(INR_hst+NE3_L+NE2_L)/(1-LR_se_H)+NE1_L
TAP_MH=(INR_hst+NE3_H+NE2_H)/(1-LR_se_M)+NE1_H
TAP_MM=(INR_hst+NE3_M+NE2_M)/(1-LR_se_M)+NE1_M
TAP_ML=(INR_hst+NE3_L+NE2_L)/(1-LR_se_M)+NE1_L
TAP_LH=(INR_hst+NE3_H+NE2_H)/(1-LR_se_L)+NE1_H
TAP_LM=(INR_hst+NE3_M+NE2_M)/(1-LR_se_L)+NE1_M
TAP_LL=(INR_hst+NE3_L+NE2_L)/(1-LR_se_L)+NE1_L
#shuchu
write.csv(TAP_HH,"OUT/TAP_HH.csv")
write.csv(TAP_HM,"OUT/TAP_HM.csv")
write.csv(TAP_HL,"OUT/TAP_HL.csv")
write.csv(TAP_MH,"OUT/TAP_MH.csv")
write.csv(TAP_MM,"OUT/TAP_MM.csv")
write.csv(TAP_ML,"OUT/TAP_ML.csv")
write.csv(TAP_LH,"OUT/TAP_LH.csv")
write.csv(TAP_LM,"OUT/TAP_LM.csv")
write.csv(TAP_LL,"OUT/TAP_LL.csv")
#计算原生铝需求量
PAP_HH=TAP_HH-SAP_hst_HH
PAP_HM=TAP_HM-SAP_hst_HM
PAP_HL=TAP_HL-SAP_hst_HL
PAP_MH=TAP_MH-SAP_hst_MH
PAP_MM=TAP_MM-SAP_hst_MM
PAP_ML=TAP_ML-SAP_hst_ML
PAP_LH=TAP_LH-SAP_hst_LH
PAP_LM=TAP_LM-SAP_hst_LM
PAP_LL=TAP_LL-SAP_hst_LL
#第一种输出方法
write.csv(PAP_HH,"OUT/PAP_HH.csv")
write.csv(PAP_HM,"OUT/PAP_HM.csv")
write.csv(PAP_HL,"OUT/PAP_HL.csv")
write.csv(PAP_MH,"OUT/PAP_MH.csv")
write.csv(PAP_MM,"OUT/PAP_MM.csv")
write.csv(PAP_ML,"OUT/PAP_ML.csv")
write.csv(PAP_LH,"OUT/PAP_LH.csv")
write.csv(PAP_LM,"OUT/PAP_LM.csv")
write.csv(PAP_LL,"OUT/PAP_LL.csv")

#第二种输出方法
PAP=data.frame(rbind(PAP_HH,PAP_HM,PAP_HL,
              PAP_MH,PAP_MM,PAP_ML,
              PAP_LH,PAP_LM,PAP_LL))
write.csv(PAP,"OUT/PAP9.csv")

