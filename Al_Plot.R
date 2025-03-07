#AL未来存量流量画图
library(ggplot2) #作图包
library(dplyr) #数据转换包

library(tidyr) #数据转换包
library(splines) #数据差值包
library(Hmisc)#用来绘制次要刻度线
library(grid)#用来排版
library(gridExtra)#用来排版
library(cowplot)#排版
library(ggpubr)#排版
library(RColorBrewer)
library(openxlsx)

setwd("E:/RBOOK/AL")
SPC_PD=read.xlsx("PLOT/Stocks&Flows_Hst.xlsx",sheet = "SPC")
TS_PD=read.xlsx("PLOT/Stocks&Flows_Hst.xlsx",sheet = "TS")
IN_PD=read.xlsx("PLOT/Stocks&Flows_Hst.xlsx",sheet = "IN")
OU_PD=read.xlsx("PLOT/Stocks&Flows_Hst.xlsx",sheet = "OU")
SPC_PD_L=gather(SPC_PD,Group,Content,-Time)#转换数据，宽变长
TS_PD_L=gather(TS_PD,Group,Content,-Time)#转换数据，宽变长
IN_PD_L=gather(IN_PD,Group,Content,-Time)#转换数据，宽变长
OU_PD_L=gather(OU_PD,Group,Content,-Time)#转换数据，宽变长
Item1=data.frame(matrix("SPC",16308,1))
Item2=data.frame(matrix("TS",16308,1))
Item3=data.frame(matrix("IN",16308,1))
Item4=data.frame(matrix("OU",16308,1))
write.csv(Item1,"PLOT/Item1.csv")
write.csv(Item2,"PLOT/Item2.csv")
write.csv(Item3,"PLOT/Item3.csv")
write.csv(Item4,"PLOT/Item4.csv")
Item=read.csv("PLOT/Item.csv",header = FALSE)
SAF=cbind(rbind(SPC_PD_L,TS_PD_L,IN_PD_L,OU_PD_L),Item)
ggplot(SAF,aes(Time,Content,colour=V1))+geom_line()+ facet_wrap(~Group, scales="free") + ggtitle('scales="free"')#分面
