########################################################################
###########################游戏事业部报表############################### 
########################################################################

###链接数据库并获取数据###
library(RPostgreSQL)
drv <- dbDriver("PostgreSQL") 
con <- dbConnect(drv, host='host', port='port', dbname='dbname',user="user",password="password")
res <- dbSendQuery(con, statement = 
       "select date,  
               sum(new_users) as new_users, 
               sum(active_users) as active_users, 
               sum(pay_users) as pay_users, 
               sum(income) as income 
        from ims.idt_mgp_android_all 
        where date>= current_date - interval '30 day'  
        group by date order by date          
        ")
users <- fetch(res, n = -1) # 查询结果
users <- na.omit(users)



###绘图###
library(ggplot2)

y1_min <- range(users$new_users)[1]
y1_max <- range(users$new_users)[2]

  # 数据标签获取最大最小值
result <- ""
fun.range <- function (x) {
  for (i in (1:length(x)))
    if(x[i]==range(x)[1]){
      result[i] <- range(x)[1] 
    } 
  else if(x[i]==range(x)[2]){
    result[i] <- range(x)[2]
  }
  else {
    result[i] <- NA
  }
  return(result)
}

new_users_range <- fun.range(users$new_users)
active_users_range <- fun.range(users$active_users)
income_range <- fun.range(users$income)
 
##新增趋势
plot1 <- ggplot(users,aes(date,new_users))+geom_bar(stat="identity",fill=I("#458B74"))+
  coord_cartesian(ylim=c(y1_min-10000,y1_max+20000))+
  labs(x=NULL,y="新增用户数",title="新增用户趋势图")+
  geom_text(aes(label =new_users_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE)  
ggsave(my_temp_file<-paste(tempfile(),".wmf",sep=""), plot=plot1) #Save the chart in Windows Metafile Format to a temporary file

##活跃趋势
plot2 <- ggplot(users,aes(date,active_users))+geom_bar(stat="identity",fill=I("#458B74"))+
  coord_cartesian(ylim=c(0,200000))+
  labs(x=NULL,y="活跃用户数",title="活跃用户趋势图")+
  geom_text(aes(label =active_users_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE)  
ggsave(file="plot2.jpeg",width=5,height=5)

##收入趋势
plot3 <- ggplot(users,aes(date,income))+geom_point()+geom_line()+ 
  coord_cartesian(ylim=c(5000,40000))+
  labs(x=NULL,y="收入",title="收入趋势图")+
  geom_text(aes(label =income_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE) 
ggsave(file="plot3.png",width=5,height=5)

##组合图片
require(gridExtra)
grid.arrange(plot1,plot2,plot3,ncol=1)  

 
dbClearResult(rs) # 清除结果
dbDisconnect(con) # 断开连接
dbUnloadDriver(drv) # 释放资源

###写入PPT###
require(R2PPT) 
myPres <- PPT.Init(method="RDCOMClient") #创建
cur_date <- Sys.Date()
myPres <- PPT.AddTitleSlide(myPres, title = "游戏事业部日报", subtitle = cur_date ) #封面
myPres <- PPT.AddTextSlide(myPres, title = "昨日数据概览",text="文字区域\r图形\r表格") #文字
 

# 表格
myPres <- PPT.AddTitleOnlySlide(myPres, title = "表格")
myPres <- PPT.AddDataFrame(myPres, df = head(users),
                           row.names=FALSE, size=c(55, 150, 600, 300))
 
# 图形
myPres <- PPT.AddTitleOnlySlide(myPres, title = "图形")
mypres <- PPT.AddGraphicstoSlide(myPres,file=my_temp_file,size= c(10, 10, 300, 500))
unlink(my_temp_file) #Delete the temporary file 


# 保存并关闭进程
myPres<-PPT.SaveAs(myPres,file=paste(getwd(),"game_report.ppt",sep="/")) #只支持R,不支持RStudio
myPres<-PPT.Close(myPres)
rm(myPres)

 
###邮件发送###
require("mailR")

inc <- paste("昨日收入",tail(users,1)[5],";")
new_u <- paste("昨日新增",tail(users,1)[2],";")
act_u <-  paste("昨日活跃",tail(users,1)[3],";")


send.mail(from = "from@qq.com",
          to = "to@126.com",
          subject =paste("游戏事业部日报",cur_date,sep="——"), 
          body = paste(inc, new_u, act_u),
          encoding = "utf-8",
          smtp = list(host.name = "smtp.qq.com", port = 465, user.name = "user.name",
                      passwd = "passwd", ssl = TRUE), authenticate = TRUE, send = TRUE,
          attach.files= paste(getwd(),"game_report.ppt",sep="/")  #添加附件
          )


###备注###
###在R和RStudio下保存图片略有区别：
###在R下直接保存ggsave(file="plot3.png",width=5,height=5)，往ppt插入时直接引用完整文件路径即可；
###在RStudio下图片保存成Windows中常见的一种图元文件格式ggsave(my_temp_file1<-paste(tempfile(),".wmf",sep=""), plot=plot1)，才能插入ppt；
###保存ppt只能在R下进行，不能再RStudio中进行，这个问题暂时没有解决；




