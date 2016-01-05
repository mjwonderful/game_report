########################################################################
###########################��Ϸ��ҵ������############################### 
########################################################################

###�������ݿⲢ��ȡ����###
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
users <- fetch(res, n = -1) # ��ѯ���
users <- na.omit(users)



###��ͼ###
library(ggplot2)

y1_min <- range(users$new_users)[1]
y1_max <- range(users$new_users)[2]

  # ���ݱ�ǩ��ȡ�����Сֵ
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
 
##��������
plot1 <- ggplot(users,aes(date,new_users))+geom_bar(stat="identity",fill=I("#458B74"))+
  coord_cartesian(ylim=c(y1_min-10000,y1_max+20000))+
  labs(x=NULL,y="�����û���",title="�����û�����ͼ")+
  geom_text(aes(label =new_users_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE)  
ggsave(my_temp_file<-paste(tempfile(),".wmf",sep=""), plot=plot1) #Save the chart in Windows Metafile Format to a temporary file

##��Ծ����
plot2 <- ggplot(users,aes(date,active_users))+geom_bar(stat="identity",fill=I("#458B74"))+
  coord_cartesian(ylim=c(0,200000))+
  labs(x=NULL,y="��Ծ�û���",title="��Ծ�û�����ͼ")+
  geom_text(aes(label =active_users_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE)  
ggsave(file="plot2.jpeg",width=5,height=5)

##��������
plot3 <- ggplot(users,aes(date,income))+geom_point()+geom_line()+ 
  coord_cartesian(ylim=c(5000,40000))+
  labs(x=NULL,y="����",title="��������ͼ")+
  geom_text(aes(label =income_range, vjust =-1, hjust =0.6, size=0.5, angle = 0), show_guide = FALSE) 
ggsave(file="plot3.png",width=5,height=5)

##���ͼƬ
require(gridExtra)
grid.arrange(plot1,plot2,plot3,ncol=1)  

 
dbClearResult(rs) # ������
dbDisconnect(con) # �Ͽ�����
dbUnloadDriver(drv) # �ͷ���Դ

###д��PPT###
require(R2PPT) 
myPres <- PPT.Init(method="RDCOMClient") #����
cur_date <- Sys.Date()
myPres <- PPT.AddTitleSlide(myPres, title = "��Ϸ��ҵ���ձ�", subtitle = cur_date ) #����
myPres <- PPT.AddTextSlide(myPres, title = "�������ݸ���",text="��������\rͼ��\r����") #����
 

# ����
myPres <- PPT.AddTitleOnlySlide(myPres, title = "����")
myPres <- PPT.AddDataFrame(myPres, df = head(users),
                           row.names=FALSE, size=c(55, 150, 600, 300))
 
# ͼ��
myPres <- PPT.AddTitleOnlySlide(myPres, title = "ͼ��")
mypres <- PPT.AddGraphicstoSlide(myPres,file=my_temp_file,size= c(10, 10, 300, 500))
unlink(my_temp_file) #Delete the temporary file 


# ���沢�رս���
myPres<-PPT.SaveAs(myPres,file=paste(getwd(),"game_report.ppt",sep="/")) #ֻ֧��R,��֧��RStudio
myPres<-PPT.Close(myPres)
rm(myPres)

 
###�ʼ�����###
require("mailR")

inc <- paste("��������",tail(users,1)[5],";")
new_u <- paste("��������",tail(users,1)[2],";")
act_u <-  paste("���ջ�Ծ",tail(users,1)[3],";")


send.mail(from = "from@qq.com",
          to = "to@126.com",
          subject =paste("��Ϸ��ҵ���ձ�",cur_date,sep="����"), 
          body = paste(inc, new_u, act_u),
          encoding = "utf-8",
          smtp = list(host.name = "smtp.qq.com", port = 465, user.name = "user.name",
                      passwd = "passwd", ssl = TRUE), authenticate = TRUE, send = TRUE,
          attach.files= paste(getwd(),"game_report.ppt",sep="/")  #���Ӹ���
          )


###��ע###
###��R��RStudio�±���ͼƬ��������
###��R��ֱ�ӱ���ggsave(file="plot3.png",width=5,height=5)����ppt����ʱֱ�����������ļ�·�����ɣ�
###��RStudio��ͼƬ�����Windows�г�����һ��ͼԪ�ļ���ʽggsave(my_temp_file1<-paste(tempfile(),".wmf",sep=""), plot=plot1)�����ܲ���ppt��
###����pptֻ����R�½��У�������RStudio�н��У����������ʱû�н����



