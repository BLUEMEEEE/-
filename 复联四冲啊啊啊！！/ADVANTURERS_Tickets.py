import urllib.request
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.header import Header
import time

username='XXXXX@qq.com'
password='XXXXX'
sender=username
receiver=username
msg_str="Tickets are On Sale!!!! GO! GO! GO!!!!"

# 发邮件给自己
def send_email():
    status=True
    try:
        msg=MIMEText(msg_str,'utf-8')
        msg['From']=formataddr(["ADVANTURERS!!!!",sender])
        msg['To']=formataddr(["ZYQ",receiver])
        msg['Subject']=msg_str
        

        server=smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(username, password)
        server.sendmail(sender,[receiver,],msg.as_string())
        server.quit()
    except smtplib.SMTPException as e:
        print(str(e))
        status=False
    return status

count=1
i=0
while i==0:
    response = urllib.request.urlopen('https://maoyan.com/cinemas?movieId=248172&showDate=2019-04-25&brandId=102642')
    result = response.read().decode('utf-8')
    index=result.find('万达影城(万胜围店)')
    res=True if index!=-1 else False
    print("**********"+str(count)+"**********"+":\nTicket_Sale_Status:"+str(res))
    count=count+1
    if(res):
        send_email()
        break
    time.sleep(15)
    