import urllib.request
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.header import Header
import time
from bs4 import BeautifulSoup

username='XXXXXX@qq.com'
password='XXXXXXX'
sender=username
receiver='XXXXXXX@qq.com'
msg_str="Tickets are On Sale!!!! GO! GO! GO!!!!"

def send_email(receiver_addr):
    status=True
    try:
        msg=MIMEText(msg_str,'utf-8')
        msg['From']=formataddr(["ADVANTURERS!!!!",sender])
        msg['To']=formataddr(["ZYQ",receiver_addr])
        msg['Subject']=msg_str
        

        server=smtplib.SMTP_SSL("smtp.qq.com", 465)
        server.login(username, password)
        server.sendmail(sender,[receiver_addr,],msg.as_string())
        server.quit()
    except smtplib.SMTPException as e:
        print(str(e))
        status=False
    return status

def check_time(start_time):
    h=int(start_time.split(':')[0])
    if h>=16 and h<=20:
        return True
    return False

count=1
while True:
    try:
        response = urllib.request.urlopen('https://maoyan.com/cinema/16769?poi=150027234&movieId=248172')
        html_doc = response.read().decode('utf-8')
        soup = BeautifulSoup(html_doc, 'html.parser')
        #print(soup.prettify())
        h3s=soup.find_all('h3',class_='movie-name')
        for h3 in h3s:
            if(h3.contents[0]=='复仇者联盟4：终局之战'):
                div=h3.parent.parent.parent#.next_sibling.next_sibling.next_sibling
                break
                #print(str(div))
        search_part=BeautifulSoup(str(div), 'html.parser')
        tables=search_part.find_all('table')
        table=tables[1]
        for item in table.tbody.contents:
            if(str(item).find('tr')==-1):
                continue
            start_time=item.td.span.contents[0]
            res=check_time(start_time)
            if(res==True):
                break
        #index=html_doc.find('万达影城(万胜围店)')
        #res=True if index!=-1 else False
        print("**********"+time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))+"**********"+"\nTicket_Sale_Status:"+str(res))
        if res:
            send_email(username)
            break
        count=count+1
        #TODO:发现sleep并不能准确睡眠60s，有时候睡眠时间会延长
        time.sleep(60)
    except Exception as e:
        print(str(e))
        time.sleep(120)
    
