import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def gmail(subject,content,attach):

    html = content
    msg = MIMEMultipart()                         # 使用多種格式所組成的內容
    msg.attach(MIMEText(html, 'html', 'utf-8'))   # 加入 HTML 內容
    # 使用 python 內建的 open 方法開啟指定目錄下的檔案
    with open(attach, 'rb') as file:
        img = file.read()
    attach_file = MIMEApplication(img, Name='log.txt')    # 設定附加檔案圖片
    msg.attach(attach_file)                       # 加入附加檔案圖片

    msg['Subject']=subject
    msg['From']=''#寄件人
    msg['To']=''#收件人
    smtp = smtplib.SMTP('smtp.gmail.com', 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.login('','')#寄件人信箱,寄件人信箱金鑰
    status = smtp.send_message(msg)
    print(status)
    if status == {}:
        print('郵件傳送成功！')
    else:
        print('郵件傳送失敗...')
    smtp.quit()
