import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


from email.utils import formataddr
# 第三方 SMTP 服务
mail_host = "smtp.qq.com"  # 设置服务器
mail_user = "1930702390@qq.com"  # 用户名
mail_pass = "**********"  # 口令（发件人授权码）

sender = '1930702390@qq.com'
receivers = ['1875408506@qq.com', 'zhangly@zhchuangsi.com']  # 接收邮件人
content = MIMEText('公司打卡记录')
message = MIMEMultipart()
message.attach(content)
message['From'] = formataddr(["赖", sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
message['To'] = formataddr(["张", mail_user])  # 括号里的对应收件人人邮箱昵称、发件人邮箱账号
subject = 'Python SMTP 邮件测试'
message['Subject'] = Header(subject, 'utf-8')
xlsx = MIMEApplication(open('打卡记录.xlsx', 'rb').read())
xlsx["Content-Type"] = 'application/octet-stream'
xlsx.add_header('Content-Disposition', 'attachment', filename='打卡记录.xlsx')
message.attach(xlsx)
try:
    # 一般流程：链接服务器-登陆账号-发送信息-关闭
    smtpObj = smtplib.SMTP()
    smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
    smtpObj.login(mail_user, mail_pass)
    smtpObj.sendmail(sender, receivers, message.as_string())
    print("邮件发送成功")
except  Exception as e:
    print(e)
