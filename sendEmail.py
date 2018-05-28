#!/usr/bin/env python  
# _*_ coding:utf-8 _*_  
# @Author  : lusheng

import smtplib
from email.mime.text import MIMEText
from email.header import Header
import xlrd

path = 'C:\\Users\\LUS\\Desktop\\email.xlsx'
data = xlrd.open_workbook(path)
sheets=data.sheets()
sheet_1_by_function=data.sheets()[0]
sheet_1_by_index=data.sheet_by_index(0)
sheet_1_by_name=data.sheet_by_name(u'Sheet1')
n_of_rows=sheet_1_by_name.nrows
n_of_cols=sheet_1_by_name.ncols
for i in range(n_of_rows):
    print(sheet_1_by_name.row_values(i))


# 第三方 SMTP 服务
mail_host="smtp.sinometalsh.com"  #设置服务器
mail_user="lusheng@sinometalsh.com"    #用户名
mail_pass="Lu1986617"   #口令

sender = 'lusheng@sinometalsh.com'
# nameList = ['one', 'two']
# receiversList = {'one': '228383562@qq.com',
#                  'two': 'lusheng1234@126.com'}
for i in range(1, n_of_rows):
    receiversName = sheet_1_by_name.row_values(i)[0]
    receivers = sheet_1_by_name.row_values(i)[1]
    print(receiversName, receivers)
    if sheet_1_by_name.row_values(i)[2] == '男':
        receiversName2 = receiversName[0] + '先生'
    else:
        receiversName2 = receiversName[0] + '女士'
    print(receiversName2, receivers)
# receivers = ['228383562@qq.com', 'lusheng1234@126.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

# 三个参数：第一个为文本内容，第二个 plain 设置文本格式，第三个 utf-8 设置编码
#     message = MIMEText('Python 邮件发送测试... %s ...' % receiversName, 'plain', 'utf-8')
#     message['From'] = Header("菜鸟教程", 'utf-8')  # 发送者
#     message['To'] = Header("测试", 'utf-8')  # 接收者

    mail_msg = """
    <p> %s: </p>
    <p>这是Python 邮件发送测试... </p>
    <p><a href="http://www.baidu.com">这是一个链接</a></p>
    """ % receiversName2
    message = MIMEText(mail_msg, 'html', 'utf-8')
    message['From'] = Header("邮件通知", 'utf-8')
    message['To'] = Header("邮件测试", 'utf-8')



    subject = 'Python SMTP 邮件测试'
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)    # 25 为 SMTP 端口号
        smtpObj.login(mail_user,mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("邮件发送成功")
    except smtplib.SMTPException:
        print("Error: 无法发送邮件")