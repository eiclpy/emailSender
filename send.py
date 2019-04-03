# -+-coding=utf8-+-
import lmail
import argparse
import sys


# ===========================SETTING=============================
addr = lmail.default_sender
passwd = lmail.default_password
excel = lmail.excel()
excel.add('./emails/example.xlsx')#可选参数：指定文件中指定表如 excel.add(filename,'1-3,5,6')
emails = excel.emails
print('共{}人'.format(len(emails)))


content, attachment = lmail.get_file('./source/')#发送正文及附件所在的目录，正文以txt格式存

mail = lmail.Mail()
mail.set_sender('Nickname', addr)#发件昵称
mail.set_subject('Subject')  # 邮件标题
mail.set_content(content)
mail.add_attachment(attachment)

# ===============================================================
parser = argparse.ArgumentParser()
parser.add_argument('-cm', '--checkmyself',
                    help='发送给自己', action='store_true')
parser.add_argument('-sa', '--sendall',
                    help='发送给所有人', action='store_true')
parser.add_argument('-nd', '--nodebug',
                    help='关闭调试模式', action='store_true')
args = parser.parse_args()

server = lmail.Server(addr, passwd, debug=False if args.nodebug else True)
# ===============================================================

if args.checkmyself:
    emails = ['your email address']
    server.send_all_mails(emails, mail)

elif args.sendall:
    input('WARNING: Really send all?')
    server.send_all_mails(emails, mail)

else:
    print('please use -h to get the help!')
