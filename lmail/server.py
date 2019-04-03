# -+- coding=utf8 -+-
import copy
import logging
import smtplib
import sys
import time
import random
from email.utils import COMMASPACE

from .config import *

logger = logging.getLogger(log_name)
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler(log_path)
fh.setLevel(logging.DEBUG)
fmt = logging.Formatter('[%(levelname)s] %(message)s')
fh.setFormatter(fmt)
logger.addHandler(fh)


class Server():
    def __init__(self, username: str = default_sender, password: str = default_password, host: str = default_host, port: int = default_port, debug: bool = True):
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.debug = debug

        self.fail_send = []
        self.all_send = []
        self.total_send = 0

        self._login = False
        self._needstls = True
        logger.debug(self.__repr__())

    def stls(self):
        """Start TLS."""
        self.server.ehlo()
        self.server.starttls()
        self.server.ehlo()
        logger.debug('stls successful')

    def log_error(self):
        if self.fail_send:
            print('存在发送失败!!!')
            logger.critical('ERROR!!!')
            if self.fail_send:
                print(self.fail_send)
                logger.critical(' '.join(self.fail_send))

    def save_last_successful_send(self):
        with open(time.strftime('last_send_%d_%M_%S.txt'), 'w') as f:
            print(
                ' '.join([x for x in self.all_send if not x in self.fail_send]), file=f)

    def login(self):
        """Login"""
        try:
            self.server = smtplib.SMTP(self.host, self.port)
        except:
            print('无法连接服务器，请检查网络是否连通')
            self.log_error()
            self.save_last_successful_send()
            sys.exit(0)
        if self.debug:
            self._login = True
            self._needstls = False
            logger.debug('Login successful (DEBUG MODE)')
            return

        if self._login:
            logger.warning('{} duplicate login!'.format(self.__repr__()))
            return

        if self._needstls:
            self.stls()
            self._needstls = False

        try:
            self.server.login(self.username, self.password)
        except smtplib.SMTPAuthenticationError:
            print('服务器认证错误，请检查用户名密码')
            logger.error('username or password error')
            self.log_error()
            self.save_last_successful_send()
            sys.exit(0)

        self._login = True
        logger.debug('Login successful')

    def logout(self):
        """Logout"""
        if self.debug:
            self._login = False
            self._needstls = True
            self.server = None
            logger.debug('Logout successful (DEBUG MODE)')
            return

        if not self._login:
            logger.warning('{} Logout before login!'.format(self.__repr__()))
            return

        try:
            code, message = self.server.docmd("QUIT")
            if code != 221:
                raise smtplib.SMTPResponseException(code, message)
        except smtplib.SMTPServerDisconnected:
            pass
        finally:
            self.server.close()

        self.server = None
        self._login = False
        self._needstls = True
        logger.debug('Logout successful')

    def check_available(self) -> bool:
        """test server"""
        try:
            self.login()
            self.logout()
            return True
        except Exception as e:
            logger.error('server does not available'+e)
            return False

    def is_login(self) -> bool:
        return self._login

    def __repr__(self):
        return '<{} username:{} password:{} {}:{}>'.format(self.__class__.__name__, self.username, self.password, self.host, self.port)

    def _send_mails(self, reciver, msg):
        if not self.is_login():
            self.login()
        msg['Bcc'] = COMMASPACE.join(reciver)
        msg = msg.as_string()
        print('发送给%d人 ' % len(reciver), reciver)
        logger.info(('Send to %d people ' % len(reciver))+(' '.join(reciver)))
        ret = None
        try:
            if self.debug:
                print('(DEBUG MODE)Send to ', reciver)
            else:
                ret = self.server.sendmail(
                    self.username, reciver, msg)
        except Exception as e:
            print(e)
            self.fail_send.extend(reciver)
            op = input('=======Send failed. continue?======[Y/n]')
            if 'n' in op:
                self.log_error()
                self.save_last_successful_send()
                sys.exit(0)

        self.all_send.extend(reciver)
        self.total_send += len(reciver)
        if ret:
            self.fail_send.extend(ret.keys())
        print('==========当前尝试发送数====== ', self.total_send)

    def send_all_mails(self, reciver, mail):
        last = len(reciver)
        msg = mail.make_mail()
        if not msg:
            print('邮件错误，请检查邮件设置')
            logger.error('msg error')
            sys.exit(0)
        logger.info('Sender: '+mail.nickname+' '+mail.addr)
        print('发送者: '+mail.nickname+'\n邮箱：'+mail.addr)
        logger.info('Subject: '+mail.subject)
        print('标题: '+mail.subject)
        logger.info('Reciver: {}'.format(last))
        print('接受者人数: {}'.format(last))
        logger.info('Content: \n'+mail.content_txt)
        print('正文内容: \n'+mail.content_txt)
        logger.info('Attachment: '+(' '.join(mail.attachmentname)))
        print('附件: '+(' '.join(mail.attachmentname)))
        print('调试模式：', 'ON' if self.debug else 'OFF')
        input('===================ENTER TO BEGIN===================')
        if not self.is_login():
            self.login()
        send_index = 0
        counter = 0
        login_cnt = 1
        while last > 0:
            if last >= mails_per_send:
                self._send_mails(
                    reciver[send_index:send_index+mails_per_send], copy.deepcopy(msg))
                send_index += mails_per_send
                counter += mails_per_send
                last -= mails_per_send
            elif last > 0:
                self._send_mails(reciver[send_index:], copy.deepcopy(msg))
                send_index += last
                counter += last
                last = 0

            if last > 0 and counter+mails_per_send > mails_per_login:
                print('重新登录')
                logger.info('Relogin')
                self.logout()
                if last > 0 and login_cnt > reset_per_n_login:
                    login_cnt = 0
                    print('登录间隔', reset_interval)
                    logger.info('Interval')
                    if not self.debug:
                        time.sleep(reset_interval)
                else:
                    if not self.debug:
                        time.sleep(interval_between_login)
                self.login()
                counter = 0
                login_cnt += 1
        self.logout()
        print('发送完成')
        logger.info('Finish send')
        self.log_error()
        self.save_last_successful_send()
