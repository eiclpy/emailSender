from email.mime.multipart import MIMEMultipart
from email.utils import COMMASPACE
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import chardet
import email
import os
import base64


class Mail:
    def __init__(self):
        self.attachment = []
        self.attachmentname = []
        self._sender = False
        self._subject = False
        self._content = False

    def set_sender(self, nickname, addr):
        self.nickname = nickname
        self.addr = addr
        self._sender = True

    def set_subject(self, subject):
        self.subject = subject
        self._subject = True

    def set_content(self, content):
        self.content_txt = content
        self._content = True

    def add_attachment(self, attachmentFile: str or list):
        if not isinstance(attachmentFile, list):
            attachmentFile = [attachmentFile]
        for eachfile in attachmentFile:
            basename = os.path.basename(eachfile)
            with open(eachfile, 'rb') as f:
                part = MIMEApplication(f.read())
                filename = base64.b64encode(basename.encode('utf-8'))
                part.add_header('Content-Disposition', 'attachment',
                                filename='=?UTF-8?B?' + filename.decode() + '?=')
            self.attachment.append(part)
            self.attachmentname.append(basename)

    def make_mail(self):
        data = MIMEMultipart()
        if self._sender and self._subject and self._content:
            self.sender = formataddr(
                (Header(self.nickname, 'utf-8').encode(), self.addr))
            data['From'] = self.sender
            data['To'] = COMMASPACE.join([])
            data['Subject'] = Header(self.subject, 'utf-8').encode()
            try:
                self.content = MIMEText(self.content_txt, 'plain', 'utf-8')
            except UnicodeEncodeError:
                self.content = MIMEText(self.content_txt, 'plain', 'gb2312')
            data.attach(self.content)
            for each in self.attachment:
                data.attach(each)
            return data
        return None
