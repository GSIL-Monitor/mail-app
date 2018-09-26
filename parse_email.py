#!/usr/local/bin/python3
from pdf2image import convert_from_path,convert_from_bytes

from email.header import decode_header
import imaplib
import email
import sys
import os
import re
import requests
from io import BytesIO
import logging


class Mail(object):
    user = 'invoice@rrs.com'
    password = 'Haier@2018'
    imap = 'imap.exmail.qq.com'

    def __init__(self):
        self.conn = None
        self.correct_receiver = False
        self.to = None
        self.unseen = []
        self.all = []
        try:
            self.conn = imaplib.IMAP4_SSL(self.imap)
            self.conn.login(self.user, self.password)
        except imaplib.IMAP4.error as e:
            print("登录失败: %s" % e)
            sys.exit(1)
        print("登录成功")
        self.conn.select("&UXZO1mWHTvZZOQ-/&kK5O9o9sefs-", readonly=False)

    def unseen_mail(self):
        """ 未读邮件 """
        result, data = self.conn.search(None, 'UNSEEN')
        if result == 'OK':
            self.unseen = data[0].split()
            print('未读邮件数量:%s' % len(self.unseen))
            # print(' '.join([str(i) for i in self.unseen]))

    def all_mail(self): #暂时不使用
        """ 所有邮件 """
        result, data = self.conn.search(None, 'ALL')
        if result == 'OK':
            self.all = data[0].split()
            print('所有邮件数量:%s' % len(self.all))

    def parse_header(self, msg):
        data, charset = email.header.decode_header(msg['subject'])[0]
        charset = charset or 'utf8'
        self.charset = charset
        print("header编码是" + charset)
        if type(data) == str:
            subject = data
        else:
            subject = str(data, charset)

        to = email.utils.parseaddr(msg['to'])[1]
        if re.match(r'\d{11,11}@rrs.com', to):
            self.correct_receiver = True
            self.to = to.replace('@rrs.com','')
        else:
            self.correct_receiver = False

        print("Subject: ", subject)
        print("From: ", email.utils.parseaddr(msg['From'])[1])
        print("To: ", email.utils.parseaddr(msg['To'])[1])
        print("Date: ", msg['Date'])

    def parse_part_to_str(self, part):
        charset = part.get_charset() or self.charset
        print("body编码是" + charset)
        payload = part.get_payload(decode=True)
        if not payload:
            return
        return str(part.get_payload(decode=True), charset)

    def parse_body(self, msg):
        if not self.correct_receiver:
            return

        for part in msg.walk():
            if not part.is_multipart():
                charset = part.get_charset()
                contenttype = part.get_content_type()
                name = part.get_param("name")
                if name:
                    fh = email.header.Header(name)
                    fdh = email.header.decode_header(fh)
                    fname = fdh[0][0]
                    print('附件名 before decode:', fname)

                    fileName = part.get_filename()

                    try:
                        fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
                    except:
                        print('do not need to decode filename')

                    print('附件名 afeter decode:', fileName)
                    if not fileName.endswith('.pdf'):
                        return

                    if bool(fileName):
                        filePath = os.path.join('.', 'attachments', fileName)
                        print(filePath)
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                        image = convert_from_path(filePath)
                        image[0].save('temp.png')
                        sfiles={'file': open('temp.png','rb')}
                        res=requests.post('http://180.76.188.189:8890/api/v1/invoices/invoice/email/qrcode',files=sfiles, data={'mobile': self.to} )
                        print (res.text)
                else:
                    print('no attachments')
                    #print(self.parse_part_to_str(part)) # print 邮件正文


    def parse(self):
        #nums = self.all[1:3]
        nums = self.unseen
        for num in nums:
            try:
                result, data = self.conn.fetch(num, '(RFC822)')
                if result == 'OK':
                    msg = email.message_from_string(data[0][1].decode())
                    print('Message %s' % num.decode())
                    self.parse_header(msg)
                    print('-'* 20)
                    self.parse_body(msg)
                    typ, data = self.conn.store(num,'+FLAGS','\\Seen')
            except Exception as e:
                print('Message %s 解析错误:%s' % (num, e))
                logging.exception("解析错误")


    def over(self):
        self.conn.close()
        self.conn.logout()

    def run(self):
        while True:
            self.unseen_mail()
            #self.all_mail()
            self.parse()
        self.over()


if __name__ == '__main__':
    mail = Mail()
    mail.run()
