#!/usr/local/bin/python3
from pdf2image import convert_from_path

from email.header import decode_header
import imaplib
import email

import sys
import os
import re
import uuid
import requests
import logging
from logging.handlers import RotatingFileHandler
import urllib.request
import configparser


# Setup the log handlers to stdout and file.
log = logging.getLogger('parser_email')
log.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s | %(name)s | %(levelname)s | %(message)s'
    )
handler_stdout = logging.StreamHandler(sys.stdout)
handler_stdout.setLevel(logging.DEBUG)
handler_stdout.setFormatter(formatter)
log.addHandler(handler_stdout)
handler_file = RotatingFileHandler(
    'parse_email.log',
    mode='a',
    maxBytes=1048576,
    backupCount=9,
    encoding='UTF-8',
    delay=True
)
handler_file.setLevel(logging.DEBUG)
handler_file.setFormatter(formatter)
log.addHandler(handler_file)


class Mail(object):

    def __init__(self, user, password, host, folder):
        self.conn = None
        self.correct_receiver = False
        self.to = None
        self.send_from = None
        self.unseen = []
        self.all = []
        try:
            self.conn = imaplib.IMAP4_SSL(host)
            self.conn.login(user, password)
        except imaplib.IMAP4.error as e:
            log.error("登录失败: %s" % e)
            sys.exit(1)
        log.info("登录成功")
        self.conn.select(folder, readonly=False)

    def unseen_mail(self):
        """ 未读邮件 """
        result, data = self.conn.search(None, 'UNSEEN')
        if result == 'OK':
            self.unseen = data[0].split()
            log.info('未读邮件数量:%s' % len(self.unseen))
            # print(' '.join([str(i) for i in self.unseen]))

    def all_mail(self): #暂时不使用
        """ 所有邮件 """
        result, data = self.conn.search(None, 'ALL')
        if result == 'OK':
            self.all = data[0].split()
            log.info('所有邮件数量:%s' % len(self.all))

    def parse_header(self, msg):
        data = None
        charset = None
        try:
            data, charset = email.header.decode_header(msg['subject'])[0]
        except:
            pass
        charset = charset or 'utf8'
        self.charset = charset
        log.debug("header编码是" + charset)
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
        #self.correct_receiver = True
        self.send_from = email.utils.parseaddr(msg['From'])[1]

        log.info("Subject: ", subject)
        log.info("From: ", email.utils.parseaddr(msg['From'])[1])
        log.info("To: ", email.utils.parseaddr(msg['To'])[1])
        log.info("Date: ", msg['Date'])

    def parse_part_to_str(self, part):
        charset = part.get_charset() or self.charset
        log.info("body编码是" + charset)
        payload = part.get_payload(decode=True)
        if not payload:
            return
        return str(part.get_payload(decode=True), charset)

    def parse_body(self, msg):
        if not self.correct_receiver:
            log.info('忽略无关邮件')
            return

        for part in msg.walk():
            if not part.is_multipart():
                charset = part.get_charset()
                contenttype = part.get_content_type()
                name = part.get_param("name")
                filePath = None
                if name:
                    fh = email.header.Header(name)
                    fdh = email.header.decode_header(fh)
                    fname = fdh[0][0]

                    fileName = part.get_filename()

                    try:
                        fileName = decode_header(fileName)[0][0].decode(decode_header(fileName)[0][1])
                    except:
                        pass

                    log.info('正在处理附件:', fileName)
                    if not fileName.endswith('.pdf'):
                        log.info('文件名称' + fileName + '不是pdf, 跳过不下载')
                        continue

                    if bool(fileName):
                        filePath = os.path.join('.', 'attachments', fileName)
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                else: # 需要在正文找链接下载的情况（比如京东）
                    log.info('处理正文')
                    if self.send_from == 'customer_service@jd.com':
                        log.info('分析京东的邮件，找到发票下载地址')
                        mail_contents = self.parse_part_to_str(part) # print 邮件正文
                        log.debug(mail_contents)
                        m = re.search(r'.*<a href="(.*)">电子普通发票下载</a>.*', mail_contents)
                        download_url = m.group(1)
                        log.info('京东的附件下载地址:' + download_url)
                        filePath = './attachments/' + str(uuid.uuid1()) + 'jd_attachment.pdf'
                        response = urllib.request.urlopen(download_url)
                        with open(filePath,'wb') as output:
                          output.write(response.read())

                if bool(filePath):
                    log.info('pdf file path is:' + filePath)
                    image = convert_from_path(filePath)
                    image[0].save('temp.png')
                    sfiles={'file': open('temp.png','rb')}
                    res=requests.post('http://180.76.188.189:8890/api/v1/invoices/invoice/email/qrcode',files=sfiles, data={'mobile': self.to} )
                    log.debug (res.text)

    def parse(self):
        nums = self.unseen
        for num in nums:
            try:
                result, data = self.conn.fetch(num, '(RFC822)')
                if result == 'OK':
                    msg = email.message_from_string(data[0][1].decode())
                    log.info('Message %s' % num.decode())
                    self.parse_header(msg)
                    print('-'* 20)
                    self.parse_body(msg)
                    typ, data = self.conn.store(num,'+FLAGS','\\Seen')
            except Exception as e:
                log.error('Message %s 解析错误:%s' % (num, e))
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

    config = configparser.ConfigParser()
    config.read('parse_email.ini')
    config_user = config.get('imap', 'username')
    config_password = config.get('imap', 'password')
    config_host = config.get('imap', 'host')
    config_folder = config.get('imap', 'folder')

    log.info('邮箱基本信息, user: {0}, passowrd:{1}, host: {2}, floder:{3} '.format(
        config_user, config_password, config_host, config_folder
    ))

    mail = Mail(config_user, config_password, config_host, config_folder)
    mail.run()
