#!/usr/bin/env python3
# -*- coding: gb2312 -*-
import poplib
import email
import datetime
import time
from email.parser import Parser
from email.utils import parseaddr
from email.header import decode_header
import traceback
import sys
import telnetlib
import os
from threading import Timer

# from email.utils import parseaddr

class c_step4_get_email:
    # �ַ�����ת��
    @staticmethod
    def decode_str(str_in):
        value, charset = decode_header(str_in)[0]
        if charset:
            value = value.decode(charset)
        return value
    '''
    #��ַ�ļ���Ӧ
    @staticmethod
    def exchange(addr):
        ad = addr
        #��ȡ�����ļ�
        e=os.path.abspath(os.curdir)
        with open(e+'\\'+'seting.ini', 'r', encoding="gb2312") as f:
            s = [i[:-1].split('<f>') for i in f.readlines()]
        n = len(s)
        #print(s)
        names = locals()
        c = []
        e = []
        for i in range(len(s)):
            names['b' + str(i)] = s[i]
            c.append(names['b' + str(i)][0])
            e.append(names['b' + str(i)][1])
        if ad in c:
            t=c.index(ad)
            ad=e[t]
        else:
            ad = addr
        return ad
    '''
    # �����ʼ�,��ȡ����
    @staticmethod
    def get_att(msg_in, str_day_in,path):
        # import email
        attachment_files=[]
        for part in msg_in.walk():
            #��ȡ�����˵ĵ�ַ
            hdr, addr = parseaddr(msg_in['From'])
            #name �������������ƣ� addr �����������ַ
            #name, charset = decode_header(hdr)[0]
            #if charset:
                #name = name.decode(charset)
            addr, charset1=decode_header(addr)[0]
            if charset1:
                addr=addr.decode(charset1)
            path1 = path+'\\'+'20200108'+'\\'+addr
            if not os.path.exists(path1):
                os.makedirs(path1)
            # ��ȡ������������
            file_name = part.get_filename()
            # contType = part.get_content_type()
            if file_name:
                h = email.header.Header(file_name)
                # �Ը������ƽ��н���
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # ���������ƿɶ���
                    filename = c_step4_get_email.decode_str(str(filename, dh[0][1]))
                    #filename = filename.encode("gb2312")
                    #print(filename)
                # ���ظ���
                #
                for root, dirs, files in os.walk(path1):
                    if filename not in files:
                        data = part.get_payload(decode=True)
                        # ��ָ��Ŀ¼�´����ļ���ע��������ļ���Ҫ��wbģʽ��
                        att_file = open(path1 + '\\' + filename, 'wb')
                        attachment_files.append(filename)
                        att_file.write(data)  # ���渽��
                        att_file.close()
                    else:
                        print()
                         # print(filename+'�Ѵ��ڣ�')
        return attachment_files

    @staticmethod
    def run_ing():
        # �����ʼ���ַ, �����POP3��������ַ:
        email_user = '@.com'
        # �˴���������Ȩ��,���ڵ�¼�������ʼ��ͻ���
        password = ''
        pop3_server = 'imap.exmail.qq.com'
        # ���ڸ�ֵ
        day = datetime.date.today()
        #��ȡ���������ʼ�
        day = day+datetime.timedelta(days=-4)
        str_day = str(day).replace('-', '')
        print(str_day)
        # ���ӵ�POP3������,��Щ�����������Ҫssl���ܣ�����ʹ��poplib.POP3_SSL
        try:
            telnetlib.Telnet('imap.exmail.qq.com', 993)
            server = poplib.POP3_SSL(pop3_server, 993, timeout=10)
        except:
            time.sleep(5)
            server = poplib.POP3(pop3_server, 110, timeout=10)
        # server = poplib.POP3(pop3_server, 110, timeout=120)
        # ���Դ򿪻�رյ�����Ϣ
        # server.set_debuglevel(1)
        # ��ӡPOP3�������Ļ�ӭ����:
        print(server.getwelcome().decode('gb2312'))
        # �����֤:
        server.user(email_user)
        server.pass_(password)
        # �����ʼ�������ռ�ÿռ�:
        print('Messages: %s. Size: %s' % server.stat())
        # list()���������ʼ��ı��:
        resp, mails, octets = server.list()
        # ���Բ鿴���ص��б�����[b'1 82923', b'2 2184', ...]
        print(mails)
        print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
        index = len(mails)
        # ��������ʼ�
        # for i in range(index, 0, -1):
        
        # ˳������ʼ�
        # for i in range(1, index + 1):
        attachment_files = []
        path = "D:\\��ֵ����\\zhqssj_email\\"
        if not os.path.exists(path):
            os.makedirs(path)
        for i in range(index, 0, -1):
            resp, lines, octets = server.retr(i)
            # lines�洢���ʼ���ԭʼ�ı���ÿһ��,
            # �ʼ���ԭʼ�ı�:
            msg_content = b'\r\n'.join(lines).decode('gb2312')
            # �����ʼ�:
            msg = Parser().parsestr(msg_content)

            # ��ȡ�ʼ�ʱ��,��ʽ���ռ�ʱ��
            date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')
            # �ʼ�ʱ���ʽת��
            date2 = time.strftime("%Y%m%d", date1)
            if date2 < str_day:
                # ������break
                # break
                # ˳����continue
                break
            else:
                # ��ȡ����
                c_step4_get_email.get_att(msg, str_day,path)
        # print_info(msg)
        server.quit()
        #����ظ�ִ��
        #t = Timer(2,c_step4_get_email.run_ing)
        #t.start()
if __name__ == '__main__':
    origin = sys.stdout
    day1 = datetime.date.today()
        #��ȡ���������ʼ�
    #day = day+datetime.timedelta(days=-2)
    str_day1 = str(day1).replace('-', '')
    ll=os.path.abspath(os.curdir)
    log_path=ll+'\\log\\'
    if not os.path.exists(log_path):
        os.makedirs(log_path)
    file =log_path+str_day1+'.txt'
    f = open(file, 'w')
    sys.stdout = f
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    try:
        c_step4_get_email.run_ing()
    except Exception as e:
        s = traceback.format_exc()
        print(e)
        tra = traceback.print_exc()
    print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    sys.stdout = origin
    f.close()
