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
    # 字符编码转换
    @staticmethod
    def decode_str(str_in):
        value, charset = decode_header(str_in)[0]
        if charset:
            value = value.decode(charset)
        return value
    '''
    #地址文件对应
    @staticmethod
    def exchange(addr):
        ad = addr
        #获取配置文件
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
    # 解析邮件,获取附件
    @staticmethod
    def get_att(msg_in, str_day_in,path):
        # import email
        attachment_files=[]
        for part in msg_in.walk():
            #获取发件人的地址
            hdr, addr = parseaddr(msg_in['From'])
            #name 发送人邮箱名称， addr 发送人邮箱地址
            #name, charset = decode_header(hdr)[0]
            #if charset:
                #name = name.decode(charset)
            addr, charset1=decode_header(addr)[0]
            if charset1:
                addr=addr.decode(charset1)
            path1 = path+'\\'+'20200108'+'\\'+addr
            if not os.path.exists(path1):
                os.makedirs(path1)
            # 获取附件名称类型
            file_name = part.get_filename()
            # contType = part.get_content_type()
            if file_name:
                h = email.header.Header(file_name)
                # 对附件名称进行解码
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # 将附件名称可读化
                    filename = c_step4_get_email.decode_str(str(filename, dh[0][1]))
                    #filename = filename.encode("gb2312")
                    #print(filename)
                # 下载附件
                #
                for root, dirs, files in os.walk(path1):
                    if filename not in files:
                        data = part.get_payload(decode=True)
                        # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                        att_file = open(path1 + '\\' + filename, 'wb')
                        attachment_files.append(filename)
                        att_file.write(data)  # 保存附件
                        att_file.close()
                    else:
                        print()
                         # print(filename+'已存在！')
        return attachment_files

    @staticmethod
    def run_ing():
        # 输入邮件地址, 口令和POP3服务器地址:
        email_user = '@.com'
        # 此处密码是授权码,用于登录第三方邮件客户端
        password = ''
        pop3_server = 'imap.exmail.qq.com'
        # 日期赋值
        day = datetime.date.today()
        #收取最近三天的邮件
        day = day+datetime.timedelta(days=-4)
        str_day = str(day).replace('-', '')
        print(str_day)
        # 连接到POP3服务器,有些邮箱服务器需要ssl加密，可以使用poplib.POP3_SSL
        try:
            telnetlib.Telnet('imap.exmail.qq.com', 993)
            server = poplib.POP3_SSL(pop3_server, 993, timeout=10)
        except:
            time.sleep(5)
            server = poplib.POP3(pop3_server, 110, timeout=10)
        # server = poplib.POP3(pop3_server, 110, timeout=120)
        # 可以打开或关闭调试信息
        # server.set_debuglevel(1)
        # 打印POP3服务器的欢迎文字:
        print(server.getwelcome().decode('gb2312'))
        # 身份认证:
        server.user(email_user)
        server.pass_(password)
        # 返回邮件数量和占用空间:
        print('Messages: %s. Size: %s' % server.stat())
        # list()返回所有邮件的编号:
        resp, mails, octets = server.list()
        # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
        print(mails)
        print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
        index = len(mails)
        # 倒序遍历邮件
        # for i in range(index, 0, -1):
        
        # 顺序遍历邮件
        # for i in range(1, index + 1):
        attachment_files = []
        path = "D:\\估值数据\\zhqssj_email\\"
        if not os.path.exists(path):
            os.makedirs(path)
        for i in range(index, 0, -1):
            resp, lines, octets = server.retr(i)
            # lines存储了邮件的原始文本的每一行,
            # 邮件的原始文本:
            msg_content = b'\r\n'.join(lines).decode('gb2312')
            # 解析邮件:
            msg = Parser().parsestr(msg_content)

            # 获取邮件时间,格式化收件时间
            date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')
            # 邮件时间格式转换
            date2 = time.strftime("%Y%m%d", date1)
            if date2 < str_day:
                # 倒叙用break
                # break
                # 顺叙用continue
                break
            else:
                # 获取附件
                c_step4_get_email.get_att(msg, str_day,path)
        # print_info(msg)
        server.quit()
        #添加重复执行
        #t = Timer(2,c_step4_get_email.run_ing)
        #t.start()
if __name__ == '__main__':
    origin = sys.stdout
    day1 = datetime.date.today()
        #收取最近三天的邮件
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
