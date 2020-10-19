#!/usr/bin/env python3 
# -*- coding: utf-8 -*-

import requests
import time
import re
import publicFun
from lxml import etree

def get_email(email):
    ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3542.0 Safari/537.36'
    headers = {
        'User-Agent': ua,
        'Connection': 'close'
    }
    proxy = {
        "http": "192.168.1.250:7890"
    }
    url = 'https://email-fake.com/' + email
    response = requests.get(url,headers=headers,proxies=proxy,timeout = 10)
    content_url = ''
    tree = etree.HTML(response.text)
    a_list = tree.xpath('//*[@id="email-table"]/a')
    for a in a_list:
        From = a.xpath('./div[1]/text()')
        if From[0] == 'NO_Reply@mccd.edu':
            href = a.xpath('./@href')[0]
            content_url = 'https://email-fake.com' + href
            break

    resp = requests.get(content_url,headers=headers,proxies=proxy,timeout = 10)
    content_tree = etree.HTML(resp.text)
    mail_content = content_tree.xpath('//*[@id="iddelet2"]/div[4]/div[3]/p/text()')
    stu_id = re.search('\d{7}', mail_content[0]).group()
    #print(stu_id)
    edu_email = content_tree.xpath('//*[@id="iddelet2"]/div[4]/div[3]/p/a[4]/text()')[0]
    #print(edu_email)
    return stu_id,edu_email

def change_date(date):
    m, d, y = date.split('/')
    m = m.zfill(2)
    d = d.zfill(2)
    y = y[-2:]
    edu_pwd = m + d + y
    return edu_pwd


def main():
    stu_register = publicFun.get_user_detail(2)
    if stu_register:
        email = stu_register['email']
        Birthday = stu_register['birthday']
        print(email, Birthday)
        global stu_id, edu_email
        stu_id, edu_email = '', ''
        edu_pwd = change_date(Birthday)
        try:
            stu_id, edu_email = get_email(email)
        except Exception as e:
            publicFun.update_user_tag(email, 5)
            publicFun.logger.info('%s获取邮箱信息失败'%email)
            print('获取邮箱信息失败')
        print(stu_id, edu_email, edu_pwd)
        if stu_id and edu_email and edu_pwd:
            collect_time = time.strftime("%Y-%m-%d %X", time.localtime())
            publicFun.add_email_detail([stu_id, edu_email, edu_pwd,collect_time])
            publicFun.delete_user_detail_succeed(email)
        else:
            publicFun.logger.info('%s未获得完整学生邮箱信息'%email)
            print('未获得完整学生邮箱信息')
    else:
        time.sleep(300)
    return stu_register

if __name__ == '__main__':
    while True:
        try:
            main()
        except Exception as e:
            print(e)
            time.sleep(300)