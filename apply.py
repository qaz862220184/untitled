# -*- encoding=utf8 -*-
import asyncio
import pyppeteer
import openpyxl as xl
import os
import requests
import time
import iat_ws_python3 as iat3
from pyppeteer import launch
from pyppeteer.dialog import Dialog
import random
import shutil
import socket
import fake_useragent
import publicFun
import traceback

pyppeteer.DEBUG = True

def GetSessionProxy():
    super_proxy = socket.gethostbyname('zproxy.lum-superproxy.io')
    url = "http://%s-country-us-session-%s:%s@" + super_proxy + ":%d"
    port = 22225
    session_id = random.randint(0, 116225344)
    return url % ('lum-customer-hl_f5a6deb2-zone-sellerbdata', session_id, 'bcvjau85n9e9', port)

class Browser(object):
    browser = None

    def __init__(self,user_detail):
        self.user_detail = user_detail

    async def newbrowser(self):
        try:
            api_url = "https://dev.kdlapi.com/api/setipwhitelist"
            data = {
                'orderid': '919944088273725',
                'signature': 'qsehuojro4gehrwm35vb6wx78vl8oyav',
            }
            requests.post(url=api_url, data=data)
            api_url = 'http://dps.kdlapi.com/api/getdps/?orderid=919944088273725&num=1&pt=1&sep=1'
            proxy_ip = requests.get(api_url).text
            self.browser = await launch({
                'userAgent': fake_useragent.UserAgent,
                'executablePath': pyppeteer.launcher.executablePath(),
                'headless': False,
                'dumpio': True,
                'autoClose': False,
                'handleSIGTERM': True,
                'handleSIGHUP': True,
                # 'ignoreDefaultArgs' : ['--enable-automation'],
                'args': [
                    '--disable-zero-browsers-open-for-tests',
                    '--crash-loop-before',
                    '--no-sandbox',
                    "--start-maximized",
                    '--disable-infobars',
                    '--disable-extensions',
                    '--hide-scrollbars',
                    '--disable-bundled-ppapi-flash',
                    '--mute-audio',
                    '--disable-setuid-sandbox',
                    '--disable-gpu',
                    '--proxy-server={}'.format(proxy_ip)
                    # '--proxy-server=192.168.1.250:7890'
                ]
            }, userDataDir='./apply/%s' % self.user_detail['email'])
        except:
            print("浏览器打开失败")

    async def handle_dialog(dialog: Dialog):
        await dialog.dismiss()

    async def apply(self,browser):
        psd = 'qaz2020'
        PIN = '9210'
        await browser.createIncognitoBrowserContext()
        await asyncio.sleep(1)
        page = await browser.newPage()
        #page.on("dialog", self.handle_dialog)
        width, height = publicFun.screen_size()
        await page.setViewport({  # 最大化窗口
            "width": width,
            "height": 900
        })
        await page.goto('https://www.opencccapply.net/gateway/apply?cccMisCode=531')
        await page.waitFor(5000)
        try:
            await asyncio.wait([
                page.click('#portletContent_u16l1n18 > div > div.ccc-page-section > div > a:nth-child(2)'),
                page.waitForNavigation({'timeout': 1000 * 60}),
            ])
        except Exception:
            await page.close()
            await browser.close()
            print('页面加载超时')
        await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
                                         '{ webdriver:{ get: () => false } }) }')
        await asyncio.wait([
            page.waitForNavigation({'timeout': 1000 * 60}),
            page.click('#accountFormSubmit'),
        ])
        # Legal Name
        await page.type('#inputFirstName', self.user_detail['fullName'].split(' ')[0])
        await page.waitFor(1000)
        await page.type('#inputMiddleName', self.user_detail['fullName'].split(' ')[1])
        await page.waitFor(1000)
        await page.type('#inputLastName', self.user_detail['fullName'].split(' ')[2])
        await page.waitFor(1000)
        await page.click('#hasOtherNameNo')
        await page.waitFor(1000)
        await page.click('#hasPreferredNameNo')
        # Date of Birth
        await page.select('#inputBirthDateMonth', self.user_detail['birthday'].split('/')[0])
        await page.waitFor(1000)
        await page.select('#inputBirthDateDay', self.user_detail['birthday'].split('/')[1])
        await page.waitFor(1000)
        await page.type('#inputBirthDateYear', self.user_detail['birthday'].split('/')[2])
        await page.waitFor(1000)
        await page.select('#inputBirthDateMonthConfirm', self.user_detail['birthday'].split('/')[0])
        await page.waitFor(1000)
        await page.select('#inputBirthDateDayConfirm', self.user_detail['birthday'].split('/')[1])
        await page.waitFor(1000)
        await page.type('#inputBirthDateYearConfirm', self.user_detail['birthday'].split('/')[2])
        await page.waitFor(1000)
        # Social Security Number
        await page.click('#-have-ssn-yes')
        await page.waitFor(2000)
        await page.type('#-ssn-input1', self.user_detail['ssn'].replace('-', ''))
        await page.waitFor(1000)
        await page.type('#-ssn-input2', self.user_detail['ssn'].replace('-', ''))
        await page.waitFor(1000)
        await asyncio.wait([
            page.waitForNavigation(),
            page.click('#accountFormSubmit'),
        ])
        # Email
        await page.type('#inputEmail', self.user_detail['email'])
        await page.waitFor(1000)
        await page.type('#inputEmailConfirm', self.user_detail['email'])
        await page.waitFor(1000)
        # Telephone
        await page.type('#inputSmsPhone', self.user_detail['phoneNumber'])
        await page.waitFor(1000)
        # Permanent Address
        await page.type('#inputStreetAddress1', self.user_detail['street'])
        await page.waitFor(1000)
        await page.type('#inputStreetAddress2', self.user_detail['street'])
        await page.waitFor(1000)
        await page.type('#inputCity', self.user_detail['city'])
        await page.waitFor(1000)
        await page.select('#inputState', self.user_detail['state'])
        await page.waitFor(1000)
        await page.type('#inputPostalCode', self.user_detail['zipCode'])
        await page.waitFor(1000)
        # 提交报错
        await page.click('#accountFormSubmit'),
        await page.waitFor(5000)
        await page.click('#messageFooterLabel'),
        await page.waitFor(2000)
        await page.click('#inputAddressValidationOverride'),
        await page.waitFor(2000)
        await asyncio.wait([
            page.waitForNavigation({'timeout': 60000}),
            page.click('#accountFormSubmit'),
        ])
        # Username and Password
        await page.type('#inputUserId', self.user_detail['userName'])
        await page.waitFor(1000)
        await page.type('#inputPasswd', psd)
        await page.waitFor(1000)
        await page.type('#inputPasswdConfirm', psd)
        await page.waitFor(1000)
        # Security PIN
        await page.type('#inputPin', PIN)
        await page.waitFor(1000)
        await page.type('#inputPinConfirm', PIN)
        await page.waitFor(1000)
        # Security Questions
        await page.select('#inputSecurityQuestion1', '1')
        await page.waitFor(1000)
        await page.type('#inputSecurityAnswer1', 'first')
        await page.waitFor(1000)
        await page.select('#inputSecurityQuestion2', '2')
        await page.waitFor(1000)
        await page.type('#inputSecurityAnswer2', 'second')
        await page.waitFor(1000)
        await page.select('#inputSecurityQuestion3', '3')
        await page.waitFor(1000)
        await page.type('#inputSecurityAnswer3', 'thirdly')
        await page.waitFor(1000)
        await page.click('#recaptcha > div > div > iframe')
        print('点击人机验证')
        await page.waitFor(2000)
        now_path = os.path.abspath(os.path.abspath(os.path.dirname(__file__)))
        now_path = now_path + '_' + self.user_detail['userName']

        try:
            frame = page.frames
            for f in frame:
                title = await f.title()
                if title == 'reCAPTCHA':
                    await f.click('#recaptcha-audio-button')
                    await page.waitFor(2000)
                    link = await f.Jeval('#audio-source', 'el => el.src')
                    path_text = now_path + '\payload.mp3'
                    await page.waitFor(5000)
                    page1 = await browser.newPage()
                    cdp = await page1.target.createCDPSession()
                    await cdp.send('Page.setDownloadBehavior', {
                        'behavior': 'allow',  # 允许所有下载请求
                        'downloadPath': now_path  # 设置下载路径
                    })
                    width, height = publicFun.screen_size()
                    await page.setViewport({  # 最大化窗口
                        "width": width,
                        "height": height
                    })
                    await page1.goto(link)
                    await page1.waitFor(1000)

                    await page1.hover('body > video ')
                    start_x = page1.mouse._x
                    start_y = page1.mouse._y
                    print(start_x, start_y)
                    await page1.mouse.click(start_x + 125, start_y + 55, {'delay': 50})
                    await page1.waitFor(1000)
                    await page1.mouse.click(start_x + 125, start_y + 55, {'delay': 50})
                    await page1.keyboard.press('Enter')
                    await page1.waitFor(3000)
                    await page1.close()

                    await page.waitFor(2000)
                    data = iat3.run(path_text)
                    await page.waitFor(2000)
                    await f.type('#audio-response', data)
                    await page.waitFor(1000)
                    await f.click('#recaptcha-verify-button')
                    await page.waitFor(2000)

                    break
        except:
            await page.close()
            await browser.close()
            print("人机验证失败，请重试")
            return False
        finally:
            try:
                shutil.rmtree(now_path)
            except WindowsError:
                pass

        await page.waitFor(2000)
        await asyncio.wait([
            page.waitForNavigation({'timeout': 1000 * 60}),
            page.click('#accountFormSubmit'),
        ])
        await page.waitForSelector('#registrationSuccess')
        CCCID = await page.querySelector(
            '#registrationSuccess > main > div.column > div > div > div > p:nth-child(1) > strong')
        tag = False
        if CCCID == None:
            print('申请失败')
            await browser.close()
        else:
            tag = True
        await asyncio.wait([
            page.waitForNavigation({'timeout': 1000 * 30}),
            page.click('#registrationSuccess > main > div.column > div > div > button')
        ])
        await page.close()
        await browser.close()
        return tag

    async def open(self):
        await self.newbrowser()
        await self.apply(self.browser)

    async def close(self):
        await self.browser.close()

    def main(self):
        try:
            loop = asyncio.get_event_loop()
            tag = loop.run_until_complete(self.open())
            if tag == True:
                publicFun.update_user_tag_time(user_detail['email'], 1)
            else:
                publicFun.update_user_tag_time(user_detail['email'], 3)
                print('没有得到cccid')
        except:
            publicFun.update_user_tag_time(user_detail['email'], 3)
            print('申请失败')
        finally:
            print(time.strftime("%Y-%m-%d %X", time.localtime()))
            asyncio.get_event_loop().run_until_complete(self.close())
            time.sleep(10)

async def main(user_detail):
    api_url = "https://dev.kdlapi.com/api/setipwhitelist"
    data = {
        'orderid': '919944088273725',
        'signature': 'qsehuojro4gehrwm35vb6wx78vl8oyav',
    }
    requests.post(url=api_url, data=data)
    api_url = 'http://dps.kdlapi.com/api/getdps/?orderid=919944088273725&num=1&pt=1&sep=1'
    proxy_ip = requests.get(api_url).text
    print(proxy_ip)
    #proxy_ip = GetSessionProxy()
    psd = 'qaz2020'
    PIN = '9210'
    browser = await launch({
        'userAgent' : fake_useragent.UserAgent,
        'executablePath' : pyppeteer.launcher.executablePath(),
        #'headless': False,
        'dumpio': True,
        'autoClose': True,
        'handleSIGTERM':True,
        'handleSIGHUP':True,
        #'ignoreDefaultArgs' : ['--enable-automation'],
        'args': [
            '--disable-zero-browsers-open-for-tests',
            '--crash-loop-before',
            '--no-sandbox',
            "--start-maximized",
            '--disable-infobars',
            '--disable-extensions',
            '--hide-scrollbars',
            '--disable-bundled-ppapi-flash',
            '--mute-audio',
            '--disable-setuid-sandbox',
            '--disable-gpu',
            '--proxy-server={}'.format(proxy_ip)
            #'--proxy-server=192.168.1.250:7890'
        ]
    },userDataDir='./apply/%s'%user_detail['email'] )
    #pid = browser.process.pid
    #创建一个新的隐身浏览器上下文。这不会与其他浏览器上下文共享 cookie /缓存
    await browser.createIncognitoBrowserContext()
    await asyncio.sleep(1)
    page = await browser.newPage()
    width, height = publicFun.screen_size()
    await page.setViewport({  # 最大化窗口
        "width": width,
        "height": 900
    })
    await page.goto('https://www.opencccapply.net/gateway/apply?cccMisCode=531')
    await page.waitFor(5000)
    try:
        await asyncio.wait([
            page.click('#portletContent_u16l1n18 > div > div.ccc-page-section > div > a:nth-child(2)'),
            page.waitForNavigation({'timeout': 1000 * 60}),
        ])
    except Exception:
        await page.close()
        await browser.close()
        print('页面加载超时')
    await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
                                     '{ webdriver:{ get: () => false } }) }')
    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000*60}),
        page.click('#accountFormSubmit'),
    ])
    # Legal Name
    await page.type('#inputFirstName',user_detail['fullName'].split(' ')[0])
    await page.waitFor(1000)
    await page.type('#inputMiddleName', user_detail['fullName'].split(' ')[1])
    await page.waitFor(1000)
    await page.type('#inputLastName', user_detail['fullName'].split(' ')[2])
    await page.waitFor(1000)
    await page.click('#hasOtherNameNo')
    await page.waitFor(1000)
    await page.click('#hasPreferredNameNo')
    # Date of Birth
    await page.select('#inputBirthDateMonth',user_detail['birthday'].split('/')[0])
    await page.waitFor(1000)
    await page.select('#inputBirthDateDay',user_detail['birthday'].split('/')[1])
    await page.waitFor(1000)
    await page.type('#inputBirthDateYear',user_detail['birthday'].split('/')[2])
    await page.waitFor(1000)
    await page.select('#inputBirthDateMonthConfirm',user_detail['birthday'].split('/')[0])
    await page.waitFor(1000)
    await page.select('#inputBirthDateDayConfirm',user_detail['birthday'].split('/')[1])
    await page.waitFor(1000)
    await page.type('#inputBirthDateYearConfirm',user_detail['birthday'].split('/')[2])
    await page.waitFor(1000)
    #Social Security Number
    await page.click('#-have-ssn-yes')
    await page.waitFor(2000)
    await page.type('#-ssn-input1',user_detail['ssn'].replace('-', ''))
    await page.waitFor(1000)
    await page.type('#-ssn-input2',user_detail['ssn'].replace('-', ''))
    await page.waitFor(1000)
    await asyncio.wait([
        page.waitForNavigation(),
        page.click('#accountFormSubmit'),
    ])
    # Email
    await page.type('#inputEmail',user_detail['email'])
    await page.waitFor(1000)
    await page.type('#inputEmailConfirm',user_detail['email'])
    await page.waitFor(1000)
    # Telephone
    await page.type('#inputSmsPhone',user_detail['phoneNumber'])
    await page.waitFor(1000)
    # Permanent Address
    await page.type('#inputStreetAddress1',user_detail['street'])
    await page.waitFor(1000)
    await page.type('#inputStreetAddress2',user_detail['street'])
    await page.waitFor(1000)
    await page.type('#inputCity',user_detail['city'])
    await page.waitFor(1000)
    await page.select('#inputState',user_detail['state'])
    await page.waitFor(1000)
    await page.type('#inputPostalCode',user_detail['zipCode'])
    await page.waitFor(1000)
    # 提交报错
    await page.click('#accountFormSubmit'),
    await page.waitFor(5000)
    await page.click('#messageFooterLabel'),
    await page.waitFor(2000)
    await page.click('#inputAddressValidationOverride'),
    await page.waitFor(2000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':60000}),
        page.click('#accountFormSubmit'),
    ])
    # Username and Password
    await page.type('#inputUserId',user_detail['userName'])
    await page.waitFor(1000)
    await page.type('#inputPasswd',psd)
    await page.waitFor(1000)
    await page.type('#inputPasswdConfirm',psd)
    await page.waitFor(1000)
    # Security PIN
    await page.type('#inputPin',PIN)
    await page.waitFor(1000)
    await page.type('#inputPinConfirm',PIN)
    await page.waitFor(1000)
    # Security Questions
    await page.select('#inputSecurityQuestion1','1')
    await page.waitFor(1000)
    await page.type('#inputSecurityAnswer1','first')
    await page.waitFor(1000)
    await page.select('#inputSecurityQuestion2', '2')
    await page.waitFor(1000)
    await page.type('#inputSecurityAnswer2', 'second')
    await page.waitFor(1000)
    await page.select('#inputSecurityQuestion3', '3')
    await page.waitFor(1000)
    await page.type('#inputSecurityAnswer3', 'thirdly')
    await page.waitFor(1000)
    await page.click('#recaptcha > div > div > iframe')
    print('点击人机验证')
    await page.waitFor(2000)
    now_path = os.path.abspath(os.path.abspath(os.path.dirname(__file__)))
    now_path = now_path +'_'+ user_detail['userName']

    try:
        frame =  page.frames
        for f in frame:
            title = await f.title()
            if title == 'reCAPTCHA':
                await f.click('#recaptcha-audio-button')
                await page.waitFor(2000)
                link = await f.Jeval('#audio-source', 'el => el.src')
                path_text = now_path + '\payload.mp3'
                await page.waitFor(5000)
                page1 = await browser.newPage()
                cdp = await page1.target.createCDPSession()
                await cdp.send('Page.setDownloadBehavior', {
                    'behavior': 'allow',  # 允许所有下载请求
                    'downloadPath': now_path  # 设置下载路径
                })
                width, height = publicFun.screen_size()
                await page.setViewport({  # 最大化窗口
                    "width": width,
                    "height": height
                })
                await page1.goto(link)
                await page1.waitFor(1000)

                await page1.hover('body > video ')
                start_x = page1.mouse._x
                start_y = page1.mouse._y
                print(start_x, start_y)
                await page1.mouse.click(start_x + 125, start_y + 55, {'delay': 50})
                await page1.waitFor(1000)
                await page1.mouse.click(start_x + 125, start_y + 55, {'delay': 50})
                await page1.keyboard.press('Enter')
                await page1.waitFor(3000)
                await page1.close()

                await page.waitFor(2000)
                data = iat3.run(path_text)
                await page.waitFor(2000)
                await f.type('#audio-response', data)
                await page.waitFor(1000)
                await f.click('#recaptcha-verify-button')
                await page.waitFor(2000)

                break
    except :
        await page.close()
        await browser.close()
        print("人机验证失败，请重试")
        return False
    finally:
        try:
            shutil.rmtree(now_path)
        except WindowsError:
            pass

    await page.waitFor(2000)
    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000 * 60}),
        page.click('#accountFormSubmit'),
    ])
    await page.waitForSelector('#registrationSuccess')
    CCCID = await page.querySelector(
        '#registrationSuccess > main > div.column > div > div > div > p:nth-child(1) > strong')
    tag = False
    if CCCID == None:
        print('申请失败')
        await browser.close()
    else:
        tag = True
    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000 * 30}),
        page.click('#registrationSuccess > main > div.column > div > div > button')
    ])
    await page.close()
    await browser.close()
    return tag

if __name__ == '__main__':

    while True:
        user_detail = publicFun.get_user_detail(0)
        if user_detail:
            try :
                print(user_detail)
                try:
                    tag = asyncio.get_event_loop().run_until_complete(main(user_detail))
                    if tag == True:
                        publicFun.update_user_tag(user_detail['email'],1)
                        flag = True
                    else:
                        publicFun.update_user_tag(user_detail['email'], 3)
                        print('人机验证失败')
                except:
                    publicFun.update_user_tag(user_detail['email'], 3)
                    #traceback.print_exc()
                    print('申请失败')
            except Exception:
                traceback.print_exc()
                continue
            finally:
                print(time.strftime("%Y-%m-%d %X", time.localtime()))
                time.sleep(10)

