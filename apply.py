import asyncio
import pyppeteer
import openpyxl as xl
import os
import requests
import pyautogui
import time
import iat_ws_python3 as iat3
#from apscheduler.schedulers.blocking import BlockingScheduler
import socket
from lxml import etree
from pyppeteer import launch,launcher
import random
#launcher.DEFAULT_ARGS.remove("--enable-automation")
pyppeteer.DEBUG = True
import fake_useragent

def screen_size():
    # 使用tkinter获取屏幕大小
    import tkinter
    tk = tkinter.Tk()
    width = tk.winfo_screenwidth()
    height = tk.winfo_screenheight()
    tk.quit()
    return width, height


url1 = 'https://www.fakeaddressgenerator.com/'
url2 = 'https://www.opencccapply.net/gateway/apply?cccMisCode=531'

async def main(user_detail):
    api_url = 'http://dps.kdlapi.com/api/getdps/?orderid=979670906001794&num=1&pt=1&sep=1'
    proxy_ip = requests.get(api_url).text
    print(proxy_ip)
    psd = 'qaz2020'
    PIN = '9210'

    #launcher.DEFAULT_ARGS.remove('--enable-automation')
    browser = await launch({
        'userAgent' : fake_useragent.UserAgent,
        #'userAgent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3542.0 Safari/537.36',
        'executablePath' : pyppeteer.launcher.executablePath(),
        #'headless': False,
        'dumpio': True,
        'autoClose': True,
        #'ignoreDefaultArgs' : ['--enable-automation'],
        'args': [
            '--no-sandbox',
            "--start-maximized",
            '--disable-infobars',
            '--disable-extensions',
            '--hide-scrollbars',
            '--disable-bundled-ppapi-flash',
            '--mute-audio',
            #'--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-gpu',
            '--proxy-server={}'.format(proxy_ip)
            #'--proxy-server={}'.format('sp4e19a4a3:ny123456@us.smartproxy.com:10000')
            #'--proxy-server=192.168.1.250:7890',
            #'--proxy-server=192.168.1.239:24000'
        ]
        # , userDir=


    })
    #创建一个新的隐身浏览器上下文。这不会与其他浏览器上下文共享 cookie /缓存
    await browser.createIncognitoBrowserContext()

    await asyncio.sleep(1)
    page = await browser.newPage()

    width, height = screen_size()
    #print(width,height)
    await page.setViewport({  # 最大化窗口
        "width": width,
        "height": 900
    })

    await page.goto(url2)

    # await page.evaluate('''() =>{ Object.defineProperties(navigator,{ webdriver:{ get: () => false } }) }''')  # 以下为插入中间js，将淘宝会为了检测浏览器而调用的js修改其结果。
    # await page.evaluate("() =>{ window.navigator.chrome = { runtime: {},  }; }")
    # await page.evaluate("() =>{ Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en'] }); }")
    # await page.evaluate("() =>{ Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5,6], }); }")

    await page.waitFor(5000)
    try:
        await asyncio.wait([

            page.click('#portletContent_u16l1n18 > div > div.ccc-page-section > div > a:nth-child(2)'),
            page.waitForNavigation({'timeout': 1000 * 100}),
        ])
    except Exception:
        print('页面加载超时')
        return 0


    await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
                                     '{ webdriver:{ get: () => false } }) }')

    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000*60}),
        page.click('#accountFormSubmit'),
    ])

    # await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
    #                                  '{ webdriver:{ get: () => false } }) }')

    # Legal Name
    await page.type('#inputFirstName',user_detail['Full Name'].split(' ')[0])
    await page.waitFor(1000)
    await page.type('#inputMiddleName', user_detail['Full Name'].split(' ')[1])
    await page.waitFor(1000)
    await page.type('#inputLastName', user_detail['Full Name'].split(' ')[2])
    await page.waitFor(1000)
    await page.click('#hasOtherNameNo')
    await page.waitFor(1000)
    await page.click('#hasPreferredNameNo')
    # Date of Birth
    await page.select('#inputBirthDateMonth',user_detail['Birthday'].split('/')[0])
    await page.waitFor(1000)
    await page.select('#inputBirthDateDay',user_detail['Birthday'].split('/')[1])
    await page.waitFor(1000)
    await page.type('#inputBirthDateYear',user_detail['Birthday'].split('/')[2])
    await page.waitFor(1000)
    await page.select('#inputBirthDateMonthConfirm',user_detail['Birthday'].split('/')[0])
    await page.waitFor(1000)
    await page.select('#inputBirthDateDayConfirm',user_detail['Birthday'].split('/')[1])
    await page.waitFor(1000)
    await page.type('#inputBirthDateYearConfirm',user_detail['Birthday'].split('/')[2])
    await page.waitFor(1000)
    #Social Security Number
    await page.click('#-have-ssn-yes')
    await page.waitFor(2000)
    await page.type('#-ssn-input1',user_detail['Social Security Number'].replace('-', ''))
    await page.waitFor(1000)
    await page.type('#-ssn-input2',user_detail['Social Security Number'].replace('-', ''))
    await page.waitFor(1000)

    await asyncio.wait([
        page.waitForNavigation(),
        page.click('#accountFormSubmit'),
    ])
    # await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
    #                                  '{ webdriver:{ get: () => false } }) }')

    # Email
    await page.type('#inputEmail',user_detail['email'])
    await page.waitFor(1000)
    await page.type('#inputEmailConfirm',user_detail['email'])
    await page.waitFor(1000)
    # Telephone
    await page.type('#inputSmsPhone',user_detail['Phone Number'])
    await page.waitFor(1000)
    # Permanent Address
    await page.type('#inputStreetAddress1',user_detail['Street'])
    await page.waitFor(1000)
    await page.type('#inputStreetAddress2',user_detail['Street'])
    await page.waitFor(1000)
    await page.type('#inputCity',user_detail['City'])
    await page.waitFor(1000)
    await page.select('#inputState',user_detail['State'])
    await page.waitFor(1000)
    await page.type('#inputPostalCode',user_detail['Zip Code'])
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
    # await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
    #                                  '{ webdriver:{ get: () => false } }) }')

    # Username and Password
    await page.type('#inputUserId',user_detail['username'])
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

    # await page.hover('#recaptcha')
    # print(page.mouse._x,page.mouse._y)
    #
    # await page.hover('#recaptcha > div > div > iframe')
    # print(page.mouse._x, page.mouse._y)

    await page.click('#recaptcha > div > div > iframe')
    print('点击人机验证')
    await page.waitFor(4000)
    '''
    await page.mouse.click(478,754)
    print('点击人机验证')
    await page.waitFor(4000)

    
    await page.mouse.move(550,880)
    await page.mouse.click(576,877)

    print('点击音频验证')
    await page.waitFor(2000)
    '''



    try:
        frame =  page.frames
        for f in frame:
            title = await f.title()
            if title == 'reCAPTCHA':
                await f.click('#recaptcha-audio-button')
                print('点击音频验证')
                await page.waitFor(2000)
                # await f.click('#recaptcha-reload-button')
                # await page.waitFor(2000)
                #f.Jeval('body > div > div > div:nth-child(1) > div.rc-doscaptcha-body > div')
                link = await f.Jeval('#audio-source','el => el.src')
                print(link)
                # 跳转到链接网站进行下载
                # await page._client.send('Page.setDownloadBehavior',{
                #     'behavior': 'allow',
                #     'downloadPath':'./'
                # })
                #download_dir = r'C:\Users\Win\Downloads'
                now_path = os.path.abspath(os.path.dirname(__file__))

                path_text = now_path + '\payload.mp3'

                if os.path.exists(path_text):
                    os.remove(path_text)
                    if os.path.exists(path_text):
                        print('删除失败')
                    else:
                        print('已删除')
                await page.waitFor(5000)
                page1 = await browser.newPage()
                cdp = await page1.target.createCDPSession()
                await cdp.send('Page.setDownloadBehavior',{
                'behavior': 'allow', # 允许所有下载请求
                'downloadPath': now_path # 设置下载路径
                })
                width, height = screen_size()
                await page.setViewport({  # 最大化窗口
                    "width": width,
                    "height": height
                })
                await page1.goto(link)
                await page1.waitFor(1000)


                await page1.hover('body > video ')
                start_x = page1.mouse._x
                start_y = page1.mouse._y
                print(start_x,start_y)
                await page1.mouse.click(start_x+125,start_y+55,{'delay':50})
                print('dianji420')
                await page1.waitFor(1000)
                await page1.mouse.click(start_x+125,start_y+55,{'delay':50})
                print('dianji430')
                await page1.keyboard.press('Enter')


                await page1.waitFor(3000)
                await page1.close()

                data = ''
                await page.waitFor(2000)
                data = iat3.run(path_text)
                print(data)
                await page.waitFor(2000)
                await f.type('#audio-response',data)
                await page.waitFor(1000)
                await f.click('#recaptcha-verify-button')
                await page.waitFor(2000)

    except :
        print("人机验证失败，请重试")
    finally:
        pass
        # try:
        #     os.remove(r'C:\Users\Win\Downloads\payload.mp3')
        # except WindowsError:
        #     pass



    # 现在前期页面注册已完成
    await page.waitFor(2000)
    #await page.click('#accountFormSubmit')
    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000 * 60}),
        page.click('#accountFormSubmit'),

    ])
    print('人机验证通过')
    await page.waitForSelector('#registrationSuccess')
    CCCID = await page.querySelector(
        '#registrationSuccess > main > div.column > div > div > div > p:nth-child(1) > strong')
    tag = False
    print('现在是可以看见CCCID的页面')
    print(CCCID)
    if CCCID == None:
        print('申请失败')
        await browser.close()

    else:

        tag = True


    await asyncio.wait([
        page.waitForNavigation({'timeout': 1000 * 30}),
        page.click('#registrationSuccess > main > div.column > div > div > button')
    ])

    await browser.close()
    return tag




def write_excel_file(folder_path,user_detail):
    result_path = os.path.join(folder_path, "stu_apply.xlsx")
    headers = ['username', 'Birthday', 'email', 'Gender']
    print(result_path)
    apply_detail = []
    apply_detail.append(user_detail['username'])
    apply_detail.append(user_detail['Birthday'])
    apply_detail.append(user_detail['email'])
    apply_detail.append(user_detail['Gender'])
    print('***** 开始写入excel文件 ' + result_path + ' ***** \n')
    if os.path.exists(result_path):
        print('***** excel已存在，在表后添加数据 ' + result_path + ' ***** \n')
        workbook = xl.load_workbook(result_path)
        sheet = workbook.active
        sheet.append(apply_detail)
        workbook.save(result_path)
    else:
        print('***** excel不存在，创建excel ' + result_path + ' ***** \n')
        workbook = xl.Workbook()
        workbook.save(result_path)

        sheet = workbook.active
        sheet.append(headers)

        sheet.append(apply_detail)
        workbook.save(result_path)
    print('***** 生成Excel文件 ' + result_path + ' ***** \n')


def red_excel_file(folder_path):
    result_path = os.path.join(folder_path, "user_detail.xlsx")
    user_detail = {}
    wb2 = xl.load_workbook(result_path)
    sheet = wb2.active
    cell = sheet[1]
    cell1 = sheet[2]
    for i in range(len(cell)):
        user_detail.update({cell[i].value:cell1[i].value})

    sheet.delete_rows(2)
    wb2.save(result_path)
    return user_detail


def updata_excel(folder_path,flag,user_detail):
    now = time.strftime("%d_%m_%Y")
    result_path = os.path.join(folder_path, "qt_%s.xlsx" % now)

    wb2 = xl.load_workbook(result_path)
    sheet = wb2.active

    for row in sheet:
        if row[0].value == user_detail['Full Name']:
            if flag == True:
                row[2].value = '注册中'
            else:
                row[2].value = '失败'
    wb2.save(result_path)

if __name__ == '__main__':
    while True:
        #flag = False

        try :
            user_detail = {}


            # 快代理
            #api_url = "http://dps.kdlapi.com/api/getdps/?orderid=959731018601581&num=1&pt=1&dedup=1&sep=1"
            # 芝麻免费代理
            api_url1 = 'http://http.tiqu.alicdns.com/getip3?num=1&type=1&pro=&city=0&yys=0&port=1&pack=113939&ts=0&ys=0&cs=0&lb=1&sb=0&pb=45&mr=1&regions=&gm=4'
            # IPIDEA 国际代理
            api_url_global = 'http://tiqu.linksocket.com:81/abroad?num=1&type=1&lb=1&sb=0&flow=1&regions=us&n=0'
            # 讯代理
            api_url_xun = 'http://api.xdaili.cn/xdaili-api//privateProxy/getDynamicIP/DD20208187336LbMXZb/8104491c519511e79d9b7cd30abda612?returnType=1'

            username = ''.join(random.sample("1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", 8))

            user_detail = red_excel_file('./')
            user_detail.update({'username': username})
            print(user_detail)
            # name = user_detail['Full Name']
            #
            # first_name = name.split(' ')[0]
            #
            # middle_name = name.split(' ')[1]
            #
            # last_name = name.split(' ')[2]
            #
            # Birthday = user_detail['Birthday']
            #
            # month = Birthday.split('/')[0]
            #
            # day = Birthday.split('/')[1]
            #
            # year = Birthday.split('/')[2]
            #
            # SSN = user_detail['Social Security Number'].replace('-', '')
            try:
                tag = asyncio.get_event_loop().run_until_complete(main(user_detail))
                if tag == True:
                    write_excel_file("./",user_detail)
                    flag = True
                else:
                    print('申请失败')
            except:
                print('申请失败')
        except Exception:
            continue
        finally:
            #updata_excel('/',flag)
            launch().close()
            time.sleep(30)
