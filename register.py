import asyncio
import pyppeteer
import openpyxl as xl
import os,time
import requests
from pyppeteer import launch,launcher
import random
import fake_useragent
#launcher.DEFAULT_ARGS.remove("--enable-automation")
pyppeteer.DEBUG = True
url2 = 'https://www.opencccapply.net/gateway/apply?cccMisCode=531'
def screen_size():
    # 使用tkinter获取屏幕大小
    import tkinter
    tk = tkinter.Tk()
    width = tk.winfo_screenwidth()
    height = tk.winfo_screenheight()
    tk.quit()
    return width, height


async def regist(apply_detail):
    api_url = 'http://dps.kdlapi.com/api/getdps/?orderid=979670906001794&num=1&pt=1&sep=1'
    proxy_ip = requests.get(api_url).text
    print(proxy_ip)
    browser = await launch({
        'userAgent' : fake_useragent.UserAgent,

        'executablePath': pyppeteer.launcher.executablePath(),
        #'headless': False,
        'dumpio': True,
            'args': [
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
            #'--proxy-server=192.168.1.250:7890',
            #'--proxy-server=192.168.1.239:24010'
        ]
    })
    # 创建一个新的隐身浏览器上下文。这不会与其他浏览器上下文共享 cookie /缓存
    await browser.createIncognitoBrowserContext()

    await asyncio.sleep(1)
    page = await browser.newPage()

    width, height = screen_size()
    # print(width,height)
    await page.setViewport({  # 最大化窗口
        "width": width,
        "height": 900
    })
    # await page.evaluateOnNewDocument('() =>{ Object.defineProperties(navigator,'
    #                                  '{ webdriver:{ get: () => false } }) }')

    await page.goto(url2)

    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#portal-sign-in-link')
    ])

    await page.type('#inputJUsername', apply_detail['username'])
    await page.waitFor(500)
    await page.type('#inputJPassword', 'qaz2020')

    await asyncio.wait([
        page.waitForNavigation({'timeout': 60000}),
        page.click('#loginPane > div > form > div:nth-child(2) > div.col-sm-2.sign-in-button-cell > button'),
        # page.hover('#loginPane > div > form > div:nth-child(2) > div.col-sm-2.sign-in-button-cell > button'),
        # page.mouse.click(page.mouse._x,page.mouse._y)
    ])
    print('强制等待')
    await page.waitFor(10000)
    print('强制等待完成')

    await page.waitForSelector('#applyForm')
    print('等待出线')


    sign = await page.querySelector('#beginApplicationButton')
    await sign.click()

    print('页面跳转成功')
    await page.waitFor(10000)
    # Enrollment
    print('Enrollment')
    await page.select('#inputTermId', 'CAP_3870')
    await page.waitFor(1000)
    await page.select('#inputEduGoal', 'C')
    await page.waitFor(1000)
    await page.select('#inputMajorCategory', 'School of Humanities Languages Fine and Performing Arts')
    await page.waitFor(1000)
    await page.select('#inputMajorId', 'CAP_13694')
    await page.waitFor(1000)
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # account
    await page.waitFor(3000)
    print('account')

    await page.waitFor(3000)
    await page.click('#inputAddressSame')
    await page.waitFor(1000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # education
    print('education')
    await page.waitFor(10000)
    await page.select('#inputEnrollmentStatus', '1')
    await page.waitFor(1000)
    await page.select('#inputHsEduLevel', '2')
    await page.waitFor(1000)
    await page.click('#inputHsAttendance3')
    await page.waitFor(1000)
    await page.waitFor(3000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # citizenship/military
    print('citizenship')
    await page.waitFor(2000)
    await page.select('#inputCitizenshipStatus', '1')
    await page.waitFor(1000)
    await page.select('#inputMilitaryStatus', '1')
    await page.waitFor(1000)
    await page.waitFor(3000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # residency
    print('residency')
    await page.waitFor(10000)
    await page.click('#inputCaRes2YearsYes')
    await page.waitFor(1000)
    await page.click('#inputIsEverInFosterCareNo')
    await page.waitFor(1000)
    await page.click('#inputIsEverInFosterCareNo')
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # need & interests
    print('need')
    await page.waitFor(10000)
    await page.click('#inputEnglishYes')
    await page.waitFor(1000)
    await page.click('#inputFinAidInfoNo')
    await page.waitFor(1000)
    await page.click('#inputAssistanceNo')
    await page.waitFor(1000)
    await page.click('#inputAthleticInterest3')
    await page.waitFor(1000)

    await page.click('#inputHealthServices')
    await page.waitFor(1000)
    await page.click('#inputOnlineClasses')
    await page.waitFor(1000)
    await page.click('#inputCalWorks')
    await page.waitFor(1000)
    await page.click('#inputCareerPlanning')
    await page.waitFor(1000)
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # demographic information
    print('demographic')
    await page.waitFor(10000)
    await page.select('#inputGender', apply_detail['Gender'].capitalize())
    await page.waitFor(1000)
    await page.select('#inputTransgender', 'No')
    await page.waitFor(1000)
    await page.select('#inputOrientation', 'StraightHetrosexual')
    await page.waitFor(1000)
    await page.select('#inputParentGuardianEdu1', '3')
    await page.waitFor(1000)
    await page.select('#inputParentGuardianEdu2', '3')
    await page.waitFor(1000)
    await page.click('#inputHispanicNo')
    await page.waitFor(1000)
    await page.click('#inputRaceEthnicity800')
    await page.waitFor(1000)
    await page.click('#inputRaceEthnicity801')
    await page.waitFor(1000)
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # supplemental questions
    print('supplemental')
    await page.waitFor(10000)
    await page.waitFor(2000)
    await page.click('#YESNO_1_no')
    await page.waitFor(1000)
    await page.click('#YESNO_2_no')
    await page.waitFor(1000)
    await page.click('#YESNO_3_no')
    await page.waitFor(1000)
    await page.click('#YESNO_4_no')
    await page.waitFor(1000)
    await page.click('#YESNO_5_no')
    await page.waitFor(1000)
    await page.click('#YESNO_6_no')
    await page.waitFor(1000)
    await page.click('#YESNO_7_no')
    await page.waitFor(1000)
    await page.click('#YESNO_8_no')
    await page.waitFor(1000)
    await page.click('#_supp_CHECK_1')
    await page.waitFor(1000)
    await page.click('#_supp_CHECK_2')
    await page.waitFor(1000)
    await page.click('#_supp_CHECK_3')
    await page.waitFor(1000)
    await page.click('#_supp_CHECK_4')
    await page.waitFor(1000)
    await page.click('#YESNO_9_yes')
    # await page.click('#YESNO_10_yes')
    await page.waitFor(1000)
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click(
            '#applyForm > main > div.column.column2 > div.buttonBox.ccc-page-section > ol > li.save-continue > button')
    ])

    # submission
    print('submission')
    await page.waitFor(10000)
    await page.waitFor(2000)
    await page.click('#inputConsentYes')
    await page.waitFor(1000)
    await page.click('#inputESignature')
    await page.waitFor(1000)
    await page.click('#inputFinancialAidAck')
    await page.waitFor(1000)
    await page.waitFor(5000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#submit-application-button')
    ])

    job_page = await page.content()
    await page.waitFor(3000)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#applyForm > main > div.column.column1 > div.ccc-page-section.buttonBox > ol > li > button')
    ])
    await page.waitFor(5000)
    await page.click('#inputEnglishVerySatisfied')
    await page.waitFor(1000)
    await page.click('#RecommendYes')
    await page.waitFor(1000)
    await page.click('#applyForm > main > div.column.column1 > div.ccc-page-section.current-fieldset > button')
    await page.waitFor(2000)

    await browser.close()


def write_excel_file(folder_path,apply_detail):
    result_path = os.path.join(folder_path, "stu_register.xlsx")
    print(result_path)
    headers = ['email', 'Birthday']
    register_detail = []
    register_detail.append(apply_detail['email'])
    register_detail.append(apply_detail['Birthday'])
    print('***** 开始写入excel文件 ' + result_path + ' ***** \n')
    if os.path.exists(result_path):
        print('***** excel已存在，在表后添加数据 ' + result_path + ' ***** \n')
        workbook = xl.load_workbook(result_path)
        sheet = workbook.active
        sheet.append(register_detail)
        workbook.save(result_path)
    else:
        print('***** excel不存在，创建excel ' + result_path + ' ***** \n')
        workbook = xl.Workbook()
        workbook.save(result_path)

        sheet = workbook.active
        sheet.append(headers)

        sheet.append(register_detail)
        workbook.save(result_path)
    print('***** 生成Excel文件 ' + result_path + ' ***** \n')


def red_excel_file(folder_path):
    result_path = os.path.join(folder_path, "stu_apply.xlsx")
    apply_detail = {}
    wb2 = xl.load_workbook(result_path)
    sheet = wb2.active
    cell = sheet[1]
    cell1 = sheet[2]
    for i in range(len(cell)):
        apply_detail.update({cell[i].value:cell1[i].value})
    return apply_detail
    #sheet.delete_rows(2)
    #wb2.save(result_path)


def remove_excel_file(folder_path):
    result_path = os.path.join(folder_path, "stu_apply.xlsx")
    wb2 = xl.load_workbook(result_path)
    sheet = wb2.active
    sheet.delete_rows(2)
    wb2.save(result_path)
    print('删除成功')


def updata_excel(folder_path,flag,apply_detail):
    now = time.strftime("%d_%m_%Y")
    result_path = os.path.join(folder_path, "qt_%s.xlsx" % now)

    wb2 = xl.load_workbook(result_path)
    sheet = wb2.active

    for row in sheet:
        if row[1].value == apply_detail['email']:
            if flag == True:
                row[2].value = '注册成功'
            else:
                row[2].value = '失败'
    wb2.save(result_path)


if __name__ == '__main__':
    #api_url = "http://dps.kdlapi.com/api/getdps/?orderid=959731018601581&num=1&pt=1&dedup=1&sep=1"
    while True:
        flag = False
        apply_detail ={}
        try:
            apply_detail = red_excel_file('./')
            print(apply_detail)
        except:
            print('文件内无内容')

        if apply_detail is None:
            print('文件中没有已申请账号')
            time.sleep(300)
            continue

        try:
            asyncio.get_event_loop().run_until_complete(regist(apply_detail))
            write_excel_file('./',apply_detail)
            flag = True
        except :
            print('注册失败')
        finally:
            remove_excel_file('./')
            #updata_excel('./',flag)
            #print('删除成功')

        time.sleep(30)
