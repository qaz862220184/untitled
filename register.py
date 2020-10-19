import asyncio
import pyppeteer
import openpyxl as xl
import os,time
import requests
from pyppeteer import launch
import publicFun
import fake_useragent

pyppeteer.DEBUG = True

async def regist(apply_detail):
    url = 'https://www.opencccapply.net/gateway/apply?cccMisCode=531'
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
            #'--proxy-server=192.168.1.250:7890'
            #'--proxy-server={}'.format(proxy_ip)
        ]
    })
    #pid = browser.process.pid
    # 创建一个新的隐身浏览器上下文。这不会与其他浏览器上下文共享 cookie /缓存
    await browser.createIncognitoBrowserContext()
    await asyncio.sleep(1)
    page = await browser.newPage()
    width, height = publicFun.screen_size()
    await page.setViewport({  # 最大化窗口
        "width": width,
        "height": 900
    })
    await page.goto(url)
    await asyncio.wait([
        page.waitForNavigation({'timeout':1000*60}),
        page.click('#portal-sign-in-link')
    ])
    await page.type('#inputJUsername', apply_detail['userName'])
    await page.waitFor(500)
    await page.type('#inputJPassword', 'qaz2020')
    await asyncio.wait([
        page.waitForNavigation({'timeout': 60000}),
        page.click('#loginPane > div > form > div:nth-child(2) > div.col-sm-2.sign-in-button-cell > button'),
    ])
    await page.waitFor(10000)
    await page.waitForSelector('#applyForm')

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
    await page.select('#inputGender', apply_detail['gender'].capitalize())
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
    #return pid

if __name__ == '__main__':
    pid = ''
    while True:
        flag = False
        apply_detail = publicFun.get_user_detail(1)
        if apply_detail:
            try:
                asyncio.get_event_loop().run_until_complete(regist(apply_detail))
                publicFun.update_user_tag_time(apply_detail['email'],2)
                flag = True
                publicFun.logger.info('%s 申请成功'%apply_detail['email'])
            except Exception as e:
                publicFun.update_user_tag_time(apply_detail['email'], 4)
                publicFun.logger.info('%s,%s'%(apply_detail['email'],e))
            finally:
                #os.system('taskkill /pid' + str(pid) + '-t -f')
                time.sleep(30)
        else:
            time.sleep(300)
