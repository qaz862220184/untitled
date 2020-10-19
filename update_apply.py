import asyncio
import pyppeteer
from pyppeteer import launch
import publicFun
import traceback


async def delete_register(userName):
    url = 'https://www.opencccapply.net/gateway/apply?cccMisCode=531'
    browser = await launch({
        'executablePath': pyppeteer.launcher.executablePath(),
        #'headless': False,
        'dumpio': True,
        'autoClose': False,
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
            #'--proxy-server=192.168.1.239:24002'
        ]
    })
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
        page.click('#portal-sign-in-link'),
        page.waitForNavigation({'timeout':1000*60})
    ])
    await page.type('#inputJUsername', userName)
    await page.waitFor(500)
    await page.type('#inputJPassword', 'qaz2020')
    await asyncio.wait([
        page.waitForNavigation({'timeout': 60000}),
        page.click('#loginPane > div > form > div:nth-child(2) > div.col-sm-2.sign-in-button-cell > button'),
    ])
    await page.waitFor(10000)
    await page.waitForSelector('#applyForm')

    b = await page.querySelector('.portlet-form-label')
    if b:
        await page.click('#applyForm > main > div.column.column1 > div:nth-child(6) > div > div > div.card-body > div.actions > button.btn.icon-btn.btn-link.secondary-action.delete')
        await page.waitFor(2000)
        await page.click('#delete-confirmation-ok-button')
        await page.waitFor(1000)

    #await browser.close()


if __name__ == '__main__':
    while True:
        conn = publicFun.connect_db()
        sql = 'select userName from user_detail where tag=4 limit 1'
        sql1 = 'update user_detail set tag=1 where userName= %s'
        cursor = conn.cursor()
        cursor.execute(sql)
        userName = cursor.fetchone()
        print(userName)
        try:
            asyncio.get_event_loop().run_until_complete(delete_register(userName))
            result = cursor.execute(sql1,userName)
            conn.commit()
            print(result)
        except:
            traceback.print_exc()
            print('更改状态失败')
        finally:
            cursor.close()
            conn.close()


