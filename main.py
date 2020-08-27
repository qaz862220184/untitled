# coding:utf-8


from fastapi import FastAPI
import get_stu_edu
import apply
import random
import register
import asyncio
import user
app = FastAPI()
import nest_asyncio
# 获取随机信息
@app.get('/getuser')
async def root():

    detail = []

    headers = []

    #nest_asyncio.apply()
    #asyncio.new_event_loop().run_until_complete(user.main())
    user_detail = await user.get_user_detail()
    #print(user_detail)
    user.write_excel_file("./",user_detail)
    user.write_ui_qt('./',user_detail)

# 学生申请
@app.get('/apply')
async def root1():

    flag = False
    username = ''.join(random.sample("1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", 8))
    user_detail = apply.red_excel_file('./')
    user_detail.update({'username': username})
    print(user_detail)
    try:
        tag = await apply.main(user_detail)
        if tag == True:
            apply.write_excel_file("./",user_detail)
            flag = True
        else:
            print('申请失败')
    except:
        print('申请失败1')
    finally:
        apply.updata_excel('./',flag,user_detail)


# 信息注册
@app.get('/register')
async def root2():
    flag = False
    print('开始获取申请信息')
    apply_detail = {}
    try:
        apply_detail = register.red_excel_file('./')
        print(apply_detail)
    except:
        print('文件内无内容')
    if apply_detail is None:
        print('文件中没有已申请账号')
    try:
        await register.regist(apply_detail)
        register.write_excel_file('./',apply_detail)
        flag = True
    except:
        print('注册失败')
    finally:
        register.remove_excel_file('./')
        register.updata_excel('./',flag,apply_detail)


# 获取edu邮箱
@app.get('/getemail')
async def root3():
    stu_register = get_stu_edu.red_excel_file('./')
    print(stu_register)
    email = stu_register['email']
    Birthday = stu_register['Birthday']
    print(email, Birthday)
    m, d, y = Birthday.split('/')
    m = m.zfill(2)
    d = d.zfill(2)
    y = y[-2:]
    edu_pwd = m + d + y
    stu_id = ''
    edu_email = ''
    try:
        stu_id, edu_email = await get_stu_edu.get_edu(email)
    except Exception as e:
        print(e)
    if stu_id and edu_email and edu_pwd:
        get_stu_edu.write_excel_file('./',stu_id,edu_email,edu_pwd)
        get_stu_edu.remove_excel_file('./')
    else:
        print('未获得完整学生邮箱信息')

