import base64
import os
import platform
import re
import subprocess
import sys
import time
from http import cookiejar

import muggle_ocr
import requests
from PIL import Image


# 超星登录
class chaoxing_login(object):

    def __init__(self):
        self.session = requests.session()
        self.session.cookies = cookiejar.LWPCookieJar(filename='core/chaoxing_cookies.txt')
        self.login_headers = {
            'Origin': 'http://passport2.chaoxing.com',
            'Referer': 'http://passport2.chaoxing.com/login?loginType=3&newversion=true&fid=-1&refer=http%3A%2F%2Fi.chaoxing.com',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36',
            'X-Requested-With': 'XMLHttpRequest',
            'Host': 'passport2.chaoxing.com',
        }
        # 登录完成的请求头
        self.login_complete_headers = {
            'Host': 'i.chaoxing.com',
            'Referer': 'http://passport2.chaoxing.com/',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36',
        }
        self.account_url = 'http://passport2.chaoxing.com/unitlogin'
        # 机构登录data
        self.account_data = {
            'fid': '18638',
            'uname': '',
            'numcode': '',
            'password': '',
            'refer': 'http%3A%2F%2Fi.chaoxing.com',
            't': 'true',
        }
        self.phone_url = 'http://passport2.chaoxing.com/fanyalogin'
        # 手机号登录data
        self.phone_data = {
            'fid': '-1',
            'uname': '',
            'password': '',
            'refer': 'http%3A%2F%2Fi.chaoxing.com',
            't': 'true',
            'forbidotherlogin': '0',
        }
        self.QR_code_headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36',
            'Host': 'passport2.chaoxing.com',
            'Referer': 'http://passport2.chaoxing.com/login?fid=&newversion=true&refer=http%3A%2F%2Fi.chaoxing.com',
            'Upgrade-Insecure-Requests': '1',
        }
        self.session.headers = self.login_headers

    # 图片展示
    def show_img(self, file_name):
        userPlatform = platform.system()
        if userPlatform == 'Darwin':  # Mac
            subprocess.call(['open', file_name])
        elif userPlatform == 'Linux':  # Linux
            subprocess.call(['xdg-open', file_name])
        else:  # Windows
            os.startfile(file_name)

    # 13位时间戳
    def get_time_stamp(self):
        time_stamp = str(int(time.time() * 1000))
        return time_stamp

    # 获取验证码
    def get_captcha(self):
        print('验证码获取中......')
        captcha_url = "http://passport2.chaoxing.com/num/code?{}".format(self.get_time_stamp())
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.104 Safari/537.36",
        }
        response = self.session.get(url=captcha_url, headers=headers)
        if response.status_code == 200:
            print('获取验证码成功')
            content = response.content
            with open("core\chaoxing_captcha.png", "wb") as f:
                f.write(content)
            self.show_img(file_name='core\chaoxing_captcha.png')
        else:
            print('抱歉,获取验证码失败\n'
                  '程序将自动终止，请重新打开程序')
            sys.exit()

    # 密码加密
    def password_encrypt(self, password):
        password = base64.b64encode(password.encode())
        password = password.decode()
        return password

    # 检查cookies
    def check_cookies(self):
        self.session.headers = self.login_complete_headers
        try:
            # 加载cookies
            self.session.cookies.load(ignore_discard=True)

            url = "http://i.chaoxing.com/"
            response = self.session.get(url=url, allow_redirects=True)
            if response.status_code == 200:
                # print(response.text)
                return True
            else:
                return False
        except FileNotFoundError:
            return "无cookie文件"

    # 账号密码输入
    def input(self):
        uname = input('请输入账户：')
        password = input('请输入密码：')
        password = self.password_encrypt(password)
        if self.num == '1':
            self.account_data['uname'] = uname
            self.account_data['password'] = password
            self.get_captcha()
            numcode = input('请输入验证码：')
            self.account_data['numcode'] = numcode
        elif self.num == '2':
            self.phone_data['uname'] = uname
            self.phone_data['password'] = password

    # 扫码登入所需的uuid,enc
    def get_uuid_enc(self):
        url = 'http://passport2.chaoxing.com/login'
        params = {
            'fid': '',
            'newversion': 'true',
            'refer': 'http://i.chaoxing.com',
        }
        response = self.session.get(url=url, params=params)
        text = response.text
        self.uuid = re.findall('<input.*?value="(.*?)".*?"uuid"/>', text)[0]
        self.enc = re.findall('<input.*?value="(.*?)".*?"enc"/>', text)[0]

    def QR_png(self):
        print('二维码获取中......')
        self.get_uuid_enc()
        url = 'http://passport2.chaoxing.com/createqr'
        params = {
            'uuid': self.uuid,
            'fid': '-1',
        }
        self.session.headers = self.QR_code_headers
        response = self.session.get(url=url, params=params)
        if response.status_code == 200:
            print('二维码获取成功')
            content = response.content
            # 这里照片数据为bytes形式，所以为'wb'
            with open('core\QR.png', 'wb') as f:
                f.write(content)
            print('二维码保存成功')
            self.show_img(file_name='core\QR.png')
            # self.getauthstatus()
        else:
            print('抱歉,获取二维码失败\n'
                  '程序将自动终止，请重新打开程序')
            sys.exit()

    # 扫码登录状态获取
    def getauthstatus(self):
        count = 0
        while True:
            getauthstatus_url = 'http://passport2.chaoxing.com/getauthstatus'
            data = {
                'enc': self.enc,
                'uuid': self.uuid,
            }
            response = self.session.post(url=getauthstatus_url, data=data)
            text = response.text
            if '未登录' not in text:
                dic = response.json()
                if dic['status'] == False:
                    self.uid = dic['uid']
                    self.nickname = dic['nickname']
                    print('用户==》{}《==请您确认登录'.format(self.nickname))
                elif dic['status'] == True:
                    print('用户==》{}《==您已确认登录'.format(self.nickname))
                    return True
            else:
                print('不要让人家苦苦等待嘛,请您扫一下二维码')
            # 请求50次，二维码将刷新一次
            count += count
            if count == 150:
                return False
            time.sleep(1)

    # 登入信息判断,扫码登录不能用
    def login_info_judge(self):
        response = self.session.post(url=self.url, data=self.data)
        text = response.text
        if 'captcha is incorrect' in text or '验证码错误' in text:
            return '验证码错误'
        elif 'account or passport is wrong' in text or '用户名或密码错误' in text:
            return '用户名或密码错误'
        else:
            return True

    # 机构登录
    def account_login(self):
        while True:
            self.input()
            mes = self.login_info_judge()
            if mes != True:
                print('登录失败！')
                print(mes)
                if mes == '验证码错误':
                    self.get_captcha()
                    numcode = input('请输入验证码：')
                    self.account_data['numcode'] = numcode
                    mes = self.login_info_judge()
                else:
                    mes = self.login_info_judge()
            else:
                print('登录成功！')
                self.session.cookies.save()
                print('cookie保存成功！')
                return ''

    # 号码登录
    def phone_sign(self):
        while True:
            self.input()
            mes = self.login_info_judge()
            if mes != True:
                print('登录失败！')
                print(mes)
                mes = self.login_info_judge()
            else:
                print('登录成功！')
                self.session.cookies.save()
                print('cookie保存成功！')
                return ''

    # 扫码登录
    def QR_code_sign(self):
        self.QR_png()
        while True:
            judge_info = self.getauthstatus()
            if judge_info == True:
                break
            else:
                self.QR_png()
        return ''

    # 登录入口
    def login(self):
        b = self.check_cookies()
        if b == True:
            print('超星cookie有效')

        else:
            print('不建议用机构方式登录，会涉及手动输入验证码，推荐扫码登录')
            if b == False:
                print('超星cookies失效')
            else:
                print('没有超星cookie文件')
            self.session.headers = self.login_headers
            print('1代表机构登录\n'
                  '2代表号码登录\n'
                  '3代表扫码登录')
            self.num = input('请选择登入方式:')
            if self.num == '1':
                self.url = self.account_url
                self.data = self.account_data
                self.account_login()
            elif self.num == '2':
                self.url = self.phone_url
                self.data = self.phone_data
                self.phone_sign()
            elif self.num == '3':
                self.QR_code_sign()
                self.session.headers = self.login_complete_headers
                url = "http://i.chaoxing.com/"
                response = self.session.get(url=url, allow_redirects=True)
                if response.status_code == 200:
                    print('登录成功')
                    # print(response.text)
                    self.session.cookies.save()
                    print('cookie保存成功')
        return self.session


# 教务系统
class Educational_administration_system_login():

    def __init__(self):
        self.session = requests.session()
        self.sdk = muggle_ocr.SDK(model_type=muggle_ocr.ModelType.Captcha)
        self.session.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.150 Safari/537.36'
        }
        self.login_data = {
            '__LASTFOCUS': '',
            '__VIEWSTATE': '',
            '__VIEWSTATEGENERATOR': '9BD98A7D',
            '__EVENTTARGET': '',
            '__EVENTARGUMENT': '',
            'txtUserName': '',
            'TextBox2': '',
            'txtSecretCode': '',
            'RadioButtonList1': '学生',
            'Button1': '登录',
        }

    def get_VIEWSTATE_and_SafeKey(self):

        url = 'http://jwxt.nit.net.cn/'
        response = self.session.get(url=url)
        text = response.text
        self.VIEWSTATE = re.findall('<input.*?name="__VIEWSTATE".*?value="(.*?)" />', text)[0]
        self.SafeKey = re.findall('<img id="icode".*?SafeKey=(.*?)".*', text)[0]
        return ''

    def get_captcha(self):
        url = 'http://jwxt.nit.net.cn/CheckCode.aspx?SafeKey={}'.format(self.SafeKey)
        response = self.session.get(url=url)
        content = response.content
        print('正在获取验证码......')
        with open('core\jw_captcha.jpg', 'wb') as f:
            f.write(content)
        print('获取验证码成功.')
        txtSecretCode = self.captcha_ocr()
        return txtSecretCode

    def captcha_ocr(self):
        def binarizing(img, threshold):
            """传入image对象进行灰度、二值处理"""
            img = img.convert("L")  # 转灰度
            pixdata = img.load()
            w, h = img.size
            # 遍历所有像素，大于阈值的为黑色
            for y in range(h):
                for x in range(w):
                    if pixdata[x, y] < threshold:
                        pixdata[x, y] = 0
                    else:
                        pixdata[x, y] = 255
            return img

        img = Image.open('core\jw_captcha.jpg')
        img = img.convert("L")
        img.save('core\jw_gray.jpg')
        with open('core\jw_gray.jpg', "rb") as f:
            b = f.read()

        txtSecretCode = self.sdk.predict(image_bytes=b)
        return txtSecretCode


    def login(self):
        url = 'http://jwxt.nit.net.cn/'
        txtUserName = input('请输入您的账号:')
        TextBox2 = input('请输入您的密码:')
        self.login_data['txtUserName'] = txtUserName
        self.login_data['TextBox2'] = TextBox2
        while True:
            self.get_VIEWSTATE_and_SafeKey()
            txtSecretCode = self.get_captcha()
            if len(txtSecretCode) >= 5:
                txtSecretCode = txtSecretCode[:5]
            print(txtSecretCode)
            self.login_data['__VIEWSTATE'] = self.VIEWSTATE
            self.login_data['txtSecretCode'] = txtSecretCode
            response = self.session.post(url=url, data=self.login_data)
            text = response.text
            if '安全退出' in text:
                print('登录成功')
                url = response.url
                return self.session, url


