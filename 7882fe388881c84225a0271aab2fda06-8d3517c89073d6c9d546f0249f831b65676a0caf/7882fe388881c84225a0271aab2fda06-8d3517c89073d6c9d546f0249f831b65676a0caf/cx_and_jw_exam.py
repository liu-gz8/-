import json
import os
import platform
import subprocess
import sys
from math import ceil
from time import sleep
from lxml import etree
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class login():
    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--headless')  #隐藏浏览器
        chrome_options.add_argument(
            '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36')  # 设置请求头的User-Agent
        chrome_options.add_argument('--disable-infobars')  # 禁用浏览器正在被自动化程序控制的提示
        self.driver = webdriver.Chrome(options=chrome_options,executable_path='core\chromedriver.exe')
        self.login_url = "http://passport2.chaoxing.com/login?fid=&newversion=true&refer=http%3A%2F%2Fi.chaoxing.com"

    def phone_login(self):
        phone = input('请输入手机号:')
        self.driver.find_element_by_css_selector("#phone").send_keys(phone)
        password = input('请输入密码:')
        self.driver.find_element_by_css_selector("#pwd").send_keys(password)
        self.driver.find_element_by_css_selector("#loginBtn").click()

    def jigou_login(self):
        # 点击其它登录
        self.driver.find_element_by_xpath('//*[@id="otherlogin"]').click()
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="inputunitname"]').send_keys('浙大宁波理工学院')
        sleep(0.75)
        uname = input('请输入学号或者工号:')
        self.driver.find_element_by_xpath('//*[@id="uname"]').send_keys(uname)
        sleep(0.75)
        password = input('请输入密码:')
        self.driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
        sleep(0.75)
        self.driver.find_element_by_xpath('//*[@id="numVerCode"]').screenshot('core\chaoxing_captcha.png')
        file_name = 'core\chaoxing_captcha.png'
        self.show_img(file_name)
        txtSecretCode = input('请输入验证码:')
        self.driver.find_element_by_xpath('//*[@id="vercode"]').send_keys(txtSecretCode)
        self.driver.find_element_by_xpath('//*[@id="loginBtn"]').click()

    def show_img(self, file_name):
        userPlatform = platform.system()
        if userPlatform == 'Darwin':  # Mac
            subprocess.call(['open', file_name])
        elif userPlatform == 'Linux':  # Linux
            subprocess.call(['xdg-open', file_name])
        else:  # Windows
            os.startfile(file_name)

    def QR_login(self):
        self.driver.find_element_by_xpath('//*[@id="quickCode"]').screenshot('core\QR.png')
        file_name = 'core\QR.png'
        self.show_img(file_name)

    def check_state(self):
        self.driver.get('http://i.chaoxing.com/')
        sleep(1)
        f = '账号管理' in self.driver.page_source
        if f == True:
            print('登录成功')
            return True
        else:
            print('登录失败')
            return False

    def save_cookie(self):
        cookie_list = self.driver.get_cookies()
        jsonCookies = json.dumps(cookie_list)
        with open('core\chaoxing_cookies.json', 'w') as f:
            f.write(jsonCookies)
        print('保存cookie成功')

    def read_cookie(self):
        with open('core\chaoxing_cookies.json', 'r') as f:
            list_cookie = json.loads(f.read())
        return list_cookie

    def check_cookie(self, list_cookie):
        self.driver.get(self.login_url)
        self.driver.delete_all_cookies()
        for cookie in list_cookie:
            self.driver.add_cookie(cookie)
        sleep(2)
        self.driver.get(self.login_url)
        sleep(1)
        self.driver.get('http://i.chaoxing.com/')
        f = self.check_state()
        if f == True:
            print('cookie有效')
            return True
        else:
            print('cookie无效')
            return False

    def check_file(self):
        f = os.path.exists('core\chaoxing_cookies.json')
        if f == True:
            list_cookie = self.read_cookie()
            f = self.check_cookie(list_cookie)
            if f == True:
                return True
            else:
                return False
        return False

    def login(self):
        self.driver.maximize_window()
        self.driver.get(self.login_url)
        f = self.check_file()
        if f == False:
            while True:
                print('请选择登录方式:\n'
                      '1为电话号码登录\n'
                      '2为学号或者工号登录\n'
                      '3为扫码登录\n'
                      '4为验证码登录\n'
                      '输入其它字符为退出')
                num = input('请输入数字:')
                if num == '1':
                    self.phone_login()
                elif num == '2':
                    self.jigou_login()
                elif num == '3':
                    self.QR_login()
                elif num == '4':
                    pass
                else:
                    sys.exit()
                f = self.check_state()
                if f == True:
                    self.save_cookie()
                    break
        else:
            pass
        return self.driver


class exam():

    def __init__(self):
        self.driver = login().login()
        self.dict = {}
        self.subject_xpath_dict = {
            "单选题": {'delete': '//*[@id="typetrid1"]/span[2]/a', 'score': '//*[@id="0_score"]',
                    'subject_num': '//*[@id="0_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "多选题": {'delete': '//*[@id="typetrid2"]/span[2]/a', 'score': '//*[@id="1_score"]',
                    'subject_num': '//*[@id="1_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "填空题": {'delete': '//*[@id="typetrid3"]/span[2]/a', 'score': '//*[@id="2_score"]',
                    'subject_num': '//*[@id="2_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "判断题": {'delete': '//*[@id="typetrid4"]/span[2]/a', 'score': '//*[@id="3_score"]',
                    'subject_num': '//*[@id="3_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "简答题": {'delete': '//*[@id="typetrid5"]/span[2]/a', 'score': '//*[@id="4_score"]',
                    'subject_num': '//*[@id="4_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "名词解释": {'delete': '//*[@id="typetrid6"]/span[2]/a', 'score': '//*[@id="5_score"]',
                     'subject_num': '//*[@id="5_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "论述题": {'delete': '//*[@id="typetrid7"]/span[2]/a', 'score': '//*[@id="6_score"]',
                    'subject_num': '//*[@id="6_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "计算题": {'delete': '//*[@id="typetrid8"]/span[2]/a', 'score': '//*[@id="7_score"]',
                    'subject_num': '//*[@id="7_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "分录题": {'delete': '//*[@id="typetrid9"]/span[2]/a', 'score': '//*[@id="9_score"]',
                    'subject_num': '//*[@id="9_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "资料题": {'delete': '//*[@id="typetrid10"]/span[2]/a', 'score': '//*[@id="10_score"]',
                    'subject_num': '//*[@id="10_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "连线题": {'delete': '//*[@id="typetrid11"]/span[2]/a', 'score': '//*[@id="11_score"]',
                    'subject_num': '//*[@id="11_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "排序题": {'delete': '//*[@id="typetrid13"]/span[2]/a', 'score': '//*[@id="13_score"]',
                    'subject_num': '//*[@id="13_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "完型填空": {'delete': '//*[@id="typetrid14"]/span[2]/a', 'score': '//*[@id="14_score"]',
                     'subject_num': '//*[@id="14_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "阅读理解": {'delete': '//*[@id="typetrid15"]/span[2]/a', 'score': '//*[@id="15_score"]',
                     'subject_num': '//*[@id="15_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "程序题": {'delete': '//*[@id="typetrid17"]/span[2]/a', 'score': '//*[@id="17_score"]',
                    'subject_num': '//*[@id="17_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "口语题": {'delete': '//*[@id="typetrid18"]/span[2]/a', 'score': '//*[@id="18_score"]',
                    'subject_num': '//*[@id="18_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "听力题": {'delete': '//*[@id="typetrid19"]/span[2]/a', 'score': '//*[@id="19_score"]',
                    'subject_num': '//*[@id="19_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "共用选项题": {'delete': '//*[@id="typetrid20"]/span[2]/a', 'score': '//*[@id="20_score"]',
                      'subject_num': '//*[@id="20_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
            "其它": {'delete': '//*[@id="typetrid21"]/span[2]/a', 'score': '//*[@id="8_score"]',
                   'subject_num': '//*[@id="8_TypeDiv"]/li[2]/div[2]/div[1]/p[2]/input[2]'},
        }
        self.choice_dict = {
            "单选题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[1]/input',
            "多选题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[2]/input',
            "填空题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[3]/input',
            "判断题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[4]/input',
            "简答题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[5]/input',
            "名词解释": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[6]/input',
            "论述题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[7]/input',
            "计算题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[8]/input',
            "分录题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[9]/input',
            "资料题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[10]/input',
            "连线题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[11]/input',
            "排序题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[12]/input',
            "完型填空": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[13]/input',
            "阅读理解": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[14]/input',
            "程序题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[15]/input',
            "口语题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[16]/input',
            "听力题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[17]/input',
            "共用选项题": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[18]/input',
            "其它": '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[1]/label[19]/input',
        }

    def template_exam(self, url, paper_num):
        # 输入试卷标题
        print('标题字符数为4到40个字符')
        while True:
            title = input('请输入标题:')
            if len(title) >= 4 and len(title) <= 40:
                break
            else:
                print('输入有误,标题标题字符数为4到40个字符')
        for p in range(paper_num):
            try:
                self.driver.get(url)
                self.driver.refresh()
                # 清理文本输入框
                self.driver.find_element_by_xpath('//*[@id="title"]').clear()
                # 输入标题
                self.driver.find_element_by_xpath('//*[@id="title"]').send_keys(title)
                # 输入试卷数目
                self.driver.find_element_by_xpath('//*[@id="pageNum"]').send_keys(1)
                # 点击保存
                self.driver.find_element_by_xpath('//*[@id="actionTab"]/a[1]').click()
                sleep(2)
                # 定位到弹窗
                alert = self.driver.switch_to.alert
                sleep(1)
                # 确认组卷进行确定
                alert.accept()
                print("第{}张试卷组卷成功!!!".format(p + 1))
            except:
                print("第{}张试卷组卷失败!!!".format(p + 1))
            sleep(2)

    def random_exam(self, url):
        self.driver.get(url)
        sleep(1.5)
        # 前往资料库
        self.driver.find_element_by_xpath('/html/body/div[4]/div[2]/div[1]/div[2]/ul/li[4]/a').click()
        # 切换资料库页面
        self.driver.switch_to.window(self.driver.window_handles[-1])
        # 前往题库
        self.driver.find_element_by_xpath('//*[@id="RightCon"]/div/div/div[1]/ul/li[2]/a').click()
        html = self.driver.page_source
        page_num = self.get_pageNum_subject(html)
        print("正在统计第1页题目")
        self.statistical(html)
        for p in range(2, page_num + 1):
            print("正在统计第{}页题目".format(p))
            self.driver.find_element_by_xpath(
                '//div[@id="pagination"]/a[@onclick="changePageAdd({})"]'.format(p)).click()
            html = self.driver.page_source
            self.statistical(html)
            sleep(1.5)
        self.driver.find_element_by_xpath('//*[@id="RightCon"]/div/div/div[1]/ul/li[4]/a').click()
        # 等待"创建试卷"关键句出现
        sleep(1)
        WebDriverWait(self.driver, 3).until(EC.visibility_of_element_located((By.LINK_TEXT, "创建试卷")))
        # 点击创建试卷
        self.driver.find_element_by_xpath('//*[@id="qform"]/a[3]').click()
        # 等待"下一步"关键句出现
        sleep(1)
        WebDriverWait(self.driver, 3).until(EC.visibility_of_element_located((By.LINK_TEXT, "下一步")))
        # 点击自动创建试卷按钮
        self.driver.find_element_by_xpath('//*[@id="chooseForm"]/div/p[2]/label/input').click()
        # 点击下一步
        self.driver.find_element_by_xpath('//*[@id="chooseForm"]/div/div/a').click()
        # 等待更多题型加载
        sleep(1)
        WebDriverWait(self.driver, 3).until(EC.visibility_of_element_located((By.LINK_TEXT, "更多题型")))
        # 切换页面为当前
        self.driver.switch_to.window(self.driver.window_handles[-1])
        # 试卷数量设置为1
        self.driver.find_element_by_xpath('//*[@id="pageNum"]').send_keys(1)
        # 设置试卷标题
        print("请注意试卷标题至少为4个字符")
        title = input("请输入试卷标题:")
        self.driver.find_element_by_xpath('//*[@id="title"]').send_keys(title)
        sleep(2)
        # 随机试卷初始化含有以下三种类型,把它们全部删除
        origin_list = ['单选题', '多选题', '填空题']
        for i in origin_list:
            self.driver.find_element_by_xpath(self.subject_xpath_dict[i]['delete']).click()
        sleep(2)
        print("------------------------------")
        print("您题库含有试题及数量如下:")
        self.type_list = []
        x = 0
        for key, value in self.dict.items():
            if value != 0:
                self.type_list.append(key)
                print("序号:{},您有 {}:{}道".format(x, key, value))
                x += 1
        print("------------------------------")
        # 选择课程:
        print("请输入你要选择试题的序号,输入负数代表结束！")
        choice_list = []
        while True:
            num = int(input("请输入数字:"))
            if num >= 0:
                choice_list.append(num)
            else:
                break
        self.choice_subject(choice_list)
        sleep(1)
        print("下面是每个题型的总分以及题型数量的信息填写,请您分配好分数\n同时试卷默认总分为100\n如果计算出总分不为100,将会自动按照比例更改分数\n并满足总分为100")
        self.input_info(choice_list)
        n = input('是否同时保存为模板\n'
                  '是输入1\n'
                  '否输入0\n'
                  '请输入:')
        # 确定同时保持为模板
        if n == '1':
            self.driver.find_element_by_xpath('//*[@id="savePaperTemplateCheck"]').click()
        sleep(2)
        # 点击保存
        self.driver.find_element_by_xpath('//*[@id="actionTab"]/a[1]').click()
        sleep(2)
        # 定位到弹窗
        alert = self.driver.switch_to.alert
        sleep(1)
        # 确认组卷进行确定
        alert.accept()
        sleep(2)

    # 输入每个题型的分数和数目
    def input_info(self, choice_list):
        score_list = []
        for num in choice_list:
            score = input("请输入{}总分数:".format(self.type_list[num]))
            score_list.append(int(score))
            self.driver.find_element_by_xpath(self.subject_xpath_dict[self.type_list[num]]['score']).send_keys(score)
            sleep(0.75)
            subject_num = input('请输入{}的题目数:'.format(self.type_list[num]))
            self.driver.find_element_by_xpath(self.subject_xpath_dict[self.type_list[num]]['subject_num']).send_keys(
                subject_num)
            sleep(0.75)
        Sum = sum(score_list)
        if Sum != 100:
            print('试卷总分不为一百，正在进行更改')
            for num in range(len(score_list)):
                if num != len(score_list) - 1:
                    # 按照比例进行分配
                    score_list[num] = int((score_list[num] / Sum) * 100)
                else:
                    score_list[num] = 100 - sum(score_list[:-1])
        else:
            pass
        n = 0
        for num in choice_list:
            score = score_list[n]
            self.driver.find_element_by_xpath(self.subject_xpath_dict[self.type_list[num]]['score']).clear()
            sleep(0.25)
            self.driver.find_element_by_xpath(self.subject_xpath_dict[self.type_list[num]]['score']).send_keys(score)
            n += 1
            sleep(0.75)
        print('更改完成')

    # 题型选择
    def choice_subject(self, choice_list):
        # 点击更多题型
        self.driver.find_element_by_xpath('//*[@id="newMore"]').click()
        for num in choice_list:
            # 对每个题型打上对号
            self.driver.find_element_by_xpath(self.choice_dict[self.type_list[num]]).click()
        # 点击确定
        self.driver.find_element_by_xpath(
            '//*[@id="setPaperStructure"]/div[1]/div/div[2]/div[2]/a[1]/span').click()

    # 题库信息统计
    def statistical(self, html):
        html = etree.HTML(html)
        tr_list = html.xpath('//*[@id="tableId"]/tr')
        for tr in tr_list:
            key = tr.xpath('td[3]/text()')
            if key != []:
                key = key[0].strip()
                self.dict[key] += 1
            else:
                pass

    # 对网页处理
    def get_pageNum_subject(self, html):
        html = etree.HTML(html)
        str_num = html.xpath('//*[@id="RightCon"]/div/div/div[4]/span[2]/text()')[0]
        page_num = ceil(int(str_num) / 20)
        # print(page_num)
        option_list = html.xpath('//*[@id="qTypeSelect"]/option')
        for i in range(1, len(option_list)):
            key = option_list[i].xpath('text()')[0].strip()
            self.dict[key] = 0
        return page_num
