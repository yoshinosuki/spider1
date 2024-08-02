import datetime
import os
import json
import time
import PyOfficeRobot
import schedule
from PyOfficeRobot.api.chat import wx
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import InvalidArgumentException


# 全局变量

who = '芳乃'  # 默认用户
watch_time = 3600  # 播放时间，单位s


# 定义接收
def receive_message1(who='文件传输助手', txt='userMessage.txt', output_path='./'):
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
    current_time = datetime.datetime.now()
    print('--' * 88)
    with open(os.path.join(output_path, txt), 'w') as output_file:
        output_file.write(str(receive_msg))
        output_file.write('\n')


# 将cookie转换成selenium适用的格式
def convert_cookies(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as file:
        data = json.load(file)
    cookies = data.get('cookies', [])
    selenium_cookies = []
    for cookie in cookies:
        selenium_cookie = {
            'name': cookie['name'],
            'value': cookie['value'],
            'domain': cookie['domain'],
            'path': cookie.get('path', '/'),
            'secure': cookie.get('secure', False),
            'httpOnly': cookie.get('httpOnly', False),
            'expiry': int(cookie['expirationDate']) if 'expirationDate' in cookie else None
        }
        if selenium_cookie['expiry'] is None:
            del selenium_cookie['expiry']
        selenium_cookies.append(selenium_cookie)
    with open(output_file, 'w', encoding='utf-8') as file:
        json.dump(selenium_cookies, file, ensure_ascii=False, indent=4)


# 读取cookie
def load_cookies(driver, cookies_file):
    with open(cookies_file, 'r', encoding='utf-8') as file:
        cookies = json.load(file)
        for cookie in cookies:
            driver.add_cookie(cookie)


# 检查并点击相关元素
def click_if_exists(driver, xpath):
    try:
        element = WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, xpath)))  # 等待元素可见
        ActionChains(driver).move_to_element(element).click().perform()  # 直接点击
        print(f"Clicked on element: {xpath}")
    except TimeoutException:
        print(f"Element not found (timeout): {xpath}")  # 增加超时提示
    except NoSuchElementException:
        print(f"Element not found: {xpath}")  # 增加找不到元素的提示
    except Exception as e:
        print(f"An error occurred: {str(e)}")


# 使用微信启动
def wxstart():
    global who
    time.sleep(1)
    PyOfficeRobot.chat.send_message(who=who, message='初始化中')
    time.sleep(1)
    PyOfficeRobot.chat.send_message(who=who, message='你好，请发送')
    time.sleep(20)
    receive_message1(who, 'web.txt', './')
    time.sleep(1)
    PyOfficeRobot.chat.send_message(who=who, message='报名停止')
    time.sleep(5)
    run_browser_task()


def run_browser_task():
    try:
        global watch_time
        # 读取web.txt里的web地址
        with open('web.txt', 'r') as file:
            lines = file.readlines()

        for line in lines:
            line = line.strip()
            question_mark_index = line.find('?')
            if question_mark_index != -1:
                url = line[:question_mark_index]
            else:
                url = line
            print(url)

        # 处理cookie
        input_file = 'cookies.json'
        output_file = 'selenium_cookies.json'
        convert_cookies(input_file, output_file)
        print(f"Got the Cookies {output_file}")

        # 运行网页
        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--ignore-ssl-errors')
        options.add_argument(
            "user-agent=Mozilla/5.0 (Linux; Android 8.0.0; SM-T835) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")
        options.add_argument('--window-size=1280,800')

        driver = webdriver.Chrome(options=options)
        driver.get(url)

        # 等待页面加载完成（使用显式等待）
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # Load cookies
        load_cookies(driver, output_file)
        driver.refresh()
        time.sleep(5)
        print(f"start!")

        # 持续检测元素并点击
        start_time = time.time()
        while True:
            click_if_exists(driver, '//*[@id="bigScreenWrap"]/section/div/i')  # 开始播放1
            click_if_exists(driver, '//*[@id="main_page"]/section/div/i')  # 开始播放2
            click_if_exists(driver, '//*[@id="app"]/div[3]/div/footer/img')  # 不关注
            click_if_exists(driver, '//*[@id="bigScreenWrap"]/section/div/section/div/button')  # 重新尝试
            # 检查是否达到7200秒
            if time.time() - start_time > watch_time:
                print("Reached set time. Closing the browser.")
                break
            time.sleep(5)  # 每5秒检测一次
        # 检查缓存文件是否存在
        if os.path.exists(output_file):
            os.remove(output_file)
            print(f"{output_file} 缓存已成功删除")
        else:
            print(f"{output_file} 生成失败")

        # 退出
        print(f'完成任务，退出该次直播')
        driver.quit()

    except InvalidArgumentException as e:
        print("出现了无效参数异常:", e)
        # 这里可以选择记录日志或者进行其他处理
    except Exception as e:
        print("发生了其他异常:", e)


if __name__ == '__main__':
    # 定时任务:至少一分钟前开启！！！！，每个计划运行前要等待前计划运行完毕
    schedule.every().day.at("11:24").do(wxstart)
    # schedule.every().day.at("11:14").do(wxstart)
    # schedule.every().day.at("11:02").do(wxstart)
    # schedule.every().day.at("11:02").do(wxstart)

    # 星期一：monday（Mon.）
    # 星期二：tuesday（Tue.）
    # 星期三：wednesday（Wed.）
    # 星期四：thursday（Thu.）
    # 星期五：friday（Fri.）
    # 星期六：saturday（Sat.）
    # 星期日：sunday（Sun.）
    # 替换day 例：schedule.every().monday.at("11:05").do(wxstart)

    # every()里的是每月的哪天 例：schedule.every(1).day.at("11:05").do(wxstart)

    # 立即执行使用:
    # schedule.every().day.at("10:15").do(run_browser_task())
    # run_browser_task()

    while True:
        schedule.run_pending()
        time.sleep(1)  # 每秒检查一次是否有任务需要执行


