# 导入selenium中的webdriver库
from selenium import webdriver
# 导入chrome service库（selenium4开始需要使用service）
from selenium.webdriver.edge.service import Service as EdgeService
# 导入drivermanager用于自动获取浏览器驱动
from webdriver_manager.microsoft import EdgeChromiumDriverManager
# 导入By库来通过各种方法查找元素
from selenium.webdriver.common.by import By
# 导入ActionChains用于模拟鼠标点击
from selenium.webdriver.common.action_chains import ActionChains
# 导入WebDriverWait用于延时等待直到元素出现
from selenium.webdriver.support.ui import WebDriverWait
# 导入EC（expected_conditions）用于提供一组预定义的条件，以便在进行下一步操作之前等待某些条件的满足。
from selenium.webdriver.support import expected_conditions as EC
# 导入time来延时
import time
# 用于匹配当前时间
from datetime import date
import datetime
#导入subprocess打开企业微信
import subprocess
#导入pyautogui操作企业微信
import pyautogui
#检测代码所在目录
import os
#导入pynput输入汉字
from pynput.keyboard import Controller as KeyboardController
from pynput.keyboard import Key
#导入pyperclip读取剪切板
import pyperclip
#导入pandas操作excel
import pandas as pd
import random
#导入tkinter添加UI
import tkinter as tk
from tkinter import messagebox,filedialog,ttk
#导入sys来退出程序
import sys
#导入threading用于异步执行其他任务
import threading
#导入json用于加载配置
import json

keyboard = KeyboardController()

def get_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS # pyinstaller打包后的路径
    else:
        base_path = os.path.dirname(os.path.realpath(__file__)) # 当前工作目录的路径
    return os.path.normpath(os.path.join(base_path, relative_path)) # 返回实际路径

# 配置文件路径
config_file = get_path("config.json")

# 读取配置文件
def load_config():
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config
    else:
        # 如果配置文件不存在，返回默认值
        return {
            "account": "15327204399",
            "password": "123456789Lb",
            "wechat_path": "C:\\Program Files (x86)\\WXWork\\WXWork.exe",
            "time1": "22:50",
            "time2": "23:10",
            "remove_names": ['徐文洋', '伍珂萱']
        }

# 保存配置到文件
def save_config(config):
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

# 加载配置
config = load_config()

# 创建UI界面
def create_ui():
    global account, password, wechat_path, time1, time2, remove_names

    # 创建主窗口
    root = tk.Tk()
    root.title("参数配置")
    root.geometry("700x800")

    # 设置账号和密码输入框
    frame_account = tk.Frame(root)
    frame_account.pack(pady=5, fill="x")
    tk.Label(frame_account, text="学习通账号:").pack(side="left", padx=5)
    account_entry = tk.Entry(frame_account)
    account_entry.insert(0, config["account"])  # 默认填充账号
    account_entry.pack(side="left", fill="x", expand=True, padx=(5, 20))

    frame_password = tk.Frame(root)
    frame_password.pack(pady=5, fill="x")
    tk.Label(frame_password, text="密码:").pack(side="left", padx=(30,35))
    password_entry = tk.Entry(frame_password, show="*")
    password_entry.insert(0, config["password"])  # 默认填充密码
    password_entry.pack(side="left", fill="x", expand=True, padx=(5, 20))

    
    # 创建包含输入框和按钮的Frame（企业微信路径）
    frame_wechat = tk.Frame(root)
    frame_wechat.pack(pady=5, fill="x")
    tk.Label(frame_wechat, text="企业微信路径:").pack(side="left", padx=5)  # 左侧标签
    wechat_path_entry = tk.Entry(frame_wechat)
    wechat_path_entry.insert(0, config["wechat_path"])  # 默认填充路径
    wechat_path_entry.pack(side="left", fill="x", expand=True, padx=(5, 20))  # 填充剩余空间

    # 浏览按钮
    def browse_wechat_path():
        file_path = filedialog.askopenfilename(title="选择企业微信路径", filetypes=(("Executables", "*.exe"), ("All files", "*.*")))
        if file_path:
            wechat_path_entry.delete(0, tk.END)  # 清空输入框
            wechat_path_entry.insert(0, file_path)  # 填充选择的路径

    browse_button = tk.Button(frame_wechat, text="浏览", command=browse_wechat_path)
    browse_button.pack(side="right", padx=5)  # 按钮放到右边


    # 设置小时和分钟选择的下拉框
    def create_time_combobox(parent, time_str):
        # 分割输入时间
        hours = [str(i).zfill(2) for i in range(24)]  # 24小时制
        minutes = [str(i).zfill(2) for i in range(0, 60, 5)]  # 5分钟间隔
        time_combobox = ttk.Combobox(parent, values=[f"{h}:{m}" for h in hours for m in minutes], state="readonly")
        time_combobox.set(time_str)  # 设置默认时间
        time_combobox.pack(pady=5)
        # 滚轮事件绑定
        def on_wheel(event):
            current_time = time_combobox.get()
            current_hour, current_minute = map(int, current_time.split(":"))
            if event.delta > 0:  # 向上滚动
                current_minute += 5
                if current_minute >= 60:
                    current_minute = 0
                    current_hour += 1
                    if current_hour >= 24:
                        current_hour = 0
            elif event.delta < 0:  # 向下滚动
                current_minute -= 5
                if current_minute < 0:
                    current_minute = 55
                    current_hour -= 1
                    if current_hour < 0:
                        current_hour = 23
            # 更新 combobox 显示的时间
            time_combobox.set(f"{str(current_hour).zfill(2)}:{str(current_minute).zfill(2)}")
        # 绑定鼠标滚轮事件
        time_combobox.bind("<MouseWheel>", on_wheel)
        return time_combobox

    # 设置第一次提醒时间选择
    frame_time1 = tk.Frame(root)
    frame_time1.pack(pady=5, fill="x")
    tk.Label(frame_time1, text="第一次提醒时间:").pack(side="left", padx=5)
    time1_combobox = create_time_combobox(frame_time1, config["time1"])

    # 设置第二次提醒时间选择
    frame_time2 = tk.Frame(root)
    frame_time2.pack(pady=5, fill="x")
    tk.Label(frame_time2, text="第二次提醒时间:").pack(side="left", padx=5)
    time2_combobox = create_time_combobox(frame_time2, config["time2"])

    # 设置名单输入框
    frame_names = tk.Frame(root)
    frame_names.pack(pady=5, fill="x")
    tk.Label(frame_names, text="剔除提醒人名单（用英文逗号分隔）:").pack(side="left", padx=5)
    names_entry = tk.Entry(frame_names)
    names_entry.insert(0, ",".join(config["remove_names"]))  # 默认填充名单
    names_entry.pack(side="left", fill="x", expand=True, padx=(5, 20))

    # 保存按钮事件
    def save_settings():
        global config
        # 获取用户输入
        config["account"] = account_entry.get()
        config["password"] = password_entry.get()
        config["wechat_path"] = wechat_path_entry.get()
        config["time1"] = time1_combobox.get()
        config["time2"] = time2_combobox.get()
        config["remove_names"] = names_entry.get().split(',')
        # 保存配置到文件
        save_config(config)
        # 显示成功消息
        messagebox.showinfo("保存成功", "程序将继续运行")
        threading.Thread(target=run_program).start()  # 使用线程执行后续程序

    # 保存按钮
    save_button = tk.Button(root, text="保存设置", command=save_settings)
    save_button.pack(pady=20)
    # 绑定关闭窗口事件
    def on_closing():
        if messagebox.askokcancel("退出", "确定要退出程序吗？"):
            root.quit()  # 退出窗口事件循环
            sys.exit()
    # 绑定窗口关闭事件
    root.protocol("WM_DELETE_WINDOW", on_closing)
    # 启动UI
    root.mainloop()

def getnamelist():
    # 设置 Edge WebDriver 服务
    service = EdgeService(executable_path=EdgeChromiumDriverManager().install())
    # 初始化 Edge WebDriver
    driver = webdriver.Edge(service=service)
    # 打开指定的 URL
    driver.get('https://i.chaoxing.com/')

    # 找到id为"phone"的输入框元素
    input = driver.find_element(By.ID,'phone')
    # 输入账号
    input.send_keys(config["account"])
    # 找到id为"pwd"的输入框元素
    input = driver.find_element(By.ID,'pwd')
    # 输入密码
    input.send_keys(config["password"])
    # 寻找登录按钮
    clickable = driver.find_element(By.ID, "loginBtn")
    # 使用 ActionChains 来模拟鼠标左键点击
    ActionChains(driver).click(clickable).perform()
    time.sleep(3)
    # 切换到iframe
    driver.switch_to.frame("frame_content")
    # 定位并操作元素
    element = driver.find_element(By.ID, "myTeach")
    ActionChains(driver).click(element).perform()
    time.sleep(1)
    element = driver.find_element(By.XPATH, "//span[@title='23级晚签打卡']")
    element.click()

    # 等待新窗口或标签页出现
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
    # 存储原始窗口的句柄
    original_window = driver.current_window_handle
    # 切换到新窗口
    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)
    time.sleep(3)
    # 切换到iframe
    driver.switch_to.frame("frame_content-hd")
    # 在新页面上等待目标元素出现并操作它
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,  "//a[text()='活动列表']"))
    )
    #检测新页面的第一个要点击的元素元素
    new_page_element = driver.find_element(By.XPATH, "//a[text()='活动列表']")
    # 使用 ActionChains 来模拟鼠标左键点击
    ActionChains(driver).click(new_page_element).perform()

    # 等待新窗口或标签页出现
    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(3))
    # 存储原始窗口的句柄
    original_window = driver.current_window_handle
    # 切换到新窗口
    new_window = [window for window in driver.window_handles if window != original_window][1]
    driver.switch_to.window(new_window)
    # 在新页面上等待目标元素出现并操作它
    xpath = f'//li[contains(@class,"list-info")]//span[text()="23级晚签打卡{dt}"]/ancestor::li'
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    #检测新页面的第一个要点击的元素元素
    new_page_element = driver.find_element(By.XPATH, xpath)
    # 使用 ActionChains 来模拟鼠标左键点击
    ActionChains(driver).click(new_page_element).perform()

    # 等待页面加载完成
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "listLi")))
    # 获取所有包含用户名的<li>元素
    li_elements = driver.find_elements(By.CLASS_NAME, "listLi")
    # 获取所有包含用户名的<li>元素
    usernames = [li.get_attribute("username") for li in li_elements if li.get_attribute("signstatus") == "0"]
    for name in config["remove_names"]:
        if name in usernames:
            usernames.remove(name)
    else:
        pass 
    # 返回所有的用户名
    return (usernames)

def presssearch(text,button_center):
    # 移动鼠标到按钮位置并点击
    pyautogui.click(button_center)
    # 使用 pynput 输入文字
    for char in text:
        keyboard.press(char)
        time.sleep(0.1)
        keyboard.release(char)
    time.sleep(1)
    # 使用 pyautogui 模拟按下 Enter 键
    pyautogui.press('enter')

def getlocation():
    # 构造图像文件的完整路径
    image_path = get_path("search.jpg")
    # 尝试多次查找图像
    retries = 0
    max_retries = 10  # 最大重试次数
    button_location = None
    while retries < max_retries:
        try:
            # 使用 pyautogui 查找图像
            button_location = pyautogui.locateOnScreen(image_path, confidence=0.8)
            if button_location:
                print(f"找到搜索框的位置: {button_location}")
                break  # 找到后退出循环
            else:
                print(f"搜索框位置未找到，正在重新尝试... 第 {retries+1} 次")
        except pyautogui.ImageNotFoundException:
            print(f"未找到搜索框位置，正在重新尝试... 第 {retries+1} 次")

        retries += 1
        time.sleep(2)  # 等待2秒后再尝试
    # 如果在最大重试次数内没有找到，提示用户
    if button_location is None:
        print("搜索框位置未找到，请检查图像位置或手动调整。")
        return None
    # 获取按钮的中心位置
    button_center = pyautogui.center(button_location)
    return button_center

def inputmessage():
    for name in  names:
        message='同学你好，请记得及时晚签'
        presssearch(name,button_center)
        time.sleep(1)
        # 使用 pyautogui 输入文字
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl','v')
        emoji='/mg'
        pyperclip.copy(emoji)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(1)
        #使用 pyautogui 模拟按下 Enter 键
        pyautogui.press('enter')

def inputmessage2():
    i=0
    for name in  names2:
        presssearch(name,button_center)
        name=names[i]
        message='同学你好，麻烦提醒一下你的室友'+name+'记得及时晚签'
        time.sleep(1)
        # 使用 pyautogui 输入文字
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl','v')
        emoji='/mg'
        pyperclip.copy(emoji)
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl','v')
        time.sleep(1)
        i+=1
        #使用 pyautogui 模拟按下 Enter 键
        pyautogui.press('enter')

def find_dorm_leader(names):
    result = []
    # 获取当前脚本所在目录，并构造Excel文件路径
    file_path = get_path('B部寝室名单.xlsx')
    # 读取 Excel 文件
    df = pd.read_excel(file_path)
    for name in names:
        # 获取当前姓名所在的寝室号
        dorm_info = df[df['姓名'] == name]
        
        if dorm_info.empty:
            # 如果找不到对应的姓名，输出错误信息
            print(f"未找到姓名为 '{name}' 的数据,请再次提醒")
            continue
        dorm_number = dorm_info.iloc[0]['寝室号']
        dorm_leader = dorm_info.iloc[0]['寝室长']
        # 检查寝室长是否为空
        if pd.isna(dorm_leader):
            result[name] = '寝室长信息缺失'
            print(f"寝室长信息缺失，姓名：{name}，寝室号：{dorm_number}")
            continue
        # 找出同一寝室的人
        same_dorm_members = df[df['寝室号'] == dorm_number]['姓名'].tolist()
        # 如果这个人是寝室长，则排除自己，并从同寝室的人中随机选一个
        if name == dorm_leader:
            same_dorm_members.remove(name)  # 排除自己
            if same_dorm_members:
                selected_member = random.choice(same_dorm_members)
                result.append(selected_member)
            else:
                print(f" '{name}' 同寝室无其他成员,请再次提醒")
        else:
            result.append(dorm_leader)

    return result

def run_program():
    global dt,names,button_center,names2
    while True:
        today=date.today()
        dt=today.strftime("%Y%m%d")
        current_time = datetime.datetime.now()
        current_hour = current_time.hour
        current_minute = current_time.minute

        if f"{current_hour:02d}:{current_minute:02d}" == config["time1"]:
            #获取未晚签名单
            #names=getnamelist()
            names=['王皓翰','张宇航','郑俊韬']
            #names=['韦世泽', '刘畅', '周浩喆', '姚甜甜', '周乐涵', '祁薇晓', '蒋勤勤', '伊斯拉皮力·玉山木', '谢秉桦', '艾力库提·阿合买提江', '坚才索朗', '陈天月', '刘媛', '冯慧', '罗映雪', '李青颖', '明红艳']
            print(f'未晚签名单{names}')

            #打开企业微信
            subprocess.Popen(config["wechat_path"])
            time.sleep(5)
            button_center=getlocation()
            inputmessage()

        elif f"{current_hour:02d}:{current_minute:02d}" == config["time2"]:
            names=getnamelist()
            names2 = find_dorm_leader(names)
            print(f'需再次提醒{names2}')

            #打开企业微信
            subprocess.Popen(config["wechat_path"])
            time.sleep(5)
            button_center=getlocation()
            inputmessage2()
        time.sleep(60)
create_ui()