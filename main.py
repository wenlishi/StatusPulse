# 模拟真人版
import os
import time
import requests
import random
import datetime
import threading
import json
import schedule
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError, Error as PlaywrightError
from dotenv import load_dotenv


load_dotenv()

# 从测试号信息获取
appID = os.getenv("APPID")
appSecret = os.getenv("APPSECRET")
#收信人ID即 用户列表中的微信号，见上文
openId = os.getenv("OPENID")
# 审稿状态变更通知模板ID
status_template_id = os.getenv("STATUS_TEMPLATE_ID")

journal_username = os.getenv("JOURNAL_USERNAME")
journal_password = os.getenv("JOURNAL_PASSWORD")
target_url = os.getenv("TARGET_URL")


STATUS_FILE = Path("journal_data.json")
MAX_RETRIES = 3


def type_like_human(locator, text_to_type):
    """模拟真人逐字输入，并带有随机间隔"""
    locator.hover()
    time.sleep(random.uniform(0.5, 1.0))
    for char in text_to_type:
        locator.press(char)
        time.sleep(random.uniform(0.08, 0.25))
 
def get_saved_data():
    """从本地JSON文件读取上次的状态和Cookies"""
    if not STATUS_FILE.exists():
        return {"last_status": "首次运行", "storage_state": None}
    try:
        with open(STATUS_FILE, 'r') as f:
            data = json.load(f)
            return {
                "last_status": data.get("last_status", "首次运行"),
                "storage_state": data.get("storage_state")
            }
    except (json.JSONDecodeError, IOError):
        return {"last_status": "文件读取错误", "storage_state": None}
 
def save_data(status, storage_state):
    """将最新状态和Cookies保存到本地JSON文件"""
    with open(STATUS_FILE, 'w') as f:
        json.dump({"last_status": status, "storage_state": storage_state}, f, indent=4)
    print(f"  -> 最新状态 '{status}' 和会话Cookies已保存至 {STATUS_FILE}")
 
def get_access_token():
    # 获取access token的url
    url = 'https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={}&secret={}' \
        .format(appID.strip(), appSecret.strip())
    response = requests.get(url).json()
    print(response)
    access_token = response.get('access_token')
    return access_token

 
def send_status_update(access_token, last_status, current_status):
    """发送企业微信应用消息"""
    check_time_str = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M")
    url = f"https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={access_token}"

    payload = {
        "touser": openId,
        "msgtype": "text",
        "template_id": status_template_id,
      

        "data": {
            "title": {"value": "您的期刊投稿状态有新的变化！"}, # 红色标题
            "check_time": {"value": f" {check_time_str}"},
            "old_status": {"value": f" {last_status }"},
            "new_status": {"value": f" {current_status}"}
        }
    }
    try:
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()
        result = response.json()
        if result.get("errcode") == 0:
            print("  -> 企业微信通知发送成功！")
        else:
            print(f"  -> [错误] 企业微信通知发送失败: {result.get('errmsg')}")
    except requests.RequestException as e:
        print(f"  -> [错误] 发送企业微信通知时网络异常: {e}")
 
 
# ==============================================================================
#                        【核心任务函数 - 最终优化版】
# ==============================================================================
 
def check_journal_status():
    """期刊状态检查任务的完整流程"""
    ## 检查是否是允许操作时间段
    if not is_operating_time():
        return 

    initial_delay = random.randint(1, 30)
    print(f"\n【{time.strftime('%Y-%m-%d %H:%M:%S')}】 任务已触发，为模拟人类操作，将随机延迟 {initial_delay} 秒...")
    time.sleep(initial_delay)
    
    print(f"--- 开始执行期刊状态检查 ---")
    
    current_status = ""
    saved_data = get_saved_data()
 
    for attempt in range(MAX_RETRIES):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
                )
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
                    storage_state=saved_data.get('storage_state'), # 使用 .get 更安全
                    no_viewport=True,
                )
                page = context.new_page()
 
                print("  -> 正在导航至网站入口...")
                page.goto(f"{target_url}", wait_until="domcontentloaded")
 
                main_frame = page.frame_locator('#content')
                
                # 【核心逻辑】优先检查是否已直接处于最终的状态页面
                try:
                    # 'Current Status' 是最终状态页的标志性文本
                    main_frame.get_by_text("Current Status", exact=True).wait_for(state="visible", timeout=10000)
                    print("  -> 检测到有效会话，并已直接进入状态列表页面。")
                
                except TimeoutError:
                    # 如果10秒内没看到 "Current Status"，说明需要登录
                    print("  -> Session已过期或未登录，执行登录操作。")
                    login_frame = main_frame.frame_locator('iframe[name="login"]')
                    print(login_frame)
                    type_like_human(login_frame.locator("#username"), journal_username)
                    type_like_human(login_frame.locator("#passwordTextbox"), journal_password)
                    
                    login_button = login_frame.locator('#emLoginButtonsDiv > input[type=button]:nth-child(1)')
                    login_button.click()
                    
                    # +++ 修改部分 +++
                    print("  -> 登录按钮已点击，正在等待登录后页面加载...")
                    # 我们不等待网络空闲，而是明确等待登录后应该出现的 "Current Status" 文本
                    # 这证明登录成功且页面已跳转

                    main_frame = page.frame_locator('iframe[name="content"]')
                    main_frame.get_by_text("Submissions Being Processed").click()
                    page.wait_for_load_state("load", timeout=30000)




                    # main_frame.get_by_text("Current Status", exact=True).wait_for(state="visible", timeout=45000)
                    # print("  -> 登录成功，页面已加载。")
 
                # --- 抓取状态 ---
                print("  -> 正在抓取页面状态...")
                try:
                    first_row = main_frame.locator("table#datatable tr#row1")
                    first_row.wait_for(state="visible", timeout=20000)
                    
                    # 根据截图，“With Editor”在第二列。td的索引从0开始，所以用nth(1)
                    status_cell = first_row.locator("td").nth(5) 
                    
                    current_status = status_cell.text_content().strip()
                    print(f"  -> 成功抓取到当前状态：'{current_status}'")
                except TimeoutError:
                    current_status = "无在处理的投稿"
                    print("  -> 未发现正在处理的投稿。")
                except Exception as e:
                    current_status = f"抓取页面元素时出错"
                    print(f"  -> [错误] 定位状态元素时失败: {e}")
 
                # 保存最新的Cookies和Session
                latest_storage_state = context.storage_state()
                browser.close()
 
                # --- 任务成功，比较状态并通知 ---
                last_status = saved_data['last_status']
                print(f"  -> 上一次记录的状态是：'{last_status}'")
 
                if current_status and current_status != last_status:
                    print(f"  -> !!! 状态发生变化 ({last_status} -> {current_status})，准备发送微信通知 !!!")
                    access_token = get_access_token()
                    if access_token:
                        send_status_update(access_token, last_status, current_status)
                        save_data(current_status, latest_storage_state) # 只有通知成功后才保存新状态
                    else:
                        print("  -> 获取access_token失败，本次状态将不会被保存，等待下次重试。")
                else:
                    print("  -> 状态无变化，无需通知。")
                    save_data(last_status, latest_storage_state) # 状态没变，也要更新cookies保持会话活跃
 
                print("--- 本次期刊状态检查任务完成 ---\n")
                return # 成功执行后，退出函数
 
        except (PlaywrightError, TimeoutError) as e:
            print(f"  -> [严重错误] 第 {attempt + 1}/{MAX_RETRIES} 次尝试失败: {type(e).__name__}")
            if attempt < MAX_RETRIES - 1:
                sleep_time = (attempt + 1) * 60
                print(f"  -> 将在 {sleep_time} 秒后重试...")
                time.sleep(sleep_time)
            else:
                print("  -> 已达到最大重试次数，任务中断。等待下一个调度周期。")
                return

def is_operating_time(time_start: int = 1, time_end: int = 6):
    """检查当前时间是否是凌晨1点到6点,该时间段内不对网站进行爬取"""
    current_time_obj = datetime.datetime.now()
    current_time = current_time_obj.time() # 获取当前时间的时间部分
    span_start = datetime.time(time_start,0)
    span_end = datetime.time(time_end,0)
    is_blocked = False
    # 如果起始时间小于结束时间
    if span_start < span_end:
        if span_start < current_time < span_end:
            is_blocked = True
    else:
        # 如果起始时间大于结束时间(跨午夜，如22:00-5:00)
        if current_time > span_start or current_time < span_end:
            is_blocked = True
    if is_blocked:
        print(f"当前时间{current_time.strftime("%H:%M:%S")}处于模拟睡觉时间段"
              f"{span_start.strftime("%H:%M")}-{span_end.strftime("%H:%M")},任务跳过")
        return False
    return True



if __name__ == '__main__':
    # 任务2: 每小时的第5分钟，检查一次期刊状态 (例如 1:05, 2:05, 3:05...)
    # 使用 at(":05") 可以避免在整点执行，稍微错开高峰
    print("脚本启动成功！服务已初始化。")
    print(f"任务 'check_journal_status' 每小時执行一次")
    schedule.every().hour.at(":23").do(check_journal_status)
    #schedule.every(45).minutes.do(check_journal_status)
    last_heartbeat_time = time.time()

    while True:
       schedule.run_pending()# 每隔600秒打印一次心跳日志
       current_time = time.time()
       if current_time - last_heartbeat_time > 600:
            print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 服务心跳：正常运行中，等待下一个任务...")
            last_heartbeat_time = current_time
                   
       time.sleep(1)


