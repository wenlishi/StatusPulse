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
        print(f"  -> [信息] 数据文件 {STATUS_FILE} 不存在，将以首次运行模式启动。")
        return {"last_status": "首次运行", "storage_state": None}
    try:
        with open(STATUS_FILE, 'r') as f:
            data = json.load(f)
            return {
                "last_status": data.get("last_status", "首次运行"),
                "storage_state": data.get("storage_state")
            }
    except (json.JSONDecodeError, IOError) as e:
        print(f"  -> [警告] 读取 {STATUS_FILE} 失败，可能文件损坏或格式不正确：{e}。将以首次运行模式启动。")
        return {"last_status": "文件读取错误", "storage_state": None}
 
def save_data(status, storage_state):
    """将最新状态和Cookies保存到本地JSON文件"""
    try:
        with open(STATUS_FILE, 'w') as f:
            json.dump({"last_status": status, "storage_state": storage_state}, f, indent=4)
        print(f"  -> 最新状态 '{status}' 和会话Cookies已保存至 {STATUS_FILE}")
    except IOError as e:
        print(f"  -> [错误] 保存数据到 {STATUS_FILE} 失败: {e}")
 
def get_access_token():
    """获取微信 access token"""
    url = f'https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={appID.strip()}&secret={appSecret.strip()}'
    try:
        response = requests.get(url, timeout=10).json()
        access_token = response.get('access_token')
        if not access_token:
            print(f"  -> [错误] 从微信API获取access_token失败: {response.get('errmsg', '未知错误')}")
        return access_token
    except requests.RequestException as e:
        print(f"  -> [错误] 获取access_token时网络异常: {e}")
        return None
 
def send_status_update(access_token, last_status, current_status):
    """发送微信模板消息"""
    check_time_str = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M")
    url = f"https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={access_token}"

    payload = {
        "touser": openId,
        "template_id": status_template_id,
        "data": {
            "title": {"value": "您的期刊投稿状态有新的变化！", "color": "#FF0000"},
            "check_time": {"value": f" {check_time_str}"},
            "old_status": {"value": f" {last_status}"},
            "new_status": {"value": f" {current_status}"}
        }
    }
    try:
        response = requests.post(url, json=payload, timeout=15)
        response.raise_for_status()
        result = response.json()
        if result.get("errcode") == 0:
            print("  -> 微信通知发送成功！")
        else:
            print(f"  -> [错误] 微信通知发送失败: {result.get('errmsg')}, 详细: {result}")
    except requests.RequestException as e:
        print(f"  -> [错误] 发送微信通知时网络异常: {e}")
 
# ==============================================================================
#                        【核心任务函数 - 最终修正版】
# ==============================================================================
 
def check_journal_status():
    """期刊状态检查任务的完整流程"""
    if not is_operating_time():
        return

    initial_delay = random.randint(1, 15) # 稍微缩短延迟
    print(f"\n【{time.strftime('%Y-%m-%d %H:%M:%S')}】 任务已触发，为模拟人类操作，将随机延迟 {initial_delay} 秒...")
    time.sleep(initial_delay)
    
    print(f"--- 开始执行期刊状态检查 ---")
    
    saved_data = get_saved_data()
    
    browser = None
    page = None

    for attempt in range(MAX_RETRIES):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=True,
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
                )
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36", # 更新UA
                    storage_state=saved_data.get('storage_state'), 
                    no_viewport=True,
                )
                page = context.new_page()

                print(f"  -> 正在导航至网站入口: {target_url}...")
                page.goto(f"{target_url}", wait_until="domcontentloaded")

                main_app_iframe_locator = page.locator('iframe#content')
                main_app_iframe_locator.wait_for(state="attached", timeout=20000)
                # *** 修正点: `content_frame` 是属性，不是方法 ***
                main_app_frame_object = main_app_iframe_locator.content_frame
                if not main_app_frame_object:
                    raise Exception("无法获取主应用程序框架 (#content) 的内容帧。")
                print("  -> 已成功定位到主应用程序框架的内容帧。")

                current_content_frame_object = None

                try:
                    # 检测是否已登录
                    logged_in_content_iframe_locator = main_app_frame_object.locator('iframe[name="content"]')
                    logged_in_content_iframe_locator.wait_for(state="attached", timeout=10000)
                    # *** 修正点 ***
                    current_content_frame_object = logged_in_content_iframe_locator.content_frame
                    if not current_content_frame_object:
                        raise Exception("无法获取登录后内容框架 (iframe[name=\"content\"]) 的内容帧。")
                    
                    current_content_frame_object.get_by_text("Submissions Being Processed", exact=True).wait_for(state="visible", timeout=10000)
                    print("  -> 检测到有效会话，页面已显示投稿处理信息。")

                except TimeoutError:
                    print("  -> Session已过期或未登录，执行登录操作。")
                    
                    login_iframe_locator = main_app_frame_object.locator('iframe[name="login"]')
                    login_iframe_locator.wait_for(state="attached", timeout=15000)
                    # *** 修正点 ***
                    login_form_frame_object = login_iframe_locator.content_frame
                    if not login_form_frame_object:
                        raise Exception("无法获取登录框架 (iframe[name=\"login\"]) 的内容帧。")
                    print("  -> 已定位到登录iframe的内容框架。")

                    username_input = login_form_frame_object.locator("#username")
                    username_input.wait_for(state="visible", timeout=15000)
                    type_like_human(username_input, journal_username)
                    print("  -> 输入好了用户名。")

                    password_input = login_form_frame_object.locator("#passwordTextbox")
                    password_input.wait_for(state="visible", timeout=15000)
                    type_like_human(password_input, journal_password)
                    print("  -> 输入好了密码。")

                    login_button = login_form_frame_object.locator('#emLoginButtonsDiv > input[type=button]:nth-child(1)')
                    login_button.click()
                    print("  -> 登录按钮已点击，正在等待登录后页面加载...")

                    logged_in_content_iframe_locator = main_app_frame_object.locator('iframe[name="content"]')
                    logged_in_content_iframe_locator.wait_for(state="attached", timeout=45000)
                    # *** 修正点 ***
                    current_content_frame_object = logged_in_content_iframe_locator.content_frame
                    if not current_content_frame_object:
                         raise Exception("登录后无法获取内容框架的内容帧。")
                    
                    current_content_frame_object.get_by_text("Submissions Being Processed").wait_for(state="visible", timeout=15000)
                    print("  -> 登录成功，且已检测到'Submissions Being Processed'链接。")

                if not current_content_frame_object:
                    raise Exception("无法获取到有效的当前内容框架，任务中断。")

                print("  -> 尝试点击 'Submissions Being Processed' 查看列表...")
                submissions_link = current_content_frame_object.get_by_text("Submissions Being Processed")
                submissions_link.click()
                # 等待点击操作触发的导航完成
                current_content_frame_object.page.wait_for_load_state("load", timeout=30000)
                print("  -> 'Submissions Being Processed' 链接已点击，列表页面加载完成。")

                print("  -> 正在抓取页面状态...")
                try:
                    first_row = current_content_frame_object.locator("table#datatable tr#row1")
                    first_row.wait_for(state="visible", timeout=20000)
                    # 请再次确认状态列的索引是否为5（第六列）
                    status_cell = first_row.locator("td").nth(5)
                    current_status = status_cell.text_content().strip()
                    print(f"  -> 成功抓取到当前状态：'{current_status}'")
                except TimeoutError:
                    current_status = "无在处理的投稿"
                    print("  -> 未发现正在处理的投稿，或加载超时。")
                except Exception as e:
                    current_status = f"抓取页面元素时出错: {type(e).__name__} - {e}"
                    print(f"  -> [错误] 定位状态元素时失败: {e}")
 
                latest_storage_state = context.storage_state()
 
                last_status = saved_data['last_status']
                print(f"  -> 上一次记录的状态是：'{last_status}'")
 
                if current_status and "抓取页面元素时出错" not in current_status and current_status != last_status:
                    print(f"  -> !!! 状态发生变化 ({last_status} -> {current_status})，准备发送微信通知 !!!")
                    access_token = get_access_token()
                    if access_token:
                        send_status_update(access_token, last_status, current_status)
                        save_data(current_status, latest_storage_state)
                    else:
                        print("  -> 获取access_token失败，本次状态将不会被保存，等待下次重试。")
                else:
                    if "抓取页面元素时出错" in current_status:
                         print("  -> 由于抓取状态出错，本次不更新状态。")
                    else:
                        print("  -> 状态无变化或为空，无需通知。")
                        save_data(last_status, latest_storage_state)
 
                print("--- 本次期刊状态检查任务完成 ---\n")
                return # 成功，退出函数
 
        except (PlaywrightError, Exception) as e: # 捕获更广泛的错误
            print(f"  -> [严重错误] 第 {attempt + 1}/{MAX_RETRIES} 次尝试失败: {type(e).__name__} - {e}")
            if page and not page.is_closed():
                try:
                    screenshot_path = f"debug_error_attempt_{attempt + 1}.png"
                    page.screenshot(path=screenshot_path)
                    print(f"  -> [信息] 错误发生时页面截图已保存至 {screenshot_path}")
                except Exception as screenshot_e:
                    print(f"  -> [警告] 保存截图时发生错误: {screenshot_e}")
            
            if attempt < MAX_RETRIES - 1:
                sleep_time = (attempt + 1) * 30 # 缩短重试时间
                print(f"  -> 将在 {sleep_time} 秒后重试...")
                time.sleep(sleep_time)
            else:
                print("  -> 已达到最大重试次数，任务中断。等待下一个调度周期。")
        finally:
            # 只有在with块成功执行完毕或发生错误跳出循环后，Playwright才会关闭
            # 这个finally确保浏览器在每次重试后都被关闭
            if browser:
                try:
                    browser.close()
                except Exception as e:
                    # 这个警告是预期的，因为如果主任务出错，Playwright实例可能已关闭
                    print(f"  -> [信息] 关闭浏览器时发生错误（这在任务失败时可能是正常现象）: {e}")


def is_operating_time(time_start: int = 2, time_end: int = 6):
    """检查当前时间是否处于非爬取时间段（凌晨2点到6点）。"""
    current_time_obj = datetime.datetime.now()
    current_time = current_time_obj.time()
    span_start = datetime.time(time_start,0)
    span_end = datetime.time(time_end,0)
    is_blocked = False

    if span_start < span_end:
        if span_start <= current_time < span_end:
            is_blocked = True
    else:
        if current_time >= span_start or current_time < span_end:
            is_blocked = True

    if is_blocked:
        print(f"当前时间{current_time.strftime('%H:%M:%S')}处于模拟睡觉时间段"
              f"{span_start.strftime('%H:%M')}-{span_end.strftime('%H:%M')},任务跳过")
        return False
    return True


if __name__ == '__main__':
    print("脚本启动成功！服务已初始化。")
    print(f"任务 'check_journal_status' 每小时的第40分钟执行一次。")
    schedule.every().hour.at(":54").do(check_journal_status)
    
    # 立即执行一次用于测试
    # check_journal_status()

    last_heartbeat_time = time.time()

    while True:
       schedule.run_pending()
       current_time_loop = time.time()
       if current_time_loop - last_heartbeat_time > 600:
            print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 服务心跳：正常运行中，等待下一个任务...")
            last_heartbeat_time = current_time_loop
                   
       time.sleep(1)
