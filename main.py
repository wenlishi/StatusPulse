import os
import re
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
#                 【核心任务函数 - 最终确认与优化版】
# ==============================================================================
def check_journal_status():
    """
    期刊状态检查任务的完整流程（最终优化版 - 三叉戟侦察）。
    能够智能识别初始页面是“主菜单”、“详情页”还是“未登录”，并采取最优策略。
    """
    if not is_operating_time():
        return

    initial_delay = random.randint(1, 15)
    print(f"\n【{time.strftime('%Y-%m-%d %H:%M:%S')}】 任务已触发，为模拟人类操作，将随机延迟 {initial_delay} 秒...")
    time.sleep(initial_delay)

    print(f"--- 开始执行期刊状态检查 ---")

    saved_data = get_saved_data()
    browser = None

    for attempt in range(MAX_RETRIES):
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=False,  # 后台运行时请改回 True
                    args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
                )
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
                    storage_state=saved_data.get('storage_state'),
                    no_viewport=True,
                )
                page = context.new_page()

                print(f"  -> 正在导航至网站入口: {target_url}...")
                page.goto(target_url, wait_until="domcontentloaded", timeout=45000)

                # --- 1. 定位并进入核心操作框架 ---
                main_app_iframe_locator = page.locator('iframe#content')
                main_app_iframe_locator.wait_for(state="attached", timeout=20000)
                main_app_frame_object = main_app_iframe_locator.content_frame
                if not main_app_frame_object:
                    raise Exception("无法获取主应用框架 (#content)。")
                print("  -> 已定位主应用框架，正在深入核心内容框架...")
                
                content_frame_locator = page.locator('iframe[name="content"]')
                content_frame_locator.wait_for(state="attached", timeout=15000)
                op_frame = content_frame_locator.content_frame
                if not op_frame: raise Exception("无法定位到核心内容操作框架 (iframe[name='content'])")
                print("  -> 已成功进入核心内容操作框架。")


                # --- 2. 智能三叉戟状态侦察 ---
                current_status = ""
                login_is_required = False
                
                try:
                    # [侦察 1/3] 是否在主菜单?
                    print("  -> [侦察 1/3] 检查是否位于主菜单 (特征: 'New Submissions')...")
                    op_frame.get_by_text("New Submissions", exact=True).wait_for(state="visible", timeout=7000)
                    print("  -> ✔️ 确认：位于主菜单页。")
                    # 无需做任何事，让代码自然流转到第4步的菜单扫描即可

                except Exception:
                    try:
                        # [侦察 2/3] 是否在详情页?
                        print("  -> [侦察 2/3] 主菜单未找到，检查是否已在详情页 (特征: 'Manuscript Number')...")
                        op_frame.get_by_text("Manuscript Number", exact=True).wait_for(state="visible", timeout=5000)
                        print("  -> ✔️ 确认：直接位于详情页。开始提取状态...")
                        first_row = op_frame.locator("table#datatable tr#row1")
                        status_cell = first_row.locator("td").nth(5)
                        detailed_status = status_cell.text_content().strip()
                        current_status = f"详情页直达 -> {detailed_status}" # 直接获取状态，任务提前完成！
                        
                    except Exception:
                        # [侦察 3/3] 既非主菜单也非详情页，则必须登录
                        print("  -> [侦察 3/3] 未找到任何已登录标志，判定需要登录。")
                        login_is_required = True

                # --- 3. 如有需要，执行登录 ---
                if login_is_required:
                    print("  -> 执行登录操作...")
                    login_iframe_locator = main_app_frame_object.locator('iframe[name="login"]')
                    login_iframe_locator.wait_for(state="attached", timeout=15000)
                    login_form_frame_object = login_iframe_locator.content_frame
                    if not login_form_frame_object: raise Exception("无法获取登录框架。")

                    type_like_human(login_form_frame_object.locator("#username"), journal_username)
                    type_like_human(login_form_frame_object.locator("#passwordTextbox"), journal_password)
                    
                    print("  -> 登录信息已输入，点击登录并等待页面跳转...")
                    with page.expect_navigation(wait_until="domcontentloaded", timeout=45000):
                        login_form_frame_object.locator('#emLoginButtonsDiv > input[type=button]:nth-child(1)').click()
                    
                    print("  -> ✔️ 登录成功！根据规则，现在必定位于主菜单。")
                    # 登录后必须重新定位核心操作框架
                    content_frame_locator = page.locator('iframe[name="content"]')
                    content_frame_locator.wait_for(state="attached", timeout=15000)
                    op_frame = content_frame_locator.content_frame
                    if not op_frame: raise Exception("登录后无法重新定位核心内容框架！")

                # --- 4. 扫描主菜单 (仅当初始状态为菜单页或刚登录时执行) ---
                if not current_status: # 如果状态还没被“详情页直达”逻辑赋值
                    print("  -> 开始在主菜单页扫描状态...")
                    candidate_items = op_frame.locator("a[cssclass='main_menu_item_2'], span[cssclass='main_menu_item_2']")
                    active_statuses = []
                    first_clickable_link = None

                    for item_locator in candidate_items.all():
                        count_span = item_locator.locator("xpath=./following-sibling::span[@class='count'][1]")
                        if count_span.is_visible():
                            count_text = count_span.inner_text()
                            match = re.search(r'\((\d+)\)', count_text)
                            if match and int(match.group(1)) > 0:
                                status_name = item_locator.text_content().strip()
                                active_statuses.append(f"{status_name} ({match.group(1)})")
                                print(item_locator.inner_html)
                                print(item_locator.inner_text)
                                if item_locator.evaluate('element => element.tagName') == 'A' and not first_clickable_link:
                                    first_clickable_link = item_locator

                    if not active_statuses:
                        current_status = "无在处理的投稿"
                    else:
                        aggregated_status = ", ".join(active_statuses)
                        if first_clickable_link:
                            link_text = first_clickable_link.text_content().strip()
                            print(f"  -> 点击首个活动链接 '{link_text}' 查看详情...")
                            with page.expect_navigation(wait_until="domcontentloaded", timeout=30000):
                                first_clickable_link.click()
                            
                            # 导航后再次确保 op_frame 是最新的
                            op_frame = page.frame(name="content")
                            if not op_frame: raise Exception("点击链接后无法重新定位核心框架！")
                            
                            first_row = op_frame.locator("table#datatable tr#row1")
                            status_cell = first_row.locator("td").nth(5)
                            detailed_status = status_cell.text_content().strip()
                            current_status = f"{aggregated_status} -> {detailed_status}"
                        else:
                            current_status = aggregated_status
                
                # --- 5. 状态比对与通知 (逻辑不变) ---
                print(f"\n  -> 本次检查获取到的最终状态是: '{current_status}'")
                latest_storage_state = context.storage_state()
                last_status = saved_data.get('last_status', '')

                print(f"  -> 上一次记录的状态是：'{last_status}'")
                if current_status and "抓取" not in current_status and current_status != last_status:
                    print(f"  -> !!! 状态发生变化, 准备发送微信通知!!!")
                    access_token = get_access_token()
                    if access_token:
                        send_status_update(access_token, last_status, current_status)
                    save_data(current_status, latest_storage_state)
                else:
                    if "抓取" in current_status: print("  -> 由于抓取状态包含错误信息，本次不更新。")
                    else:
                        print("  -> 状态无变化或无需通知，仅更新会话信息。")
                        save_data(last_status, latest_storage_state)

                print("\n--- 本次期刊状态检查任务圆满完成 ---\n")
                return # 成功，退出函数

        except (PlaywrightError, Exception) as e:
            # 错误处理和重试逻辑保持不变
            import traceback
            print(f"  -> [操作异常] 第 {attempt + 1}/{MAX_RETRIES} 次尝试失败: {type(e).__name__} - {str(e).splitlines()[0]}")
            # traceback.print_exc()
            if page and not page.is_closed():
                try:
                    screenshot_path = f"debug_error_attempt_{attempt + 1}.png"
                    page.screenshot(path=screenshot_path, full_page=True)
                    print(f"  -> [信息] 错误发生时页面截图已保存至 {screenshot_path}")
                except Exception as screenshot_e:
                    print(f"  -> [警告] 保存截图时发生错误: {screenshot_e}")
            
            if attempt < MAX_RETRIES - 1:
                sleep_time = (attempt + 1) * 20
                print(f"  -> 将在 {sleep_time} 秒后重试...")
                time.sleep(sleep_time)
            else:
                print("  -> 已达到最大重试次数，任务彻底中断。")
        # finally:
        #     if browser and browser.is_connected():
        #         browser.close()


def is_operating_time(time_start: int = 4, time_end: int = 6):
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
    schedule.every().hour.at(":22").do(check_journal_status)
    
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