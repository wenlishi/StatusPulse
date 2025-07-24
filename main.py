# 安装依赖 pip3 install requests html5lib bs4 schedule
import os
import time
import requests
import datetime
import json
import schedule
from playwright.sync_api import sync_playwright
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

STATUS_FILE = "last_status.txt"



def get_access_token():
    # 获取access token的url
    url = 'https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={}&secret={}' \
        .format(appID.strip(), appSecret.strip())
    response = requests.get(url).json()
    print(response)
    access_token = response.get('access_token')
    return access_token


# ===================================================================
# 【新增】期刊状态检查功能
# ===================================================================
def get_last_status():
    """从文件中读取上一次保存的投稿的状态"""
    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            return f.read().strip()
    except FileNotFoundError:
        return ""


def save_current_status(status):
    """将新状态写入文件"""
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        f.write(status)

def send_status_update(access_token, old_status, new_status):
    check_time_str = datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M")
    body = {
        "touser": openId,
        "template_id": status_template_id,
        "url": "https://www.editorialmanager.com/bae/default2.aspx", # 点击消息跳转的链接
        "data": {
            "title": {"value": "您的期刊投稿状态有新的变化！", "color": "#FF0000"}, # 红色标题
            "check_time": {"value": check_time_str},
            "old_status": {"value": old_status or "（无记录）"},
            "new_status": {"value": new_status, "color": "#0000FF"} # 蓝色新状态
        }
    }
    url = f'https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={access_token}'
    print(requests.post(url, data=json.dumps(body)).text)

def check_journal_status():
    """【新增】期刊状态检查任务的完整流程"""
    print(f"\n【{time.strftime('%Y-%m-%d %H:%M:%S')}】 开始执行期刊状态检查任务...")
    current_status = ""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36")
            page = context.new_page()
 
            # 登录和导航
            page.goto("https://www.editorialmanager.com/bae/default2.aspx", timeout=60000)
            login_frame = page.frame_locator('iframe[name="content"]').frame_locator('iframe[name="login"]')
            login_frame.locator("#username").fill(journal_username)
            login_frame.locator("#passwordTextbox").fill(journal_password)
            login_frame.locator('#emLoginButtonsDiv > input[type=button]:nth-child(1)').click()
            page.wait_for_load_state("load", timeout=30000)
            
            main_frame = page.frame_locator('iframe[name="content"]')
            main_frame.get_by_text("Submissions Being Processed").click()
            page.wait_for_load_state("load", timeout=30000)
 
            # 抓取状态
            try:
                first_row = main_frame.locator("table#datatable tr#row1")
                first_row.wait_for(state="visible", timeout=15000)
                status_cell = first_row.locator("td:nth-child(6)")
                current_status = status_cell.text_content().strip()
                print(f"  -> 成功抓取到当前状态：'{current_status}'")
            except TimeoutError:
                current_status = "无在处理的投稿"
                print("  -> 未发现正在处理的投稿。")
            
            browser.close()
 
    except Exception as e:
        print(f"  -> [错误] Playwright执行失败: {e}")
        return # 出错则直接中断任务
 
    # 状态比较与通知
    last_status = get_last_status()
    print(f"  -> 上一次记录的状态是：'{last_status}'")
 
    if current_status and current_status != last_status:
        print("  -> !!! 状态发生变化，准备发送微信通知 !!!")
        access_token = get_access_token()
        if access_token:
            send_status_update(access_token, last_status, current_status)
            save_current_status(current_status)
            print("  -> 状态更新通知已发送，并已更新本地记录。")
        else:
            print("  -> 获取access_token失败，无法发送通知。")
    else:
        print("  -> 状态无变化，无需通知。")
    
    print("  -> 期刊状态检查任务完成。")


if __name__ == '__main__':
    # 任务2: 每小时的第5分钟，检查一次期刊状态 (例如 1:05, 2:05, 3:05...)
    # 使用 at(":05") 可以避免在整点执行，稍微错开高峰
    print("脚本启动成功！服务已初始化。")
    print(f"任务 'check_journal_status' 每20分钟执行一次")
    schedule.every(20).minutes.do(check_journal_status)
    last_heartbeat_time = time.time()

    while True:
       schedule.run_pending()# 每隔600秒打印一次心跳日志
       current_time = time.time()
       if current_time - last_heartbeat_time > 600:
            print(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 服务心跳：正常运行中，等待下一个任务...")
            last_heartbeat_time = current_time
                   
       time.sleep(1)
