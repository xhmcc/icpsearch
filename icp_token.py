from icp_config import *
import time
import threading
import hashlib
from icp_network import get_random_user_agent

class TokenManager:
    def __init__(self):
        self.token = None
        self.refresh_token = None
        self.expire_in = 0
        self.lock = threading.Lock()
        self.http = None
        self.init_http()

    def init_http(self):
        import requests
        self.http = requests.Session()
        if PROXIES:
            self.http.proxies = PROXIES
        self.http.verify = False

    def get_token(self):
        with self.lock:
            current_time = int(time() * 1000)
            if self.expire_in > current_time and self.token:
                return self.token
            return self._refresh_token()

    def _refresh_token(self):
        import requests
        max_retries = 10
        retry_count = 0
        while retry_count < max_retries:
            try:
                timestamp = str(int(time() * 1000))
                auth_key = hashlib.md5(f"testtest{timestamp}".encode()).hexdigest()
                headers = {
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "Referer": "https://beian.miit.gov.cn/",
                    "User-Agent": get_random_user_agent(),
                    "Cookie": "__jsluid_s = 6452684553c30942fcb8cff8d5aa5a5b"
                }
                data = {"authKey": auth_key, "timeStamp": timestamp}
                response = self.http.post("https://hlwicpfwc.miit.gov.cn/icpproject_query/api/auth", headers=headers, data=data)
                if response.status_code == 403:
                    retry_count += 1
                    if retry_count >= max_retries:
                        print("\n获取token失败次数过多，需要暂停...")
                        if should_continue_callback:
                            choice = should_continue_callback("获取token失败次数过多，需要暂停，是否继续？")
                        else:
                            choice = 'n'
                        if choice == 'y':
                            retry_count = 0
                            continue
                        else:
                            print("用户选择停止运行")
                            return None
                    print(f"获取token失败(403)，第{retry_count}次重试...")
                    time.sleep(2)
                    continue
                if response.status_code != 200:
                    print(f"获取token失败: {response.status_code}")
                    return None
                result = response.json()
                if result.get("code") != 200:
                    print(f"获取token失败: {result.get('msg')}")
                    return None
                params = result.get("params", {})
                self.token = params.get("bussiness")
                self.refresh_token = params.get("refresh")
                self.expire_in = int(timestamp) + int(params.get("expire", 0))
                return self.token
            except Exception as e:
                print(f"刷新token时出错: {str(e)}")
                return None
    def get_headers(self):
        token = self.get_token()
        if not token:
            raise Exception("无法获取有效的token")
        return {
            "Token": token,
            "Sign": "eyJ0eXBlIjozLCJleHREYXRhIjp7InZhZnljb2RlX2ltYWdlX2tleSI6IjBlNzg0YzM4YmQ1ZTQwNWY4NzQyMTdiN2E5MjVjZjdhIn0sImUiOjE3MzA5NzkzNTgwMDB9.kyklc3fgv9Ex8NnlmkYuCyhe8vsLrXBcUUkEawZryGc",
            "Content-Type": "application/json",
            "User-Agent": get_random_user_agent(),
            "Referer": "https://beian.miit.gov.cn/"
        } 