import requests
import random
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
from icp_config import PROXIES

def create_session():
    session = requests.Session()
    session.verify = False
    retry_strategy = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    if PROXIES:
        session.proxies = PROXIES
    return session

def get_random_user_agent():
    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
    ]
    return random.choice(user_agents)

def verify_proxy(proxy_url):
    try:
        if not proxy_url.startswith(('http://', 'https://', 'socks5://', 'socks5h://')):
            proxy_url = 'http://' + proxy_url
        session = requests.Session()
        session.verify = False
        session.proxies = {'http': proxy_url, 'https': proxy_url}
        response = session.get('http://httpbin.org/ip', timeout=5)
        if response.status_code == 200:
            proxy_ip = response.json().get('origin')
            print(f"代理IP: {proxy_ip}")
            return True
        return False
    except Exception as e:
        print(f"代理验证失败: {str(e)}")
        return False

def load_proxy_list(proxy_file):
    try:
        with open(proxy_file, 'r', encoding='utf-8') as f:
            proxies = []
            for line in f:
                line = line.strip()
                if line and ':' in line:
                    proxies.append(line)
        return proxies
    except Exception as e:
        print(f"加载代理列表失败: {str(e)}")
        return []

def make_request_with_timeout(session, method, url, **kwargs):
    if 'timeout' not in kwargs:
        kwargs['timeout'] = 3
    max_retries = 3
    retry_count = 0
    import time
    while retry_count < max_retries:
        try:
            response = session.request(method, url, **kwargs)
            if response.status_code == 403 or "访问频次过高" in response.text:
                print("\n访问频次过高或403错误，切换代理...")
                return None
            return response
        except requests.exceptions.Timeout as e:
            retry_count += 1
            if retry_count >= max_retries:
                print("\n请求超时，切换代理...")
                return None
            print(f"\n请求超时，第{retry_count}次重试...")
            time.sleep(2)
            continue
        except (requests.exceptions.ProxyError, requests.exceptions.ConnectionError) as e:
            retry_count += 1
            if retry_count >= max_retries:
                print("\n连接错误，切换代理...")
                return None
            print(f"\n连接错误，第{retry_count}次重试...")
            time.sleep(2)
            continue
        except Exception as e:
            raise 