from icp_config import *
from icp_token import TokenManager
from icp_network import *
from icp_utils import *
import sys
import time
import re
import json
import yaml
import os

def set_should_continue_callback(cb):
    global should_continue_callback
    should_continue_callback = cb

def get_icp_domains(company_name: str) -> dict:
    global TOKEN_MANAGER
    if TOKEN_MANAGER is None:
        TOKEN_MANAGER = TokenManager()
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            data = {
                "pageNum": "1",
                "pageSize": "40",
                "serviceType": "1",
                "unitName": company_name
            }
            headers = TOKEN_MANAGER.get_headers()
            session = create_session()
            response = make_request_with_timeout(session, 'POST', QUERY_URL, headers=headers, json=data, verify=False)
            if response is None:
                print("\n访问频次过高或403错误，切换代理...")
                return None
            if response.status_code != 200:
                print(f"查询备案域名失败: {response.status_code}")
                return {'domains': [], 'ips': []}
            result = response.json()
            if result.get("code") != 200:
                print(f"查询备案域名失败: {result.get('msg')}")
                return {'domains': [], 'ips': []}
            domains = set()
            ips = set()
            params = result.get("params", {})
            for item in params.get("list", []):
                domain = item.get("domain", "")
                if domain:
                    if is_valid_domain(domain):
                        domains.add(domain)
                    else:
                        ips.add(domain)
            if domains:
                print(f"找到域名: {', '.join(domains)}")
            else:
                print("未找到备案域名")
            if ips:
                print(f"找到IP: {', '.join(ips)}")
            else:
                print("未找到备案IP")
            return {
                'domains': list(domains),
                'ips': list(ips)
            }
        except Exception as e:
            print(f"查询备案域名时出错: {str(e)}")
            retry_count += 1
            if retry_count >= max_retries:
                print("\n查询备案域名失败次数过多，需要暂停...")
                if should_continue_callback:
                    choice = should_continue_callback("查询备案域名失败次数过多，需要暂停，是否继续？")
                else:
                    choice = 'n'
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    exit(0)
            print(f"查询备案域名失败，第{retry_count}次重试...")
            sleep(2)
            continue

def get_miniprograms(company_name: str) -> list:
    global TOKEN_MANAGER
    if TOKEN_MANAGER is None:
        TOKEN_MANAGER = TokenManager()
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            data = {
                "pageNum": 1,
                "pageSize": 10,
                "unitName": company_name,
                "serviceType": "7"
            }
            headers = TOKEN_MANAGER.get_headers()
            session = create_session()
            response = make_request_with_timeout(session, 'POST', QUERY_URL, headers=headers, json=data, verify=False)
            if response is None:
                print("\n访问频次过高或403错误，切换代理...")
                return None
            if response.status_code != 200:
                print(f"查询微信小程序失败: {response.status_code}")
                return []
            result = response.json()
            if result.get("code") != 200:
                print(f"查询微信小程序失败: {result.get('msg')}")
                return []
            items = []
            params = result.get("params", {})
            for item in params.get("list", []):
                service_name = item.get("serviceName", "")
                if service_name:
                    items.append(service_name)
            if items:
                print(f"找到微信小程序: {', '.join(items)}")
            else:
                print("未找到备案微信小程序")
            return items
        except Exception as e:
            print(f"查询微信小程序时出错: {str(e)}")
            retry_count += 1
            if retry_count >= max_retries:
                print("\n查询微信小程序失败次数过多，需要暂停...")
                if should_continue_callback:
                    choice = should_continue_callback("查询微信小程序失败次数过多，需要暂停，是否继续？")
                else:
                    choice = 'n'
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    exit(0)
            print(f"查询微信小程序失败，第{retry_count}次重试...")
            sleep(2)
            continue

def get_wechat_accounts(company_id: str) -> list:
    url = f"https://aiqicha.baidu.com/c/wechatoaAjax?pid={company_id}"
    try:
        headers = {
            "User-Agent": get_random_user_agent(),
            "Accept": "text/html, application/xhtml+xml, image/jxr, */*",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "close",
            "Referer": "https://aiqicha.baidu.com/"
        }
        session = create_session()
        response = session.get(url, headers=headers, verify=False, allow_redirects=False)
        if response.status_code != 200:
            print(f"查询微信公众号失败: {response.status_code}")
            return []
        try:
            data = response.json()
            if 'data' not in data or 'list' not in data['data']:
                print("未找到微信公众号信息")
                return []
            accounts = set()
            for item in data['data']['list']:
                if 'wechatName' in item and item['wechatName']:
                    accounts.add(item['wechatName'])
            if accounts:
                print(f"找到微信公众号: {', '.join(accounts)}")
            else:
                print("未找到备案微信公众号")
            return list(accounts)
        except Exception as e:
            print(f"解析微信公众号数据失败: {str(e)}")
            return []
    except Exception as e:
        print(f"查询微信公众号时出错: {str(e)}")
        return []

def get_apps(company_name: str) -> list:
    global TOKEN_MANAGER
    if TOKEN_MANAGER is None:
        TOKEN_MANAGER = TokenManager()
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            data = {
                "pageNum": "1",
                "pageSize": "40",
                "serviceType": "6",
                "unitName": company_name
            }
            headers = TOKEN_MANAGER.get_headers()
            session = create_session()
            response = make_request_with_timeout(session, 'POST', QUERY_URL, headers=headers, json=data, verify=False)
            if response is None:
                print("\n访问频次过高或403错误，切换代理...")
                return None
            if response.status_code != 200:
                print(f"查询备案APP失败: {response.status_code}")
                return []
            result = response.json()
            if result.get("code") != 200:
                print(f"查询备案APP失败: {result.get('msg')}")
                return []
            apps = set()
            params = result.get("params", {})
            for item in params.get("list", []):
                app_name = item.get("serviceName", "")
                if app_name:
                    apps.add(app_name)
            if apps:
                print(f"找到备案APP: {', '.join(apps)}")
            else:
                print("未找到备案APP")
            return list(apps)
        except Exception as e:
            print(f"查询备案APP时出错: {str(e)}")
            retry_count += 1
            if retry_count >= max_retries:
                print("\n查询备案APP失败次数过多，需要暂停...")
                if should_continue_callback:
                    choice = should_continue_callback("查询备案APP失败次数过多，需要暂停，是否继续？")
                else:
                    choice = 'n'
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    exit(0)
            print(f"查询备案APP失败，第{retry_count}次重试...")
            sleep(2)
            continue

def verify_proxy_ip():
    try:
        session = create_session()
        urls = [
            'https://whois.pconline.com.cn/ipJson.jsp?ip=myip&json=true',
            'https://ip.taobao.com/outGetIpInfo?ip=myip&accessKey=alibaba-inc',
            'https://ip.chinaz.com/getip.aspx',
            'https://ip.3322.net'
        ]
        non_mainland = ['香港', '台湾', '澳门']
        for url in urls:
            try:
                response = session.get(url, timeout=5)
                if response.status_code == 200:
                    data = response.json() if 'json' in url else {'ip': response.text.strip()}
                    if 'pconline.com.cn' in url:
                        country = data.get('pro', '')
                        ip = data.get('ip', '')
                    elif 'taobao.com' in url:
                        country = data.get('data', {}).get('country', '')
                        ip = data.get('data', {}).get('ip', '')
                    elif 'chinaz.com' in url:
                        country = data.get('location', '')
                        ip = data.get('ip', '')
                    elif '3322.net' in url:
                        ip = data.get('ip', '')
                        try:
                            location_response = session.get(f'https://whois.pconline.com.cn/ipJson.jsp?ip={ip}&json=true', timeout=5)
                            if location_response.status_code == 200:
                                location_data = location_response.json()
                                country = location_data.get('pro', '')
                        except:
                            country = ''
                    else:
                        continue
                    if country and ip:
                        if any(region in country for region in non_mainland):
                            return False
                        return True
            except:
                continue
        return False
    except Exception as e:
        return False

def verify_and_switch_proxy(proxy_list, current_index, proxy_file=None):
    if not proxy_list:
        return None, current_index
    start_index = current_index
    tried_proxies = set()
    while True:
        current_index = (current_index + 1) % len(proxy_list)
        if current_index == start_index or len(tried_proxies) >= len(proxy_list):
            print("\n所有代理都不可用，将不使用代理")
            return None, current_index
        proxy = proxy_list[current_index]
        if proxy in tried_proxies:
            continue
        tried_proxies.add(proxy)
        print(f"\r正在测试代理: {proxy}", end="", flush=True)
        global PROXIES
        PROXIES = {
            'http': proxy,
            'https': proxy
        }
        try:
            session = create_session()
            response = session.get('http://httpbin.org/ip', timeout=5)
            if response.status_code == 200:
                proxy_ip = response.json().get('origin', '')
                if proxy_ip:
                    if verify_proxy_ip():
                        try:
                            location_response = session.get(
                                f'https://whois.pconline.com.cn/ipJson.jsp?ip={proxy_ip}&json=true',
                                timeout=5,
                                verify=False
                            )
                            if location_response.status_code == 200:
                                location_data = location_response.json()
                                country = location_data.get('pro', '')
                                city = location_data.get('city', '')
                                print(f"\n找到可用的大陆IP代理: {proxy}")
                                print(f"出口IP: {proxy_ip} ({country} {city})")
                            else:
                                print(f"\n找到可用的大陆IP代理: {proxy}")
                                print(f"出口IP: {proxy_ip}")
                        except:
                            print(f"\n找到可用的大陆IP代理: {proxy}")
                            print(f"出口IP: {proxy_ip}")
                        return PROXIES, current_index
                    else:
                        continue
        except Exception as e:
            continue

if __name__ == "__main__":
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    exit(app.exec_())