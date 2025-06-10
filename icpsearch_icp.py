import requests
import yaml
import uuid
import re
import urllib3
import warnings
import pandas as pd
from urllib.parse import quote
import time
import os
import random
import argparse
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import json
import openpyxl
import hashlib
import base64
from typing import Optional, Dict, Any, List
import threading
import math
import sys

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

REFERER = "https://beian.miit.gov.cn/"
GET_TOKEN_URL = "https://hlwicpfwc.miit.gov.cn/icpproject_query/api/auth"
QUERY_URL = "https://hlwicpfwc.miit.gov.cn/icpproject_query/api/icpAbbreviateInfo/queryByCondition/"
SIGN = "eyJ0eXBlIjozLCJleHREYXRhIjp7InZhZnljb2RlX2ltYWdlX2tleSI6IjBlNzg0YzM4YmQ1ZTQwNWY4NzQyMTdiN2E5MjVjZjdhIn0sImUiOjE3MzA5NzkzNTgwMDB9.kyklc3fgv9Ex8NnlmkYuCyhe8vsLrXBcUUkEawZryGc"

PROXIES = None
TOKEN_MANAGER = None
should_continue_callback = None

from icp_config import *
from icp_token import *
from icp_network import *
from icp_query import *
from icp_utils import *

class TokenManager:
    def __init__(self):
        self.token: Optional[str] = None
        self.refresh_token: Optional[str] = None
        self.expire_in: int = 0
        self.lock = threading.Lock()
        self.http = requests.Session()
        if PROXIES:
            self.http.proxies = PROXIES
        self.http.verify = False

    def get_token(self) -> Optional[str]:
        with self.lock:
            current_time = int(time.time() * 1000)
            if self.expire_in > current_time and self.token:
                return self.token
            return self._refresh_token()

    def _refresh_token(self) -> Optional[str]:
        max_retries = 10
        retry_count = 0
        while retry_count < max_retries:
            try:
                timestamp = str(int(time.time() * 1000))
                auth_key = hashlib.md5(f"testtest{timestamp}".encode()).hexdigest()
                headers = {
                    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "Referer": REFERER,
                    "User-Agent": get_random_user_agent(),
                    "Cookie": "__jsluid_s = 6452684553c30942fcb8cff8d5aa5a5b"
                }
                data = {
                    "authKey": auth_key,
                    "timeStamp": timestamp
                }
                response = self.http.post(GET_TOKEN_URL, headers=headers, data=data)
                if response.status_code == 403:
                    retry_count += 1
                    if retry_count >= max_retries:
                        print("\n获取token失败次数过多，需要暂停...")
                        choice = input("是否继续尝试？(y/n): ").strip().lower()
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

    def get_headers(self) -> Dict[str, str]:
        token = self.get_token()
        if not token:
            raise Exception("无法获取有效的token")
        return {
            "Token": token,
            "Sign": SIGN,
            "Content-Type": "application/json",
            "User-Agent": get_random_user_agent(),
            "Referer": REFERER
        }

def create_session():
    session = requests.Session()
    session.verify = False
    retry_strategy = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504]
    )
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

def make_request_with_timeout(session, method, url, **kwargs):
    if 'timeout' not in kwargs:
        kwargs['timeout'] = 3
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            response = session.request(method, url, **kwargs)
            if response.status_code == 403 or "访问频次过高" in response.text:
                print("\n访问频次过高或403错误")
                return None
            return response
        except requests.exceptions.Timeout as e:
            retry_count += 1
            if retry_count >= max_retries:
                print("\n请求超时")
                return None
            print(f"\n请求超时，第{retry_count}次重试...")
            time.sleep(2)
            continue
        except (requests.exceptions.ProxyError, requests.exceptions.ConnectionError) as e:
            retry_count += 1
            if retry_count >= max_retries:
                print("\n连接错误")
                return None
            print(f"\n连接错误，第{retry_count}次重试...")
            time.sleep(2)
            continue
        except Exception as e:
            raise

def save_results(results, filename, show_message=True):
    try:
        existing_df = pd.DataFrame()
        if os.path.exists(filename):
            try:
                existing_df = pd.read_excel(filename)
            except:
                if show_message:
                    print(f"读取现有文件失败，将创建新文件: {filename}")
        new_df = pd.DataFrame(results)
        if not existing_df.empty:
            columns = ['企业名称', '备案域名', '备案IP', '备案微信小程序', '备案APP']
            existing_df = existing_df.reindex(columns=columns, fill_value='')
            new_df = new_df.reindex(columns=columns, fill_value='')
            for _, row in new_df.iterrows():
                company_name = row['企业名称']
                mask = existing_df['企业名称'] == company_name
                if mask.any():
                    existing_df.loc[mask] = row.values
                else:
                    existing_df = pd.concat([existing_df, pd.DataFrame([row])], ignore_index=True)
            result_df = existing_df
        else:
            result_df = new_df
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            max_widths = {}
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_length = max(max_length, length)
                adjusted_width = min(max_length * 1.2 + 2, 50)
                worksheet.column_dimensions[column].width = adjusted_width
                max_widths[column] = adjusted_width
            for cell in worksheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
                cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
            for row in worksheet.iter_rows(min_row=2):
                max_lines = 1
                max_chars_per_line = 0
                for cell in row:
                    if cell.value:
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            chars = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_chars_per_line = max(max_chars_per_line, chars)
                        column_width = max_widths[cell.column_letter]
                        lines_needed = max(1, math.ceil(max_chars_per_line / (column_width * 0.8)))
                        max_lines = max(max_lines, lines_needed)
                worksheet.row_dimensions[row[0].row].height = max_lines * 15 + 5
        if show_message:
            print(f"已保存结果到 {filename}")
        return True
    except Exception as e:
        print(f"保存结果失败: {str(e)}")
        return False

def is_valid_domain(domain):
    domain = domain.strip()
    ip_pattern = r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
    if re.match(ip_pattern, domain):
        parts = domain.split('.')
        for part in parts:
            if int(part) > 255:
                return True
        return False
    domain_pattern = r'^[a-zA-Z0-9][a-zA-Z0-9-]{0,61}[a-zA-Z0-9](\.[a-zA-Z]{2,})+$'
    if re.match(domain_pattern, domain):
        return True
    if '。' in domain or '．' in domain:
        return True
    return True

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
                print("\n访问频次过高或403错误")
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
                    sys.exit(0)
            print(f"查询备案域名失败，第{retry_count}次重试...")
            time.sleep(2)
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
                print("\n访问频次过高或403错误")
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
                    sys.exit(0)
            print(f"查询微信小程序失败，第{retry_count}次重试...")
            time.sleep(2)
            continue

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
                print("\n访问频次过高或403错误")
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
                    sys.exit(0)
            print(f"查询备案APP失败，第{retry_count}次重试...")
            time.sleep(2)
            continue

def set_should_continue_callback(cb):
    global should_continue_callback
    should_continue_callback = cb

def main():
    try:
        print(r"""
  ___          ____                      _     
 |_ _|___ _ __/ ___|  ___  __ _ _ __ ___| |__  
  | |/ __| '_ \___ \ / _ \/ _` | '__/ __| '_ \ 
  | | (__| |_) |__) |  __/ (_| | | | (__| | | |
 |___\___| .__/____/ \___|\__,_|_|  \___|_| |_|
         |_|                                   
    @https://github.com/xhmcc/icpsearch   By:心海
""")
        global TOKEN_MANAGER
        TOKEN_MANAGER = TokenManager()
        parser = argparse.ArgumentParser(
            description='企业备案信息查询工具',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog='''\n示例:\n  python icpsearch_icp.py -f input.xlsx -o output.xlsx\n  python icpsearch_icp.py -f input.xlsx -o output.xlsx -d 1  \n  python icpsearch_icp.py -f input.xlsx -o output.xlsx -proxy http://127.0.0.1:8008  \n  python icpsearch_icp.py -h  \n'''
        )
        parser.add_argument('-f', '--file', default='company_name.xlsx', help='指定输入Excel文件路径，默认为company_name.xlsx')
        parser.add_argument('-o', '--output', default='company_domains_result.xlsx', help='指定输出Excel文件路径，默认为company_domains_result.xlsx')
        parser.add_argument('-d', '--delay', type=float, default=0, help='设置请求间隔时间（秒），默认为0秒')
        parser.add_argument('-proxy', '--proxy', help='设置代理服务器')
        args = parser.parse_args()
        global PROXIES
        if args.proxy:
            PROXIES = {
                'http': args.proxy,
                'https': args.proxy
            }
            print(f"使用代理: {args.proxy}")
            try:
                session = create_session()
                response = session.get('http://httpbin.org/ip', timeout=5)
                if response.status_code != 200:
                    print("代理连接失败，将不使用代理")
                    PROXIES = None
                    return
                proxy_ip = response.json().get('origin')
                if not proxy_ip:
                    print("无法获取代理IP信息，将不使用代理")
                    PROXIES = None
                    return
            except Exception as e:
                print(f"代理验证失败: {str(e)}")
                print("将不使用代理")
                PROXIES = None
                return
        if args.delay > 0:
            print(f"请求间隔时间设置为: {args.delay}秒")
        if not os.path.exists(args.file):
            print(f"错误：输入文件 {args.file} 不存在")
            return
        df = pd.read_excel(args.file)
        if '企业名称' not in df.columns:
            print("Excel文件中没有'企业名称'列，请确保企业名称在'企业名称'列")
            return
        results = []
        for index, company_name in enumerate(df['企业名称']):
            if pd.isna(company_name):
                continue
            company_name = str(company_name).strip()
            if not company_name:
                continue
            if index > 0:
                print()
            if index > 0 and args.delay > 0:
                time.sleep(args.delay)
            print(f"正在查询企业 [{index + 1}/{len(df['企业名称'])}]: {company_name}")
            data = get_icp_domains(company_name)
            if data is None:
                print("跳过当前企业")
                continue
            domains = data['domains']
            ips = data['ips']
            miniprograms = get_miniprograms(company_name)
            if miniprograms is None:
                print("跳过微信小程序查询")
                miniprograms = []
            apps = get_apps(company_name)
            if apps is None:
                print("跳过备案APP查询")
                apps = []
            domains_str = '\n'.join(domains) if domains else '无备案域名'
            ips_str = '\n'.join(ips) if ips else '无备案IP'
            miniprograms_str = '\n'.join(miniprograms) if miniprograms else '无备案微信小程序'
            apps_str = '\n'.join(apps) if apps else '无备案APP'
            results.append({
                '企业名称': company_name,
                '备案域名': domains_str,
                '备案IP': ips_str,
                '备案微信小程序': miniprograms_str,
                '备案APP': apps_str
            })
            save_results(results, args.output, show_message=(index == len(df['企业名称']) - 1))
    except KeyboardInterrupt:
        print("\n用户中断程序")
        sys.exit(0)
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()