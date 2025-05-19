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
import socks
import socket
import sys

# 禁用 SSL 警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 常量定义
REFERER = "https://beian.miit.gov.cn/"
GET_TOKEN_URL = "https://hlwicpfwc.miit.gov.cn/icpproject_query/api/auth"
QUERY_URL = "https://hlwicpfwc.miit.gov.cn/icpproject_query/api/icpAbbreviateInfo/queryByCondition/"
SIGN = "eyJ0eXBlIjozLCJleHREYXRhIjp7InZhZnljb2RlX2ltYWdlX2tleSI6IjBlNzg0YzM4YmQ1ZTQwNWY4NzQyMTdiN2E5MjVjZjdhIn0sImUiOjE3MzA5NzkzNTgwMDB9.kyklc3fgv9Ex8NnlmkYuCyhe8vsLrXBcUUkEawZryGc"

# 全局变量
PROXIES = None
AQC_COOKIE = None
TOKEN_MANAGER = None

# 读取配置
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)

AQC_COOKIE = config.get('aiqicha_cookie')

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
        """获取token，如果过期则自动刷新"""
        with self.lock:
            current_time = int(time.time() * 1000)
            if self.expire_in > current_time and self.token:
                return self.token
            
            return self._refresh_token()

    def _refresh_token(self) -> Optional[str]:
        """刷新token"""
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
                    time.sleep(2)  # 等待2秒后重试
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
        """获取带有token的请求头"""
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
    """创建请求会话"""
    session = requests.Session()
    session.verify = False
    
    # 设置重试策略
    retry_strategy = Retry(
        total=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504]
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    
    # 设置代理
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
    """验证代理是否可用"""
    try:
        # 验证代理URL格式
        if not proxy_url.startswith(('http://', 'https://', 'socks5://', 'socks5h://')):
            proxy_url = 'http://' + proxy_url

        # 创建测试会话
        session = requests.Session()
        session.verify = False
        session.proxies = {
            'http': proxy_url,
            'https': proxy_url
        }
        
        # 测试代理连接
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
    """从文件加载代理列表"""
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
    """使用超时检测的请求函数"""
    if 'timeout' not in kwargs:
        kwargs['timeout'] = 3  # 默认3秒超时
    
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            response = session.request(method, url, **kwargs)
            # 检查响应状态码和内容
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
            time.sleep(2)  # 等待2秒后重试
            continue
        except (requests.exceptions.ProxyError, requests.exceptions.ConnectionError) as e:
            retry_count += 1
            if retry_count >= max_retries:
                print("\n连接错误，切换代理...")
                return None
            print(f"\n连接错误，第{retry_count}次重试...")
            time.sleep(2)  # 等待2秒后重试
            continue
        except Exception as e:
            # 其他错误不切换代理
            raise

def save_results(results, filename, show_message=True):
    try:
        # 如果文件存在，先尝试读取现有数据
        existing_df = pd.DataFrame()
        if os.path.exists(filename):
            try:
                existing_df = pd.read_excel(filename)
            except:
                if show_message:
                    print(f"读取现有文件失败，将创建新文件: {filename}")
        
        # 将新结果转换为DataFrame
        new_df = pd.DataFrame(results)
        
        if not existing_df.empty:
            # 确保列的顺序一致
            columns = ['企业名称', '备案域名', '备案IP', '备案微信小程序', '备案微信公众号', '备案APP']
            existing_df = existing_df.reindex(columns=columns, fill_value='')
            new_df = new_df.reindex(columns=columns, fill_value='')
            
            # 如果存在现有数据，更新相同企业名的记录
            for _, row in new_df.iterrows():
                company_name = row['企业名称']
                # 更新或添加记录
                mask = existing_df['企业名称'] == company_name
                if mask.any():
                    existing_df.loc[mask] = row.values
                else:
                    existing_df = pd.concat([existing_df, pd.DataFrame([row])], ignore_index=True)
            
            result_df = existing_df
        else:
            result_df = new_df
        
        # 保存所有结果
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
            # 获取工作表
            worksheet = writer.sheets['Sheet1']
            
            # 计算每列的最大字符宽度
            max_widths = {}
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # 获取列字母
                
                # 遍历列中的所有单元格
                for cell in col:
                    if cell.value:
                        # 计算每个单元格中每行的最大长度
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            # 中文字符按2个字符宽度计算
                            length = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_length = max(max_length, length)
                
                # 设置列宽（字符数 * 1.2 + 2，确保有适当边距）
                adjusted_width = min(max_length * 1.2 + 2, 50)  # 限制最大宽度为50
                worksheet.column_dimensions[column].width = adjusted_width
                max_widths[column] = adjusted_width
            
            # 设置标题行格式
            for cell in worksheet[1]:
                cell.font = openpyxl.styles.Font(bold=True)
                cell.alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # 设置数据行格式并自动调整行高
            for row in worksheet.iter_rows(min_row=2):  # 从第二行开始（跳过标题行）
                max_lines = 1
                max_chars_per_line = 0
                
                for cell in row:
                    if cell.value:
                        # 设置单元格格式
                        cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
                        
                        # 计算每行的字符数
                        lines = str(cell.value).split('\n')
                        for line in lines:
                            # 中文字符按2个字符宽度计算
                            chars = sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in line)
                            max_chars_per_line = max(max_chars_per_line, chars)
                        
                        # 计算需要的行数
                        column_width = max_widths[cell.column_letter]
                        lines_needed = max(1, math.ceil(max_chars_per_line / (column_width * 0.8)))  # 0.8是考虑到字符间距
                        max_lines = max(max_lines, lines_needed)
                
                # 设置行高（每行15像素，额外加5像素的padding）
                worksheet.row_dimensions[row[0].row].height = max_lines * 15 + 5
        
        if show_message:
            print(f"已保存结果到 {filename}")
        return True
    except Exception as e:
        print(f"保存结果失败: {str(e)}")
        return False

def update_config_cookie(new_cookie):
    """更新config.yaml文件中的cookie"""
    try:
        # 读取现有配置
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        # 更新cookie
        config['aiqicha_cookie'] = new_cookie
        
        # 写回文件
        with open('config.yaml', 'w', encoding='utf-8') as f:
            yaml.dump(config, f, allow_unicode=True)
        
        print("已更新config.yaml文件中的cookie")
        return True
    except Exception as e:
        print(f"更新config.yaml失败: {str(e)}")
        return False

def get_company_id_aiqicha(company_name, is_first_query=False):
    """获取爱企查企业ID"""
    global AQC_COOKIE
    encoded_name = quote(company_name)
    url = f"https://aiqicha.baidu.com/s?q={encoded_name}&t=0"
    
    while True:
        try:
            headers = {
                "User-Agent": get_random_user_agent(),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                "Accept-Encoding": "gzip, deflate",
                "Connection": "close",
                "Referer": "https://aiqicha.baidu.com/",
                "Cookie": AQC_COOKIE,
                "Sec-Ch-Ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
                "Sec-Ch-Ua-Mobile": "?0",
                "Sec-Ch-Ua-Platform": '"Windows"',
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "same-origin",
                "Sec-Fetch-User": "?1",
                "Upgrade-Insecure-Requests": "1",
                "Cache-Control": "max-age=0"
            }
            
            session = create_session()
            response = session.get(url, headers=headers, verify=False, timeout=15, allow_redirects=False)
            
            if response.status_code == 302:
                if is_first_query:
                    print("\n爱企查Cookie无效，需要重新设置...")
                    choice = input("是否现在设置Cookie？(y/n): ").strip().lower()
                    if choice == 'y':
                        new_cookie = input("请输入Cookie: ").strip()
                        if new_cookie:
                            update_config_cookie(new_cookie)
                            AQC_COOKIE = new_cookie
                            continue
                        else:
                            print("未输入Cookie，爱企查查询不可用，跳过微信公众号查询")
                            return None
                    else:
                        print("未设置Cookie，爱企查查询不可用，跳过微信公众号查询")
                        return None
                else:
                    return None
            
            if response.status_code == 200:
                try:
                    # 查找包含JSON数据的script标签
                    json_match = re.search(r'window\.pageData\s*=\s*({.*?});', response.text, re.DOTALL)
                    if json_match:
                        json_str = json_match.group(1)
                        data = json.loads(json_str)
                        
                        if 'result' in data and 'resultList' in data['result']:
                            # 遍历结果列表
                            for item in data['result']['resultList']:
                                if 'entName' in item:
                                    # 移除HTML标签后比较企业名称
                                    clean_name = item['entName'].replace('<em>', '').replace('</em>', '')
                                    
                                    # 如果企业名称完全匹配，返回对应的pid
                                    if clean_name == company_name:
                                        company_id = item['pid']
                                        return company_id
                            
                            # 如果没有完全匹配的，返回第一个结果的pid
                            if data['result']['resultList']:
                                return data['result']['resultList'][0]['pid']
                    
                    print(f"未找到匹配的企业ID: {company_name}")
                    return None
                except Exception as e:
                    print(f"解析响应数据失败")
                    return None
            elif response.status_code == 429:
                print("\n请求过于频繁，需要暂停...")
                choice = input("是否继续运行？(y/n): ").strip().lower()
                if choice == 'y':
                    print("继续运行，重试上一个请求...")
                    continue
                else:
                    print("用户选择停止运行")
                    sys.exit(0)
            else:
                return None
        except Exception as e:
            print(f"请求出错")
            return None

def is_valid_domain(domain):
    """检查是否为有效的域名"""
    # 清理域名中的空格
    domain = domain.strip()
    
    # 检查是否是IP地址格式
    ip_pattern = r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$'
    if re.match(ip_pattern, domain):
        # 验证IP地址的每个数字是否在有效范围内
        parts = domain.split('.')
        for part in parts:
            if int(part) > 255:
                return True  # 如果数字超过255，可能是域名
        return False  # 是有效的IP地址
    
    # 检查是否包含常见的域名特征
    domain_pattern = r'^[a-zA-Z0-9][a-zA-Z0-9-]{0,61}[a-zA-Z0-9](\.[a-zA-Z]{2,})+$'
    if re.match(domain_pattern, domain):
        return True
    
    # 检查是否包含中文域名特征
    if '。' in domain or '．' in domain or '．' in domain or '．' in domain:
        return True
    
    # 如果都不匹配，默认认为是域名
    return True

def get_icp_domains(company_name: str) -> dict:
    """查询企业的备案域名和IP信息"""
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
            
            if response is None:  # 如果是访问频次过高或403错误
                print("\n访问频次过高或403错误，切换代理...")
                return None
            
            if response.status_code != 200:
                print(f"查询备案域名失败: {response.status_code}")
                return {'domains': [], 'ips': []}
                
            result = response.json()
            if result.get("code") != 200:
                print(f"查询备案域名失败: {result.get('msg')}")
                return {'domains': [], 'ips': []}
                
            # 解析结果
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
                choice = input("是否继续尝试？(y/n): ").strip().lower()
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    sys.exit(0)  # 直接退出程序
            print(f"查询备案域名失败，第{retry_count}次重试...")
            time.sleep(2)  # 等待2秒后重试
            continue

def get_miniprograms(company_name: str) -> list:
    """查询企业的微信小程序信息"""
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
            
            if response is None:  # 如果是访问频次过高或403错误
                print("\n访问频次过高或403错误，切换代理...")
                return None
            
            if response.status_code != 200:
                print(f"查询微信小程序失败: {response.status_code}")
                return []
                
            result = response.json()
            if result.get("code") != 200:
                print(f"查询微信小程序失败: {result.get('msg')}")
                return []
                
            # 解析结果
            items = []
            params = result.get("params", {})
            for item in params.get("list", []):
                service_name = item.get("serviceName", "")
                if service_name:
                    items.append(service_name)
            
            # 输出找到的微信小程序信息
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
                choice = input("是否继续尝试？(y/n): ").strip().lower()
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    sys.exit(0)  # 直接退出程序
            print(f"查询微信小程序失败，第{retry_count}次重试...")
            time.sleep(2)  # 等待2秒后重试
            continue

def get_wechat_accounts(company_id: str) -> list:
    """查询企业的微信公众号信息"""
    global AQC_COOKIE
    url = f"https://aiqicha.baidu.com/c/wechatoaAjax?pid={company_id}"
    
    try:
        headers = {
            "User-Agent": get_random_user_agent(),
            "Accept": "text/html, application/xhtml+xml, image/jxr, */*",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "close",
            "Referer": "https://aiqicha.baidu.com/",
            "Cookie": AQC_COOKIE
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
                
            # 使用集合去重
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
    """查询企业的备案APP信息"""
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
            
            if response is None:  # 如果是访问频次过高或403错误
                print("\n访问频次过高或403错误，切换代理...")
                return None
            
            if response.status_code != 200:
                print(f"查询备案APP失败: {response.status_code}")
                return []
                
            result = response.json()
            if result.get("code") != 200:
                print(f"查询备案APP失败: {result.get('msg')}")
                return []
                
            # 解析结果
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
                choice = input("是否继续尝试？(y/n): ").strip().lower()
                if choice == 'y':
                    retry_count = 0
                    continue
                else:
                    print("用户选择停止运行")
                    sys.exit(0)  # 直接退出程序
            print(f"查询备案APP失败，第{retry_count}次重试...")
            time.sleep(2)  # 等待2秒后重试
            continue

def check_aiqicha_cookie():
    """检查爱企查cookie是否有效"""
    try:
        # 读取配置
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        cookie = config.get('aiqicha_cookie')
        if not cookie:
            print("\n爱企查Cookie未配置，需要设置...")
            choice = input("是否现在设置Cookie？(y/n): ").strip().lower()
            if choice == 'y':
                new_cookie = input("请输入Cookie: ").strip()
                if new_cookie:
                    update_config_cookie(new_cookie)
                    return new_cookie
                else:
                    print("未输入Cookie，爱企查查询不可用，跳过微信公众号查询")
                    return None
            else:
                print("未设置Cookie，爱企查查询不可用，跳过微信公众号查询")
                return None
        
        # 测试cookie是否有效
        headers = {
            "User-Agent": get_random_user_agent(),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
            "Accept-Encoding": "gzip, deflate",
            "Connection": "close",
            "Referer": "https://aiqicha.baidu.com/",
            "Cookie": cookie
        }
        
        session = create_session()
        response = session.get("https://aiqicha.baidu.com/", headers=headers, verify=False, timeout=15, allow_redirects=False)
        
        if response.status_code == 302:
            if 'login' in response.headers.get('Location', '').lower():
                print("\n爱企查Cookie已过期，需要更新...")
                choice = input("是否现在更新Cookie？(y/n): ").strip().lower()
                if choice == 'y':
                    new_cookie = input("请输入新的Cookie: ").strip()
                    if new_cookie:
                        update_config_cookie(new_cookie)
                        return new_cookie
                    else:
                        print("未输入新的Cookie，爱企查查询不可用，跳过微信公众号查询")
                        return None
                else:
                    print("未更新Cookie，爱企查查询不可用，跳过微信公众号查询")
                    return None
            elif PROXIES:
                print("\n当前代理无法使用爱企查，请更换代理或关闭代理后重试")
                return None
        
        return cookie
        
    except Exception as e:
        print(f"检查Cookie时出错: {str(e)}")
        return None

def verify_proxy_ip():
    """验证代理IP是否为大陆IP"""
    try:
        session = create_session()
        # 尝试多个IP查询服务
        urls = [
            'https://whois.pconline.com.cn/ipJson.jsp?ip=myip&json=true',
            'https://ip.taobao.com/outGetIpInfo?ip=myip&accessKey=alibaba-inc',
            'https://ip.chinaz.com/getip.aspx',
            'https://ip.3322.net'
        ]
        
        # 非大陆地区列表
        non_mainland = ['香港', '台湾', '澳门']
        
        for url in urls:
            try:
                response = session.get(url, timeout=5)
                if response.status_code == 200:
                    data = response.json() if 'json' in url else {'ip': response.text.strip()}
                    
                    # 根据不同API的返回格式处理
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
                        # 3322.net 只返回IP，需要额外查询
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
                        # 检查是否是非大陆地区
                        if any(region in country for region in non_mainland):
                            return False
                        return True
            except:
                continue
                
        return False
    except Exception as e:
        return False

def verify_and_switch_proxy(proxy_list, current_index, proxy_file=None):
    """验证并切换代理，找到第一个可用的大陆IP就使用"""
    if not proxy_list:
        return None, current_index
        
    # 从当前索引开始尝试
    start_index = current_index
    tried_proxies = set()  # 记录已尝试过的代理
    
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
        
        # 设置代理
        global PROXIES
        PROXIES = {
            'http': proxy,
            'https': proxy
        }
        
        # 验证代理是否生效
        try:
            session = create_session()
            response = session.get('http://httpbin.org/ip', timeout=5)
            if response.status_code == 200:
                proxy_ip = response.json().get('origin', '')
                if proxy_ip:
                    # 检查是否为大陆IP
                    if verify_proxy_ip():
                        # 获取IP地理位置
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

def main():
    try:
        # 显示logo和版本信息
        print(r"""
  ___          ____                      _     
 |_ _|___ _ __/ ___|  ___  __ _ _ __ ___| |__  
  | |/ __| '_ \___ \ / _ \/ _` | '__/ __| '_ \ 
  | | (__| |_) |__) |  __/ (_| | | | (__| | | |
 |___\___| .__/____/ \___|\__,_|_|  \___|_| |_|
         |_|                                   
    @https://github.com/xhmcc/icpsearch   By:心海
""")
        
        # 初始化TokenManager
        global TOKEN_MANAGER
        TOKEN_MANAGER = TokenManager()
        
        # 解析命令行参数
        parser = argparse.ArgumentParser(
            description='企业备案信息查询工具',
            formatter_class=argparse.RawDescriptionHelpFormatter,
            epilog='''
示例:
  python icpsearch_icp.py -f input.xlsx -o output.xlsx
  python icpsearch_icp.py -f input.xlsx -o output.xlsx -d 1  # 设置请求间隔为1秒
  python icpsearch_icp.py -f input.xlsx -o output.xlsx -proxy http://127.0.0.1:8008  # 使用代理
  python icpsearch_icp.py -f input.xlsx -o output.xlsx -proxy proxypool.txt  # 使用代理池
  python icpsearch_icp.py -h  # 显示帮助信息
'''
        )
        parser.add_argument('-f', '--file', default='company_name.xlsx', help='指定输入Excel文件路径，默认为company_name.xlsx')
        parser.add_argument('-o', '--output', default='company_domains_result.xlsx', help='指定输出Excel文件路径，默认为company_domains_result.xlsx')
        parser.add_argument('-d', '--delay', type=float, default=0, help='设置请求间隔时间（秒），默认为0秒')
        parser.add_argument('-proxy', '--proxy', help='设置代理服务器或代理池文件路径')
        args = parser.parse_args()

        # 设置代理
        global PROXIES
        proxy_list = []
        current_proxy_index = -1  # 从-1开始，这样第一次会尝试第一个代理
        skip_aiqicha = False
        proxy_switch_counter = 0  # 代理切换计数器
        successful_queries = 0  # 成功查询计数器

        if args.proxy:
            if os.path.isfile(args.proxy):
                # 如果是文件，加载代理列表
                proxy_list = load_proxy_list(args.proxy)
                if not proxy_list:
                    print("代理列表为空，将不使用代理")
                else:
                    print(f"已加载 {len(proxy_list)} 个代理")
                    # 验证并找到第一个可用的大陆IP代理
                    PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                    if PROXIES is None:
                        print("所有代理都不可用，将不使用代理")
                    else:
                        print(f"使用代理 [{current_proxy_index + 1}/{len(proxy_list)}]: {proxy_list[current_proxy_index]}")
            else:
                # 如果是单个代理，直接使用，不检查是否为大陆IP
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

        # 如果指定了延迟，显示延迟信息
        if args.delay > 0:
            print(f"请求间隔时间设置为: {args.delay}秒")

        # 检查输入文件是否存在
        if not os.path.exists(args.file):
            print(f"错误：输入文件 {args.file} 不存在")
            return

        # 读取Excel文件
        df = pd.read_excel(args.file)
        
        # 确保企业名称列存在
        if '企业名称' not in df.columns:
            print("Excel文件中没有'企业名称'列，请确保企业名称在'企业名称'列")
            return
            
        # 创建结果列表
        results = []
        
        # 遍历每个企业名称
        for index, company_name in enumerate(df['企业名称']):
            if pd.isna(company_name):  # 跳过空值
                continue
                
            company_name = str(company_name).strip()
            if not company_name:  # 跳过空字符串
                continue
            
            # 如果不是第一个企业，打印一个空行
            if index > 0:
                print()
            
            # 如果不是第一个企业且设置了延迟，则等待指定时间
            if index > 0 and args.delay > 0:
                time.sleep(args.delay)
            
            print(f"正在查询企业 [{index + 1}/{len(df['企业名称'])}]: {company_name}")
            
            # 如果使用代理池且已经成功查询了5个企业，切换代理
            if proxy_list and successful_queries >= 3:
                print(f"\n已成功查询{successful_queries}个企业，切换代理...")
                PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                if PROXIES is None:
                    print("所有代理都不可用，将不使用代理")
                else:
                    print(f"使用代理 [{current_proxy_index + 1}/{len(proxy_list)}]: {proxy_list[current_proxy_index]}")
                successful_queries = 0
            
            # 获取企业ID（用于微信公众号查询）
            company_id = None
            query_success = False  # 标记本次查询是否成功
            
            if not skip_aiqicha:
                try:
                    company_id = get_company_id_aiqicha(company_name, is_first_query=(index == 0))
                    if company_id is not None:
                        query_success = True
                    elif proxy_list:
                        # 只有在不是Cookie问题的情况下才切换代理
                        while current_proxy_index < len(proxy_list):
                            PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                            if PROXIES is None:
                                print("所有代理都已尝试，跳过微信公众号查询")
                                skip_aiqicha = True
                                break
                            if not verify_proxy_ip():
                                continue
                            try:
                                company_id = get_company_id_aiqicha(company_name, is_first_query=False)
                                if company_id is not None:
                                    query_success = True
                                    break
                            except Exception as e:
                                print(f"获取企业ID失败: {str(e)}")
                                continue
                except Exception as e:
                    print(f"获取企业ID失败: {str(e)}")
            
            # 获取备案域名和IP
            data = get_icp_domains(company_name)
            if data is None:  # 如果返回None，说明需要切换代理
                if proxy_list:
                    PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                    if PROXIES is None:
                        print("所有代理都不可用，跳过当前企业")
                        continue
                    data = get_icp_domains(company_name)
                    if data is None:
                        print("切换代理后仍然失败，跳过当前企业")
                        continue
                else:
                    print("未使用代理池，跳过当前企业")
                    continue
            
            domains = data['domains']
            ips = data['ips']
            if domains or ips:
                query_success = True
            
            # 获取微信小程序信息
            miniprograms = get_miniprograms(company_name)
            if miniprograms is None:  # 如果返回None，说明需要切换代理
                if proxy_list:
                    PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                    if PROXIES is None:
                        print("所有代理都不可用，跳过微信小程序查询")
                        miniprograms = []
                    else:
                        miniprograms = get_miniprograms(company_name)
                        if miniprograms is None:
                            print("切换代理后仍然失败，跳过微信小程序查询")
                            miniprograms = []
                else:
                    print("未使用代理池，跳过微信小程序查询")
                    miniprograms = []
            
            if miniprograms:
                query_success = True
            
            # 获取微信公众号信息（使用爱企查）
            wechat_accounts = []
            if not skip_aiqicha and company_id:
                try:
                    wechat_accounts = get_wechat_accounts(company_id)
                    if wechat_accounts:
                        query_success = True
                except Exception as e:
                    print(f"获取微信公众号信息失败: {str(e)}")
            
            # 获取备案APP信息
            apps = get_apps(company_name)
            if apps is None:  # 如果返回None，说明需要切换代理
                if proxy_list:
                    PROXIES, current_proxy_index = verify_and_switch_proxy(proxy_list, current_proxy_index, args.proxy)
                    if PROXIES is None:
                        print("所有代理都不可用，跳过备案APP查询")
                        apps = []
                    else:
                        apps = get_apps(company_name)
                        if apps is None:
                            print("切换代理后仍然失败，跳过备案APP查询")
                            apps = []
                else:
                    print("未使用代理池，跳过备案APP查询")
                    apps = []
            
            if apps:
                query_success = True
            
            # 如果本次查询成功，增加成功计数器
            if query_success:
                successful_queries += 1
                
                # 格式化输出，使用换行符分隔
                domains_str = '\n'.join(domains) if domains else '无备案域名'
                ips_str = '\n'.join(ips) if ips else '无备案IP'
                miniprograms_str = '\n'.join(miniprograms) if miniprograms else '无备案微信小程序'
                wechat_accounts_str = '\n'.join(wechat_accounts) if wechat_accounts else '无备案微信公众号'
                apps_str = '\n'.join(apps) if apps else '无备案APP'
                
                results.append({
                    '企业名称': company_name,
                    '备案域名': domains_str,
                    '备案IP': ips_str,
                    '备案微信小程序': miniprograms_str,
                    '备案微信公众号': wechat_accounts_str,
                    '备案APP': apps_str
                })
                
                # 每处理完一个企业就保存一次结果，但只在最后一个企业时显示提示
                save_results(results, args.output, show_message=(index == len(df['企业名称']) - 1))
            
    except KeyboardInterrupt:
        print("\n用户中断程序")
        sys.exit(0)
    except Exception as e:
        print(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()