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
from typing import Optional, Dict, Any
import threading
import math

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

# 读取配置
with open('config.yaml', 'r', encoding='utf-8') as f:
    config = yaml.safe_load(f)

AQC_COOKIE = config.get('aiqicha_cookie')

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

def create_session():
    """创建session，不设置自动重试"""
    session = requests.Session()
    return session

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
    global AQC_COOKIE
    # 使用具体的搜索URL，对中文进行URL编码
    encoded_name = quote(company_name)
    url = f"https://aiqicha.baidu.com/s?q={encoded_name}&t=0"
    
    while True:  # 添加循环以支持重试
        try:
            print(f"正在查询企业: {company_name}")
            
            # 每次请求都生成新的headers
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
            response = session.get(url, headers=headers, verify=False, timeout=15, proxies=PROXIES, allow_redirects=False)
            
            if response.status_code == 200:
                # 从HTML响应中提取JSON数据
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
                                    print(f"企业名称: {clean_name}")
                                    
                                    # 如果企业名称完全匹配，返回对应的pid
                                    if clean_name == company_name:
                                        company_id = item['pid']
                                        return company_id
                    
                    print(f"未找到匹配的企业ID: {company_name}")
                    return None
                except Exception as e:
                    print(f"解析响应数据失败: {str(e)}")
                    print("响应内容：", response.text[:200])  # 打印部分响应内容以便调试
                    return None
            elif response.status_code == 429:
                print("\n请求过于频繁，需要暂停...")
                choice = input("是否继续运行？(y/n): ").strip().lower()
                if choice == 'y':
                    print("继续运行，重试上一个请求...")
                    continue
                else:
                    print("用户选择停止运行")
                    return None
            elif response.status_code == 302:
                print("\nCookie已过期，需要更新...")
                choice = input("是否继续运行？(y/n): ").strip().lower()
                if choice == 'y':
                    new_cookie = input("请输入新的Cookie: ").strip()
                    if new_cookie:
                        AQC_COOKIE = new_cookie
                        headers["Cookie"] = new_cookie
                        # 更新config.yaml文件
                        update_config_cookie(new_cookie)
                        print("Cookie已更新，重试上一个请求...")
                        continue
                    else:
                        print("未输入新的Cookie，停止运行")
                        return None
                else:
                    print("用户选择停止运行")
                    return None
            else:
                print(f"请求失败，状态码: {response.status_code}")
                print("响应内容：", response.text[:200])  # 打印部分响应内容以便调试
                return None
        except Exception as e:
            print(f"请求出错: {str(e)}")
            return None

def is_valid_domain(domain):
    """检查是否为有效的域名"""
    # 简单的域名验证：不包含IP地址格式
    if re.match(r'^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$', domain):
        return False
    return True

def get_icp_domains_aiqicha(company_id):
    global AQC_COOKIE
    url = f"https://aiqicha.baidu.com/detail/intellectualPropertyAjax?pid={company_id}"
    
    while True:  # 添加循环以支持重试
        try:
            # 每次请求都生成新的headers
            headers = {
                "User-Agent": get_random_user_agent(),
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                "Accept-Encoding": "gzip, deflate",
                "Connection": "close",
                "Referer": f"https://aiqicha.baidu.com/company_detail_{company_id}?tab=certRecord",
                "Cookie": AQC_COOKIE,
                "Sec-Ch-Ua": '"Chromium";v="136", "Microsoft Edge";v="136", "Not.A/Brand";v="99"',
                "Sec-Ch-Ua-Mobile": "?0",
                "Sec-Ch-Ua-Platform": '"Windows"',
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin",
                "X-Requested-With": "XMLHttpRequest",
                "Zx-Open-Url": f"https://aiqicha.baidu.com/company_detail_{company_id}"
            }
            
            session = create_session()
            response = session.get(url, headers=headers, verify=False, proxies=PROXIES, allow_redirects=False)
            
            if response.status_code == 200:
                data = response.json()
                try:
                    # 使用集合去重
                    domains = set()
                    ips = set()
                    
                    # 检查data字段是否存在
                    if 'data' in data and 'icpinfo' in data['data'] and 'list' in data['data']['icpinfo']:
                        items = data['data']['icpinfo']['list']
                        for item in items:
                            # 提取域名
                            if 'domain' in item and isinstance(item['domain'], list):
                                for domain in item['domain']:
                                    if domain:
                                        if is_valid_domain(domain):
                                            domains.add(domain)
                                        else:
                                            ips.add(domain)
                            
                            # 提取homeSite中的域名
                            if 'homeSite' in item and isinstance(item['homeSite'], list):
                                for site in item['homeSite']:
                                    if site:
                                        # 移除www.前缀
                                        site = site.replace('www.', '')
                                        if is_valid_domain(site):
                                            domains.add(site)
                                        else:
                                            ips.add(site)
                    
                    print(f"找到域名: {', '.join(domains)}")  # 打印找到的域名
                    print(f"找到IP: {', '.join(ips) if ips else '无备案IP'}") 
                    return {
                        'domains': list(domains),
                        'ips': list(ips)
                    }
                except Exception as e:
                    print("解析域名和IP失败", e)
                    print("响应数据:", data)  # 打印完整的响应数据以便调试
                    return {'domains': [], 'ips': []}
            elif response.status_code == 429:
                print("\n请求过于频繁，需要暂停...")
                choice = input("是否继续运行？(y/n): ").strip().lower()
                if choice == 'y':
                    print("继续运行，重试上一个请求...")
                    continue
                else:
                    print("用户选择停止运行")
                    return {'domains': [], 'ips': []}
            elif response.status_code == 302:
                print("\nCookie已过期，需要更新...")
                choice = input("是否继续运行？(y/n): ").strip().lower()
                if choice == 'y':
                    new_cookie = input("请输入新的Cookie: ").strip()
                    if new_cookie:
                        AQC_COOKIE = new_cookie
                        headers["Cookie"] = new_cookie
                        # 更新config.yaml文件
                        update_config_cookie(new_cookie)
                        print("Cookie已更新，重试上一个请求...")
                        continue
                    else:
                        print("未输入新的Cookie，停止运行")
                        return {'domains': [], 'ips': []}
                else:
                    print("用户选择停止运行")
                    return {'domains': [], 'ips': []}
            else:
                print(f"备案域名请求失败 {response.status_code}")
                print("响应内容：", response.text[:200])  # 打印部分响应内容以便调试
                return {'domains': [], 'ips': []}
        except Exception as e:
            print(f"请求出错: {str(e)}")
            return {'domains': [], 'ips': []}

def get_miniprograms(company_name: str) -> list:
    """查询企业的微信小程序信息"""
    global TOKEN_MANAGER
    
    if TOKEN_MANAGER is None:
        TOKEN_MANAGER = TokenManager()
    
    try:
        # 准备请求数据
        data = {
            "pageNum": 1,
            "pageSize": 10,
            "unitName": company_name,
            "serviceType": "7"  # 7表示微信小程序
        }
        
        # 获取带有token的请求头
        headers = TOKEN_MANAGER.get_headers()
        
        # 发送请求
        response = requests.post(
            QUERY_URL,
            headers=headers,
            json=data,
            verify=False,
            proxies=PROXIES
        )
        
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
        return []

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
        response = session.get(url, headers=headers, verify=False, proxies=PROXIES, allow_redirects=False)
        
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
                    # 直接使用原始名称，因为响应已经是UTF-8编码
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

def get_apps(company_id: str) -> list:
    """查询企业的备案APP信息"""
    global AQC_COOKIE
    url = f"https://aiqicha.baidu.com/c/appinfoAjax?pid={company_id}&p=1"
    
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
        response = session.get(url, headers=headers, verify=False, proxies=PROXIES, allow_redirects=False)
        
        if response.status_code != 200:
            print(f"查询备案APP失败: {response.status_code}")
            return []
            
        try:
            data = response.json()
            if 'data' not in data or 'list' not in data['data']:
                print("未找到备案APP信息")
                return []
                
            # 使用集合去重
            apps = set()
            for item in data['data']['list']:
                if 'name' in item and item['name']:
                    apps.add(item['name'])
            
            if apps:
                print(f"找到备案APP: {', '.join(apps)}")
            else:
                print("未找到备案APP")
                
            return list(apps)
            
        except Exception as e:
            print(f"解析备案APP数据失败: {str(e)}")
            return []
            
    except Exception as e:
        print(f"查询备案APP时出错: {str(e)}")
        return []

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
  python icpsearch_aqc.py -f input.xlsx -o output.xlsx
  python icpsearch_aqc.py -f input.xlsx -o output.xlsx -proxy http://127.0.0.1:8080
  python icpsearch_aqc.py -f input.xlsx -o output.xlsx -d 1  # 设置请求间隔为1秒
  python icpsearch_aqc.py -h  # 显示帮助信息
'''
        )
        parser.add_argument('-f', '--file', default='company_name.xlsx', help='指定输入Excel文件路径，默认为company_name.xlsx')
        parser.add_argument('-o', '--output', default='company_domains_result.xlsx', help='指定输出Excel文件路径，默认为company_domains_result.xlsx')
        parser.add_argument('-proxy', help='设置代理服务器，例如: http://127.0.0.1:8080')
        parser.add_argument('-d', '--delay', type=float, default=0, help='设置请求间隔时间（秒），默认为0秒')
        args = parser.parse_args()

        # 如果指定了代理，设置代理
        global PROXIES
        if args.proxy:
            PROXIES = {
                'http': args.proxy,
                'https': args.proxy
            }
            print(f"使用代理: {args.proxy}")

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
            
            # 获取企业ID
            company_id = get_company_id_aiqicha(company_name, is_first_query=(index == 0))
            if company_id:
                # 获取备案域名和IP
                data = get_icp_domains_aiqicha(company_id)
                domains = data['domains']
                ips = data['ips']
                
                # 获取微信小程序信息
                miniprograms = get_miniprograms(company_name)
                
                # 获取微信公众号信息
                wechat_accounts = get_wechat_accounts(company_id)
                
                # 获取备案APP信息
                apps = get_apps(company_id)
                
                # 格式化输出，使用换行符分隔
                domains_str = '\n'.join(domains) if domains else '无备案域名'
                ips_str = '\n'.join(ips) if ips else '无备案IP'
                miniprograms_str = '\n'.join(miniprograms) if miniprograms else '无备案微信小程序'
                wechat_accounts_str = '\n'.join(wechat_accounts) if wechat_accounts else '无备案微信公众号'
                apps_str = '\n'.join(apps) if apps else '无备案APP'
            else:
                domains_str = '未找到企业信息'
                ips_str = '未找到企业信息'
                miniprograms_str = '未找到企业信息'
                wechat_accounts_str = '未找到企业信息'
                apps_str = '未找到企业信息'
            
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
            
    except Exception as e:
        print(f"程序执行出错: {str(e)}")


if __name__ == "__main__":
    main()