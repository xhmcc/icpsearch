from PyQt5.QtCore import QThread, pyqtSignal
import re
import importlib
from ui_signals import continue_signal

icpmod = importlib.import_module("icpsearch_icp")

class BatchSearchWorker(QThread):
    progress = pyqtSignal(int)
    row_result = pyqtSignal(object)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, company_list, proxy_mode, proxy_value, delay):
        super().__init__()
        self.company_list = company_list
        self.proxy_mode = proxy_mode
        self.proxy_value = proxy_value
        self.delay = delay

    def run(self):
        try:
            self.batch_icp()
        except Exception as e:
            self.error.emit(str(e))
        self.finished.emit()

    def batch_icp(self):
        import time as _time
        proxies = None
        if self.proxy_mode == "代理":
            proxies = {'http': self.proxy_value, 'https': self.proxy_value}
        results = []
        name2row = {}
        for idx, company in enumerate(self.company_list):
            if re.match(r'[0-9\s~!@#$%^&*()_+\-=\[\]{};:\\|,.<>/?]', company):
                row = {
                    "企业名称": company,
                    "备案域名": "企业名称错误",
                    "备案IP": "",
                    "微信小程序": "",
                    "备案APP": ""
                }
                if company in name2row:
                    row_idx = name2row[company]
                    self.row_result.emit(("update", row_idx, row))
                else:
                    self.row_result.emit(("insert", None, row))
                    name2row[company] = len(name2row)
                continue
            try:
                icpmod.PROXIES = proxies
                data = None
                retry_count = 0
                while retry_count < 3:
                    start_time = _time.time()
                    try:
                        data = icpmod.get_icp_domains(company)
                        if (
                            isinstance(data, dict)
                            and data.get("code") == 200
                            and "params" in data
                            and isinstance(data["params"], dict)
                            and "list" in data["params"]
                            and isinstance(data["params"]["list"], list)
                            and len(data["params"]["list"]) == 0
                        ):
                            row = {
                                "企业名称": company,
                                "备案域名": "企业名称错误",
                                "备案IP": "",
                                "微信小程序": "",
                                "备案APP": ""
                            }
                            if company in name2row:
                                row_idx = name2row[company]
                                self.row_result.emit(("update", row_idx, row))
                            else:
                                self.row_result.emit(("insert", None, row))
                                name2row[company] = len(name2row)
                            break
                        if data is not None:
                            break
                    except Exception as e:
                        elapsed = _time.time() - start_time
                        if ("403" in str(e) or "超时" in str(e) or "timeout" in str(e).lower()):
                            retry_count += 1
                            if elapsed < 10:
                                _time.sleep(10 - elapsed)
                            if retry_count >= 3:
                                if self.proxy_mode == "代理":
                                    self.error.emit("当前IP或代理已封禁，查询终止。")
                                    return
                                else:
                                    raise
                        else:
                            raise
                if data is not None:
                    domains = data.get('domains', []) if data else []
                    ips = data.get('ips', []) if data else []
                    miniprograms = icpmod.get_miniprograms(company) if data else []
                    apps = icpmod.get_apps(company) if data else []
                    row = {
                        "企业名称": company,
                        "备案域名": "\n".join(domains) if domains else "",
                        "备案IP": "\n".join(ips) if ips else "",
                        "微信小程序": "\n".join(miniprograms) if miniprograms else "",
                        "备案APP": "\n".join(apps) if apps else ""
                    }
                    if company in name2row:
                        row_idx = name2row[company]
                        self.row_result.emit(("update", row_idx, row))
                    else:
                        self.row_result.emit(("insert", None, row))
                        name2row[company] = len(name2row)
                else:
                    row = {
                        "企业名称": company,
                        "备案域名": f"查询异常: 未返回数据",
                        "备案IP": "",
                        "微信小程序": "",
                        "备案APP": ""
                    }
                    if company in name2row:
                        row_idx = name2row[company]
                        self.row_result.emit(("update", row_idx, row))
                    else:
                        self.row_result.emit(("insert", None, row))
                        name2row[company] = len(name2row)
            except Exception as e:
                row = {
                    "企业名称": company,
                    "备案域名": f"查询异常: {str(e)}",
                    "备案IP": "",
                    "微信小程序": "",
                    "备案APP": ""
                }
                if company in name2row:
                    row_idx = name2row[company]
                    self.row_result.emit(("update", row_idx, row))
                else:
                    self.row_result.emit(("insert", None, row))
                    name2row[company] = len(name2row)
            self.progress.emit(int((idx + 1) / len(self.company_list) * 100))
            if self.delay > 0 and idx < len(self.company_list) - 1:
                self.msleep(int(self.delay * 1000))
        self.results = [row for _, row in sorted(zip(name2row.values(), name2row.keys()))] 