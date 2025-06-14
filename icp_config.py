import yaml
import os

# 全局变量
PROXIES = None
AQC_COOKIE = None
TOKEN_MANAGER = None
should_continue_callback = None

# 读取配置
def load_config():
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    return config

def update_config_cookie(new_cookie):
    try:
        with open('config.yaml', 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        config['aiqicha_cookie'] = new_cookie
        with open('config.yaml', 'w', encoding='utf-8') as f:
            yaml.dump(config, f, allow_unicode=True)
        print("已更新config.yaml文件中的cookie")
        return True
    except Exception as e:
        print(f"更新config.yaml失败: {str(e)}")
        return False 