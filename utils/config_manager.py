import json
import os

CONFIG_FILE = "toolbox_config.json"

def load_config():
    """读取配置文件，如果不存在则返回空字典"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}

def save_config(new_data: dict):
    """保存配置，将新数据合并到原有配置中"""
    config = load_config()
    config.update(new_data)  # 合并已有配置
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)

def get_config_value(key, default=None):
    """获取配置项，支持默认值"""
    config = load_config()
    return config.get(key, default)

def set_config_value(key, value):
    """单独设置某个配置项"""
    config = load_config()
    config[key] = value
    save_config(config)
