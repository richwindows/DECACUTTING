import json
import os
import re

# 默认设置
DEFAULT_SETTINGS = {
    'default': 233.0,
    'HMST-130-01': 181.0,
    'HMST-130-01B': 181.0,
    'HMST-130-02': 181.0,
    'HMST-82-01': 238.0,
    'HMST-82-02B': 233.0,
    'HMST-82-10': 238.0,
    'HMST-82-04': 257.0,
    'HMST-82-03': 257.0,
    'HMST-82-05': 257.0,
    'last_open_directory': '',
    'last_save_directory': ''
}

# JSON 文件路径
JSON_FILE_PATH = 'material_settings.json'

def load_settings():
    if os.path.exists(JSON_FILE_PATH):
        try:
            with open(JSON_FILE_PATH, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"警告：{JSON_FILE_PATH} 文件格式错误，使用默认设置")
    return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    with open(JSON_FILE_PATH, 'w') as f:
        json.dump(settings, f, indent=4)

def get_material_length(material):
    # print("get_material_length", material)
    settings = load_settings()
    
    # 使用正则表达式匹配材料名称的前两部分
    match = re.match(r'(HMST\d+-\d+)', material)
    if match:
        key = match.group(1).replace('HMST', 'HMST-')
        # print("key", key)
        # print("match settings.get(key, settings.get('default'))", settings.get(key, settings.get('default')))
        return settings.get(key, settings.get('default'))
    else:
        # print("settings.get('default')", settings.get('default'))
        return settings.get('default')

def get_last_directory(directory_type):
    settings = load_settings()
    return settings.get(f'last_{directory_type}_directory', '')

def update_last_directory(directory_type, path):
    settings = load_settings()
    settings[f'last_{directory_type}_directory'] = path
    save_settings(settings)

def get_settings():
    return load_settings()

def update_settings(new_settings):
    settings = load_settings()
    settings.update(new_settings)
    save_settings(settings)

