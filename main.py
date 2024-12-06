import pandas as pd
import requests
import json
import time
import os
import re
from datetime import datetime
from jsonschema import validate, ValidationError

# 定义多语言消息
MESSAGES = {
    'zh': {
        'select_interface_language': '选择界面语言:\n1. 中文\n2. English',
        'invalid_selection': '无效的选择。请再试一次。',
        'enter_excel_path': '请输入Excel文件路径:',
        'file_not_found': '文件未找到。请再试一次。',
        'extract_uid_failed': '无法从文件名中提取UID。请输入UID:',
        'enter_uid': '请输入UID:',
        'select_game': '选择游戏:\n1. 原神\n2. 崩坏：星穹铁道',
        'invalid_game_selection': '无效的选择。请再试一次。',
        'select_language_code': '选择语言代码:\n1. 简体中文 (chs)\n2. 繁体中文 (cht)\n3. 日语 (jp)\n4. 英语 (en)\n5. 德语 (de)\n6. 西班牙语 (es)\n7. 法语 (fr)\n8. 印尼语 (id)\n9. 韩语 (kr)\n10. 葡萄牙语 (pt)\n11. 俄语 (ru)\n12. 泰语 (th)\n13. 越南语 (vi)',
        'invalid_language_code_selection': '无效的选择。请再试一次。',
        'download_dict_failed': '字典下载失败。请检查您的网络连接或稍后再试。',
        'dict_downloaded': '字典已下载并保存到 {path}',
        'dict_loaded': '字典已从 {path} 加载',
        'building_json': '正在构建UIGF JSON...',
        'json_validation_failed': 'JSON验证失败: {error}',
        'json_validation_success': 'JSON结构验证通过。',
        'json_exported': 'UIGF JSON已成功导出到 {path}',
        'prompt_continue': '您想继续吗？ (y/n):',
        'goodbye': '感谢使用，再见！',
        'invalid_uid': '无效的UID。请确保UID为数字。'
    },
    'en': {
        'select_interface_language': 'Select interface language:\n1. Chinese\n2. English',
        'invalid_selection': 'Invalid selection. Please try again.',
        'enter_excel_path': 'Please enter the Excel file path:',
        'file_not_found': 'File not found. Please try again.',
        'extract_uid_failed': 'Failed to extract UID from filename. Please enter UID:',
        'enter_uid': 'Please enter UID:',
        'select_game': 'Select game:\n1. Genshin Impact\n2. Honkai: Star Rail',
        'invalid_game_selection': 'Invalid selection. Please try again.',
        'select_language_code': 'Select language code:\n1. Simplified Chinese (chs)\n2. Traditional Chinese (cht)\n3. Japanese (jp)\n4. English (en)\n5. German (de)\n6. Spanish (es)\n7. French (fr)\n8. Indonesian (id)\n9. Korean (kr)\n10. Portuguese (pt)\n11. Russian (ru)\n12. Thai (th)\n13. Vietnamese (vi)',
        'invalid_language_code_selection': 'Invalid selection. Please try again.',
        'download_dict_failed': 'Failed to download dictionary. Please check your internet connection or try again later.',
        'dict_downloaded': 'Dictionary downloaded and saved to {path}',
        'dict_loaded': 'Dictionary loaded from {path}',
        'building_json': 'Building UIGF JSON...',
        'json_validation_failed': 'JSON validation failed: {error}',
        'json_validation_success': 'JSON structure validation passed.',
        'json_exported': 'UIGF JSON successfully exported to {path}',
        'prompt_continue': 'Do you want to continue? (y/n):',
        'goodbye': 'Thank you for using, goodbye!',
        'invalid_uid': 'Invalid UID. Please ensure UID is numeric.'
    }
}

# 祈愿池与其对应ID的映射
SHEET_TO_GACHA_TYPE = {
    "角色活动祈愿": "301",
    "武器活动祈愿": "302",
    "常驻祈愿": "200",
    "新手祈愿": "100"
}

# 字典下载配置
DICT_API_URL_TEMPLATE = 'https://api.uigf.org/dict/{game}/{lang}.json'
SCHEMA_URL = 'https://raw.githubusercontent.com/UIGF-org/UIGF-SchemaVerify/refs/heads/master/src/source/uigf-4.0-schema.json'

def get_message(lang, key):
    return MESSAGES[lang].get(key, '')

def select_interface_language():
    while True:
        print(MESSAGES['zh']['select_interface_language'])  # 默认显示中文提示
        selection = input().strip()
        if selection == '1':
            return 'zh'
        elif selection == '2':
            return 'en'
        else:
            print(MESSAGES['zh']['invalid_selection'])

def prompt_user(lang, key):
    return input(get_message(lang, key) + ' ').strip()

def extract_uid_from_filename(filename, lang):
    # 假设UID为连续的数字，位于文件名中的特定位置
    match = re.search(r'抽卡记录(\d+)_', filename)
    if match:
        return match.group(1)
    else:
        print(get_message(lang, 'extract_uid_failed'))
        while True:
            uid = input(get_message(lang, 'enter_uid') + ' ').strip()
            if uid.isdigit():
                return uid
            else:
                print(get_message(lang, 'invalid_uid'))

def select_game(lang):
    while True:
        print(get_message(lang, 'select_game'))
        selection = input().strip()
        if selection == '1':
            return 'genshin'
        elif selection == '2':
            return 'starrail'
        else:
            print(get_message(lang, 'invalid_game_selection'))

def select_language_code(lang):
    lang_codes = {
        '1': 'chs',
        '2': 'cht',
        '3': 'jp',
        '4': 'en',
        '5': 'de',
        '6': 'es',
        '7': 'fr',
        '8': 'id',
        '9': 'kr',
        '10': 'pt',
        '11': 'ru',
        '12': 'th',
        '13': 'vi'
    }
    while True:
        print(get_message(lang, 'select_language_code'))
        selection = input().strip()
        if selection in lang_codes:
            return lang_codes[selection]
        else:
            print(get_message(lang, 'invalid_language_code_selection'))

def download_dict(game, lang_code, local_path, lang):
    url = DICT_API_URL_TEMPLATE.format(game=game, lang=lang_code)
    try:
        response = requests.get(url)
        if response.status_code == 200:
            with open(local_path, 'w', encoding='utf-8') as f:
                f.write(response.text)
            print(get_message(lang, 'dict_downloaded').format(path=local_path))
            return json.loads(response.text)
        else:
            print(get_message(lang, 'download_dict_failed'))
            return {}
    except Exception as e:
        print(get_message(lang, 'download_dict_failed'))
        return {}

def load_dict(local_path, lang):
    try:
        with open(local_path, 'r', encoding='utf-8') as f:
            print(get_message(lang, 'dict_loaded').format(path=local_path))
            return json.load(f)
    except Exception as e:
        print(get_message(lang, 'download_dict_failed'))
        return {}

def read_excel(file_path, lang):
    # 读取所有工作表，同时解析 '时间' 列为 datetime
    try:
        xls = pd.ExcelFile(file_path)
        sheets = xls.sheet_names
        data = {}
        for sheet in sheets:
            try:
                df = pd.read_excel(xls, sheet, parse_dates=['时间'])
            except ValueError:
                # 如果 '时间' 列无法自动解析，手动指定日期格式
                df = pd.read_excel(xls, sheet, dtype={'时间': str})
                df['时间'] = pd.to_datetime(df['时间'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
            data[sheet] = df
        return data
    except Exception as e:
        print(get_message(lang, 'file_not_found'))
        return {}

def translate_names_to_ids(names, name_to_id_dict, lang):
    item_ids = []
    for name in names:
        item_id = name_to_id_dict.get(name, 0)
        if item_id == 0:
            warning_msg = "警告: 名称 '{name}' 未找到对应的物品ID。" if lang == 'zh' else f"Warning: Name '{name}' not found in dictionary."
            print(warning_msg)
        item_ids.append(item_id)
    return item_ids

def build_uigf_json(data, name_to_id_dict, uid, timezone, uigf_lang, lang):
    print(get_message(lang, 'building_json'))
    uigf = {
        "info": {
            "export_timestamp": int(time.time()),
            "export_app": "fxq2UIGF",
            "export_app_version": "1.0.0",
            "version": "v4.0"
        },
        "hk4e": []
    }

    all_records = []

    for sheet_name, df in data.items():
        # 根据预定义的映射获取祈愿类型ID
        gacha_type = SHEET_TO_GACHA_TYPE.get(sheet_name)
        if not gacha_type:
            warning_msg = f"警告: 工作表名称 '{sheet_name}' 未在映射表中找到对应的祈愿类型ID，使用默认值 '500'" if lang == 'zh' else f"Warning: Sheet name '{sheet_name}' not found in mapping table, using default '500'"
            print(warning_msg)
            gacha_type = "500"  # 默认值或处理错误

        # 提取名称列
        names = df['名称'].tolist()

        # 翻译名称为ID
        item_ids = translate_names_to_ids(names, name_to_id_dict, lang)

        # 构建记录
        for index, row in df.iterrows():
            # 获取对应的物品ID，防止翻译失败导致索引错误
            try:
                item_id = str(item_ids[index])
            except IndexError:
                item_id = '0'

            if item_id == '0':
                # 已在translate_names_to_ids中输出警告
                continue  # 跳过未找到的项目

            # 确保 '时间' 列为 datetime 对象
            if pd.isnull(row['时间']):
                warning_msg = f"警告: 工作表 '{sheet_name}' 的行 {index} 的 '时间' 值无效，跳过该记录。" if lang == 'zh' else f"Warning: Sheet '{sheet_name}', row {index} has invalid 'time' value, skipping record."
                print(warning_msg)
                continue

            # 格式化时间
            if isinstance(row['时间'], datetime):
                formatted_time = row['时间'].strftime('%Y-%m-%d %H:%M:%S')
            else:
                warning_msg = f"警告: 工作表 '{sheet_name}' 的行 {index} 的 '时间' 不是 datetime 对象，跳过该记录。" if lang == 'zh' else f"Warning: Sheet '{sheet_name}', row {index} 'time' is not a datetime object, skipping record."
                print(warning_msg)
                continue

            record = {
                "uigf_gacha_type": gacha_type,
                "gacha_type": gacha_type,
                "item_id": item_id,
                "count": "1",  # 假设每次祈愿获得1个物品
                "time": formatted_time,
                "name": row['名称'],
                "item_type": row['类别'],
                "rank_type": str(row['星级']),
                "id": str(row['祈愿 Id'])
            }
            all_records.append(record)

    # 添加到hk4e列表
    hk4e_entry = {
        "uid": uid,
        "timezone": timezone,
        "lang": uigf_lang,
        "list": all_records
    }

    uigf["hk4e"].append(hk4e_entry)
    return uigf

def validate_json(uigf_json, schema, lang):
    try:
        validate(instance=uigf_json, schema=schema)
        print(get_message(lang, 'json_validation_success'))
        return True
    except ValidationError as ve:
        print(get_message(lang, 'json_validation_failed').format(error=str(ve)))
        return False

def download_schema(schema_url):
    try:
        response = requests.get(schema_url)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Failed to download JSON Schema. Status code: {response.status_code}")
            return None
    except Exception as e:
        print(f"Error downloading JSON Schema: {e}")
        return None

def main():
    # 选择界面语言
    interface_lang = 'zh'

    while True:
        # 输入Excel文件路径
        excel_path = prompt_user(interface_lang, 'enter_excel_path')
        if not os.path.isfile(excel_path):
            print(get_message(interface_lang, 'file_not_found'))
            continue

        # 从文件名中提取UID
        filename = os.path.basename(excel_path)
        uid = extract_uid_from_filename(filename, interface_lang)

        # 选择游戏
        game = select_game(interface_lang)

        # 选择语言代码（用于字典下载）
        dict_lang_code = 'chs'

        # 设置UIGF的语言代码，根据界面语言
        uigf_lang_map = {
            'zh': 'zh-cn',
            'en': 'en-us'
        }
        uigf_lang = uigf_lang_map.get(interface_lang, 'zh-cn')

        # 设置本地字典文件路径
        dict_local_path = f'dict_{game}_{dict_lang_code}.json'

        # 检查并下载字典文件
        if not os.path.exists(dict_local_path):
            name_to_id = download_dict(game, dict_lang_code, dict_local_path, interface_lang)
        else:
            name_to_id = load_dict(dict_local_path, interface_lang)

        if not name_to_id:
            print(get_message(interface_lang, 'download_dict_failed'))
            continue

        # 读取Excel数据
        data = read_excel(excel_path, interface_lang)
        if not data:
            continue

        # 构建UIGF JSON
        timezone = 8  # 默认时区，可以根据需要调整
        uigf_json = build_uigf_json(data, name_to_id, uid, timezone, uigf_lang, interface_lang)

        # 下载并加载JSON Schema
        schema = download_schema(SCHEMA_URL)
        if not schema:
            print(get_message(interface_lang, 'json_validation_failed').format(error='Could not download schema'))
            continue

        # 验证JSON
        if not validate_json(uigf_json, schema, interface_lang):
            continue

        # 导出为JSON文件
        timestamp_str = datetime.now().strftime("%Y%m%d%H%M%S")
        output_json_filename = f'output_uigf_{uid}_{timestamp_str}.json'
        with open(output_json_filename, 'w', encoding='utf-8') as f:
            json.dump(uigf_json, f, ensure_ascii=False, indent=2)
        print(get_message(interface_lang, 'json_exported').format(path=output_json_filename))

        # 提示是否继续
        while True:
            cont = prompt_user(interface_lang, 'prompt_continue').lower()
            if cont in ['y', 'n']:
                break
            else:
                print(get_message(interface_lang, 'invalid_selection'))

        if cont != 'y':
            print(get_message(interface_lang, 'goodbye'))
            break

if __name__ == "__main__":
    main()
