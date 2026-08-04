"""用户ID授权管理：加载/保存已授权工号，解析批量输入。"""
import json
import os
import re
import sys

from constants import ADMIN_USER_ID, DEFAULT_AUTHORIZED_USER_IDS


def _ids_file_candidates():
    """授权ID存储文件：优先程序目录，其次用户目录"""
    if getattr(sys, 'frozen', False):
        # 打包后程序目录是临时解压目录，授权数据必须存到用户目录才能持久化
        return [os.path.join(os.path.expanduser('~'),
                             '.auto_report_authorized_ids.json')]
    script_dir = os.path.dirname(os.path.abspath(__file__))
    return [
        os.path.join(script_dir, 'authorized_ids.json'),
        os.path.join(os.path.expanduser('~'), '.auto_report_authorized_ids.json'),
    ]


def _default_ids():
    ids = [i.strip().upper() for i in DEFAULT_AUTHORIZED_USER_IDS if str(i).strip()]
    if ADMIN_USER_ID not in ids:
        ids.insert(0, ADMIN_USER_ID)
    return ids


def load_authorized_ids():
    """读取已授权工号列表（含管理员），读取失败时回退默认列表"""
    for path in _ids_file_candidates():
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            ids = [str(i).strip().upper() for i in data.get('authorized_ids', [])
                   if str(i).strip()]
            if ids:
                return ids
        except Exception:
            continue
    return _default_ids()


def save_authorized_ids(ids):
    """保存已授权工号列表，所有候选位置都失败时返回 False"""
    cleaned = sorted(set(str(i).strip().upper() for i in ids if str(i).strip()))
    payload = {'authorized_ids': cleaned}
    for path in _ids_file_candidates():
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            continue
    return False


def parse_id_input(text):
    """解析批量输入的工号：支持空格、逗号、分号、斜杠、顿号、换行等分隔"""
    parts = re.split(r'[\s,，;；/、|]+', text or '')
    return [p.strip().upper() for p in parts if p.strip()]
