"""Persistent storage for the approved sauce master workbook."""
import os
import shutil


APP_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLED_SAUCE_TEMPLATE_PATH = os.path.join(APP_DIR, 'data', 'Tokyo_Sauce.xlsm')
SAUCE_TEMPLATE_REVISION = '20260817-approved-count-driven-sauce-master'


def _persistent_dir():
    configured = str(os.environ.get('OCTA_PERSISTENT_DATA_DIR') or '').strip()
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(os.path.expanduser('~'), '.octafood-data')


SAUCE_STORAGE_DIR = _persistent_dir()
SAUCE_TEMPLATE_PATH = os.path.join(SAUCE_STORAGE_DIR, 'Tokyo_Sauce.xlsm')
SAUCE_REVISION_PATH = os.path.join(SAUCE_STORAGE_DIR, '.sauce-template-revision')
SAUCE_MAPPING_PATH = os.path.join(SAUCE_STORAGE_DIR, 'sauce-meal-mappings.json')


def ensure_sauce_storage():
    os.makedirs(SAUCE_STORAGE_DIR, exist_ok=True)
    try:
        with open(SAUCE_REVISION_PATH, 'r', encoding='utf-8') as stream:
            installed = stream.read().strip()
    except OSError:
        installed = ''
    if not os.path.exists(SAUCE_TEMPLATE_PATH) or installed != SAUCE_TEMPLATE_REVISION:
        if not os.path.exists(BUNDLED_SAUCE_TEMPLATE_PATH):
            raise FileNotFoundError('ملف الصوص الأساسي غير موجود داخل data')
        shutil.copy2(BUNDLED_SAUCE_TEMPLATE_PATH, SAUCE_TEMPLATE_PATH)
        with open(SAUCE_REVISION_PATH, 'w', encoding='utf-8') as stream:
            stream.write(SAUCE_TEMPLATE_REVISION)
    return SAUCE_TEMPLATE_PATH


ensure_sauce_storage()
