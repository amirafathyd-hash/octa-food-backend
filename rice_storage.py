"""Persistent storage for the approved rice master workbook."""
import os
import shutil


APP_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLED_RICE_TEMPLATE_PATH = os.path.join(APP_DIR, 'data', 'Rice_Ordering_Sheet.xlsm')
RICE_TEMPLATE_REVISION = '20260817-approved-rice-master'


def _persistent_dir():
    configured = str(os.environ.get('OCTA_PERSISTENT_DATA_DIR') or '').strip()
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(os.path.expanduser('~'), '.octafood-data')


RICE_STORAGE_DIR = _persistent_dir()
RICE_TEMPLATE_PATH = os.path.join(RICE_STORAGE_DIR, 'Rice_Ordering_Sheet.xlsm')
RICE_REVISION_PATH = os.path.join(RICE_STORAGE_DIR, '.rice-template-revision')


def ensure_rice_storage():
    os.makedirs(RICE_STORAGE_DIR, exist_ok=True)
    try:
        with open(RICE_REVISION_PATH, 'r', encoding='utf-8') as stream:
            installed = stream.read().strip()
    except OSError:
        installed = ''

    if not os.path.exists(RICE_TEMPLATE_PATH) or installed != RICE_TEMPLATE_REVISION:
        if not os.path.exists(BUNDLED_RICE_TEMPLATE_PATH):
            raise FileNotFoundError('ملف الأرز الأساسي غير موجود داخل data')
        shutil.copy2(BUNDLED_RICE_TEMPLATE_PATH, RICE_TEMPLATE_PATH)
        with open(RICE_REVISION_PATH, 'w', encoding='utf-8') as stream:
            stream.write(RICE_TEMPLATE_REVISION)
    return RICE_TEMPLATE_PATH


ensure_rice_storage()
