"""Persistent paths for the approved Tokyo master workbook.

Code deployments may replace the application directory, so the active master
and its approved baseline must live outside that directory.  On the first run
only, the bundled files are migrated into the persistent location.
"""
import os
import shutil
from datetime import datetime, timezone
from supabase import create_client


APP_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLED_TEMPLATE_PATH = os.path.join(APP_DIR, 'tokyo_ordering_template.xlsm')
BUNDLED_BASELINE_PATH = os.path.join(APP_DIR, 'tokyo_template_baseline.json')
BUNDLED_DESSERT_TEMPLATE_PATH = os.path.join(APP_DIR, 'data', 'Tokyo_Dessert_Ordering.xlsm')
DESSERT_TEMPLATE_REVISION = '20260816-day3-profiterole'
TOKYO_TEMPLATE_REVISION = '20260831-sep-w1-master'


def _persistent_dir():
    configured = str(os.environ.get('OCTA_PERSISTENT_DATA_DIR') or '').strip()
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(os.path.expanduser('~'), '.octafood-data')


TOKYO_STORAGE_DIR = _persistent_dir()
TOKYO_TEMPLATE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'tokyo_ordering_template.xlsm')
TOKYO_BASELINE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'tokyo_template_baseline.json')
DESSERT_TEMPLATE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'Tokyo_Dessert_Ordering.xlsm')
DESSERT_REVISION_PATH = os.path.join(TOKYO_STORAGE_DIR, '.dessert-template-revision')
TOKYO_REVISION_PATH = os.path.join(TOKYO_STORAGE_DIR, '.tokyo-template-revision')
TOKYO_CLOUD_BUCKET = os.environ.get('SYSTEM_ASSETS_BUCKET', 'system-assets')
TOKYO_CLOUD_TEMPLATE_KEY = 'production-masters/tokyo_ordering_template.xlsm'
TOKYO_CLOUD_BASELINE_KEY = 'production-masters/tokyo_template_baseline.json'
_REMOTE_RESTORE_CHECKED = False


def _storage_client():
    url = str(os.environ.get('SUPABASE_URL') or '').strip()
    key = str(os.environ.get('SUPABASE_SERVICE_KEY') or '').strip()
    if not url or not key:
        return None
    return create_client(url, key).storage.from_(TOKYO_CLOUD_BUCKET)


def persist_tokyo_template_to_cloud(include_baseline=True):
    """Keep the approved Tokyo master outside the server filesystem."""
    storage = _storage_client()
    if storage is None:
        raise RuntimeError('إعدادات التخزين الدائم غير موجودة على السيرفر')
    with open(TOKYO_TEMPLATE_PATH, 'rb') as stream:
        storage.upload(TOKYO_CLOUD_TEMPLATE_KEY, stream.read(), file_options={
            'content-type': 'application/vnd.ms-excel.sheet.macroEnabled.12', 'upsert': 'true'
        })
    if include_baseline and os.path.exists(TOKYO_BASELINE_PATH):
        with open(TOKYO_BASELINE_PATH, 'rb') as stream:
            storage.upload(TOKYO_CLOUD_BASELINE_KEY, stream.read(), file_options={
                'content-type': 'application/json', 'upsert': 'true'
            })


def restore_tokyo_template_from_cloud_once():
    """Restore the latest user-uploaded master once per running worker."""
    global _REMOTE_RESTORE_CHECKED
    if _REMOTE_RESTORE_CHECKED:
        return False
    _REMOTE_RESTORE_CHECKED = True
    storage = _storage_client()
    if storage is None:
        return False
    try:
        template_bytes = storage.download(TOKYO_CLOUD_TEMPLATE_KEY)
        baseline_bytes = storage.download(TOKYO_CLOUD_BASELINE_KEY)
        if not template_bytes:
            return False
        os.makedirs(TOKYO_STORAGE_DIR, exist_ok=True)
        template_tmp = TOKYO_TEMPLATE_PATH + '.cloud.tmp'
        baseline_tmp = TOKYO_BASELINE_PATH + '.cloud.tmp'
        with open(template_tmp, 'wb') as stream:
            stream.write(template_bytes)
        with open(baseline_tmp, 'wb') as stream:
            stream.write(baseline_bytes)
        os.replace(template_tmp, TOKYO_TEMPLATE_PATH)
        os.replace(baseline_tmp, TOKYO_BASELINE_PATH)
        return True
    except Exception:
        return False


def _dessert_revision():
    try:
        with open(DESSERT_REVISION_PATH, 'r', encoding='utf-8') as stream:
            return stream.read().strip()
    except OSError:
        return ''


def _install_dessert_revision():
    if not os.path.exists(BUNDLED_DESSERT_TEMPLATE_PATH):
        return
    if _dessert_revision() == DESSERT_TEMPLATE_REVISION:
        return
    # One-time repair for the old bundled workbook that omitted the Day 3
    # Pistachio Profiterole mapping. Future restarts keep the active template,
    # including any newer template uploaded from the dashboard.
    shutil.copy2(BUNDLED_DESSERT_TEMPLATE_PATH, DESSERT_TEMPLATE_PATH)
    with open(DESSERT_REVISION_PATH, 'w', encoding='utf-8') as stream:
        stream.write(DESSERT_TEMPLATE_REVISION)


def _install_tokyo_revision():
    try:
        with open(TOKYO_REVISION_PATH, 'r', encoding='utf-8') as stream:
            installed = stream.read().strip()
    except OSError:
        installed = ''
    if installed == TOKYO_TEMPLATE_REVISION or not os.path.exists(BUNDLED_TEMPLATE_PATH):
        return
    if os.path.exists(TOKYO_TEMPLATE_PATH):
        # Never overwrite an already active workbook during code startup.  The
        # dashboard upload is the only safe way to approve a new Tokyo master.
        # This preserves files uploaded by the user across worker restarts and
        # deployments instead of silently rolling them back to the bundled copy.
        with open(TOKYO_REVISION_PATH, 'w', encoding='utf-8') as stream:
            stream.write(TOKYO_TEMPLATE_REVISION)
        return
    # Install the approved master only when persistent storage is empty.
    shutil.copy2(BUNDLED_TEMPLATE_PATH, TOKYO_TEMPLATE_PATH)
    if os.path.exists(BUNDLED_BASELINE_PATH):
        shutil.copy2(BUNDLED_BASELINE_PATH, TOKYO_BASELINE_PATH)
    with open(TOKYO_REVISION_PATH, 'w', encoding='utf-8') as stream:
        stream.write(TOKYO_TEMPLATE_REVISION)


def mark_tokyo_template_user_uploaded():
    """Record that the active Tokyo workbook was intentionally uploaded.

    The value only needs to exist; startup migration respects any existing
    active workbook and will not replace it with a bundled copy.
    """
    os.makedirs(TOKYO_STORAGE_DIR, exist_ok=True)
    with open(TOKYO_REVISION_PATH, 'w', encoding='utf-8') as stream:
        stream.write(f'user-uploaded:{datetime.now(timezone.utc).isoformat()}')


def ensure_tokyo_storage():
    """Create persistent storage and seed it once from the deployed bundle."""
    os.makedirs(TOKYO_STORAGE_DIR, exist_ok=True)
    migrations = (
        (BUNDLED_TEMPLATE_PATH, TOKYO_TEMPLATE_PATH),
        (BUNDLED_BASELINE_PATH, TOKYO_BASELINE_PATH),
        (BUNDLED_DESSERT_TEMPLATE_PATH, DESSERT_TEMPLATE_PATH),
    )
    for bundled_path, persistent_path in migrations:
        if not os.path.exists(persistent_path) and os.path.exists(bundled_path):
            shutil.copy2(bundled_path, persistent_path)
    _install_tokyo_revision()
    _install_dessert_revision()
    return TOKYO_TEMPLATE_PATH, TOKYO_BASELINE_PATH


ensure_tokyo_storage()
