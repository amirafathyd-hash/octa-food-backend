"""Persistent paths for the approved Tokyo master workbook.

Code deployments may replace the application directory, so the active master
and its approved baseline must live outside that directory.  On the first run
only, the bundled files are migrated into the persistent location.
"""
import os
import shutil


APP_DIR = os.path.dirname(os.path.abspath(__file__))
BUNDLED_TEMPLATE_PATH = os.path.join(APP_DIR, 'tokyo_ordering_template.xlsm')
BUNDLED_BASELINE_PATH = os.path.join(APP_DIR, 'tokyo_template_baseline.json')
BUNDLED_DESSERT_TEMPLATE_PATH = os.path.join(APP_DIR, 'data', 'Tokyo_Dessert_Ordering.xlsm')
DESSERT_TEMPLATE_REVISION = '20260816-day3-profiterole'
TOKYO_TEMPLATE_REVISION = '20260816-w3-day4-explicit-mapping'


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
    # Install the approved W3 master once. The marker prevents refreshes and
    # worker restarts from replacing a newer workbook uploaded by the user.
    shutil.copy2(BUNDLED_TEMPLATE_PATH, TOKYO_TEMPLATE_PATH)
    if os.path.exists(BUNDLED_BASELINE_PATH):
        shutil.copy2(BUNDLED_BASELINE_PATH, TOKYO_BASELINE_PATH)
    with open(TOKYO_REVISION_PATH, 'w', encoding='utf-8') as stream:
        stream.write(TOKYO_TEMPLATE_REVISION)


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
