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


def _persistent_dir():
    configured = str(os.environ.get('OCTA_PERSISTENT_DATA_DIR') or '').strip()
    if configured:
        return os.path.abspath(os.path.expanduser(configured))
    return os.path.join(os.path.expanduser('~'), '.octafood-data')


TOKYO_STORAGE_DIR = _persistent_dir()
TOKYO_TEMPLATE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'tokyo_ordering_template.xlsm')
TOKYO_BASELINE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'tokyo_template_baseline.json')
DESSERT_TEMPLATE_PATH = os.path.join(TOKYO_STORAGE_DIR, 'Tokyo_Dessert_Ordering.xlsm')


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
    return TOKYO_TEMPLATE_PATH, TOKYO_BASELINE_PATH


ensure_tokyo_storage()
