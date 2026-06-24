import os
import time
from supabase import create_client

SUPABASE_URL = os.environ.get('SUPABASE_URL')
SUPABASE_KEY = os.environ.get('SUPABASE_SERVICE_KEY')  # service_role key (server-side only)

_client = None


def get_client():
    global _client
    if _client is None:
        if not SUPABASE_URL or not SUPABASE_KEY:
            raise RuntimeError(
                'لازم تحدد SUPABASE_URL و SUPABASE_SERVICE_KEY في متغيرات البيئة (Environment Variables)'
            )
        _client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _client


def execute_with_retry(query, max_attempts=3, delay_seconds=2):
    """Runs query.execute() with retries. The free hosting tier's outbound
    network has occasional brief hiccups (not a real bug) — retrying after a
    short pause usually succeeds without failing the whole request."""
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            return query.execute()
        except Exception as e:
            last_error = e
            if attempt < max_attempts:
                time.sleep(delay_seconds * attempt)
    raise last_error
