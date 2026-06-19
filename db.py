import os
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
