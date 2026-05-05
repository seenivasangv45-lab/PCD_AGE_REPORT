# gunicorn.conf.py
# Optimised for Render hosting — Monthly Report Dashboard (Flask)

import os

# ── Binding ────────────────────────────────────────────────────────────────────
bind = f"0.0.0.0:{os.environ.get('PORT', '5000')}"

# ── Workers ────────────────────────────────────────────────────────────────────
# Render free tier: 512 MB RAM. Keep workers low — uploads/outputs are written
# to disk per session so worker isolation matters less than memory.
workers = 2
worker_class = "sync"
threads = 2

# ── Timeouts ──────────────────────────────────────────────────────────────────
# Excel generation over multi-month ranges can be slow.
timeout = 120
keepalive = 5

# ── Upload size ────────────────────────────────────────────────────────────────
# Matches MAX_CONTENT_LENGTH = 200 MB in app.py
limit_request_line = 0
limit_request_fields = 200
limit_request_field_size = 0

# ── Logging ───────────────────────────────────────────────────────────────────
accesslog = "-"    # stdout → visible in Render log dashboard
errorlog  = "-"    # stderr
loglevel  = "info"

# ── Process naming ────────────────────────────────────────────────────────────
proc_name = "monthly_report_dashboard"
