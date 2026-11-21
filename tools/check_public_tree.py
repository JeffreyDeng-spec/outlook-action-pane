"""Check tracked and nonignored files before publication. Run from a Git checkout."""
from pathlib import Path
import re
import subprocess

root = Path(__file__).resolve().parents[1]
files = subprocess.check_output(["git", "-C", str(root), "ls-files", "-z",
                                 "--cached", "--others", "--exclude-standard"]).decode().split("\0")
secret = re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----|"
                    r"(?:gh[pousr]_[A-Za-z0-9]{25,}|github_pat_[A-Za-z0-9_]{25,})|"
                    r"\beyJ[A-Za-z0-9_-]{20,}\.eyJ[A-Za-z0-9_-]{15,}\.|"
                    r"https?://[^/\s]+\.ts\.net\b|\bS-1-5-21-\d+-\d+-\d+-\d+\b|"
                    r"(?i:[A-Z]:[\\/]Users[\\/][^\s\\/]+[\\/])")
email = re.compile(r"(?<![\w.+-])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b")
ok_ext = {".py", ".md", ".yml", ".yaml", ".css", ".html", ".js", ".xml", ".in"}
ok_name = {".gitignore", ".gitattributes", ".env.example"}
bad = []
for name in filter(None, files):
    file = root / name
    if file.is_symlink() or (file.suffix not in ok_ext and file.name not in ok_name):
        bad.append((name, "unexpected file type"))
        continue
    data = file.read_text(encoding="utf-8")
    if secret.search(data):
        bad.append((name, "credential, private hostname or user path"))
    if any(not m[0].lower().endswith((".test", ".invalid", ".example", "@example.org", "@example.com", "@example.net"))
           for m in email.finditer(data)):
        bad.append((name, "non-fixture mailbox address"))
for name, reason in bad:
    print(f"{name}: {reason}")
raise SystemExit(int(bool(bad)))
