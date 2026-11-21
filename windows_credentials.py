"""Explicitly configured Windows credential store with account-bound DPAPI."""
from hashlib import sha256
import os
from pathlib import Path
import tempfile

from windows_dpapi import protect, unprotect

class DPAPIStore:
    def __init__(self, root):
        root = Path(root)
        if not root.is_absolute() or not root.is_dir() or not (root/".outlook-triage-dpapi").is_file():
            raise ValueError("A separately provisioned DPAPI directory is required")
        self.root = root.resolve()

    def _path(self, service, account):
        key = sha256(f"{service}\0{(account or '').lower()}".encode()).hexdigest()
        result = self.root / f"{key}.dpapi"
        if result.is_symlink():
            raise ValueError("DPAPI file may not be a symlink")
        return result

    def get(self, service, account):
        path = self._path(service, account)
        if not path.exists():
            return ""
        if not path.is_file() or path.stat().st_size > 32768:
            raise ValueError("Unexpected Windows credential blob")
        return unprotect(path.read_bytes()).decode("utf-8").strip()

    def set(self, service, account, value):
        if not isinstance(value,str) or not value or len(value)>24000:
            raise ValueError("Invalid credential value")
        path = self._path(service, account)
        ciphertext = protect(value.encode("utf-8"))
        with tempfile.NamedTemporaryFile("wb",dir=self.root,prefix=".rotate-",delete=False) as f:
            tmp=Path(f.name)
            try:
                f.write(ciphertext)
                f.flush()
                os.fsync(f.fileno())
            except BaseException:
                f.close()
                tmp.unlink(missing_ok=True)
                raise
        try:
            os.replace(tmp,path)
        finally:
            tmp.unlink(missing_ok=True)

def get_secret(service, account=None):
    root=os.environ.get("OUTLOOK_TRIAGE_DPAPI_DIR")
    return DPAPIStore(root).get(service,account) if root else ""
