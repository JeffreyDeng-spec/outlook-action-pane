import unittest, tempfile, subprocess, sys, os
from pathlib import Path
from worker_safety import LockBusy, exclusive_lock, account_state, worker_state_root
from windows_dpapi import protect, unprotect

class WorkerSafetyTests(unittest.TestCase):
    def test_path_isolation(self):
        with tempfile.TemporaryDirectory() as d:
            root = worker_state_root(d)
            self.assertNotEqual(account_state(root, "a@example.test", "a"), account_state(root, "b@example.test", "b"))
            for bad in ("../outside", "a/b", r"a\b", "C:", "name.", ".", ".."):
                with self.subTest(bad=bad), self.assertRaises(ValueError):
                    account_state(root, "a@example.test", bad)

    def test_process_contention_and_release(self):
        with tempfile.TemporaryDirectory() as d:
            lock = Path(d) / "singleton.lock"
            script = """import sys
from pathlib import Path
from worker_safety import LockBusy, exclusive_lock
try:
    with exclusive_lock(Path(sys.argv[1])):
        print('acquired')
except LockBusy:
    print('busy')
"""
            def attempt():
                return subprocess.check_output([sys.executable, "-B", "-c", script, str(lock)], text=True).strip()
            with exclusive_lock(lock):
                self.assertEqual(attempt(), "busy")
            self.assertEqual(attempt(), "acquired")

    @unittest.skipUnless(os.name == "nt", "Windows DPAPI only")
    def test_dummy_dpapi(self):
        plaintext = b"dummy-worker-token-never-used-for-real"
        cipher = protect(plaintext)
        self.assertNotIn(plaintext, cipher)
        self.assertEqual(unprotect(cipher), plaintext)
        with self.assertRaises(OSError):
            unprotect(bytes(len(cipher)))

if __name__ == "__main__":
    unittest.main(verbosity=2)
