"""Synthetic regression tests. No provider, mailbox, real data, or trusted host paths."""
import subprocess
import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Event, Thread
from unittest import TestCase
from unittest.mock import patch

from mail_accounts import GRAPH_KEYCHAINS
from windows_production_trial import listener_trial, trial_ready
from windows_worker import Worker
from worker_safety import exclusive_lock

A,B=tuple(GRAPH_KEYCHAINS)

class Adapter:
    def graph_token(self): return "DUMMY"
    def folder_reader(self,folder):
        return lambda url,label,**kw: {"value":[],"@odata.deltaLink":url}
    def imap_token(self,account): return "DUMMY"

class ListenerRegression(TestCase):
    def test_strict_report_rejects_empty_missing_wrong_or_truthy(self):
        good={"inbox":True,"junkemail":True,"graph":True}
        for bad in ({},{A:good},{A:{},B:good},{A:good,B:{**good,"graph":1}},
                    {A:good,B:{**good,"junkemail":False}},{A:good,B:good,"extra":good}):
            self.assertFalse(trial_ready(bad))
        self.assertTrue(trial_ready({A:good,B:good}))

    def test_expired_startup_has_no_stale_worker_readiness(self):
        with TemporaryDirectory() as temp, patch("socket.create_connection",side_effect=AssertionError):
            worker=Worker(Path(temp),{a:Adapter() for a in (A,B)},pilot=True,folder_bound=True)
            for a in (A,B):
                worker.last_scan_ok[a]=True
                for f in ("inbox","junkemail"):
                    worker.connectivity(a,f,True)
            def watcher(folder,stop,wake,status,account,*,get_token): stop.wait()
            report,closed=listener_trial(worker,0,watcher=watcher)
            self.assertTrue(closed)
            self.assertEqual(set(report),{A,B})
            self.assertTrue(all(not any(flags.values()) for flags in report.values()))
            self.assertFalse(trial_ready(report))

    def test_live_listeners_without_current_graph_success_are_not_ready(self):
        class Broken(Adapter):
            def folder_reader(self,folder):
                if folder=="junkemail": raise OSError("synthetic Graph failure")
                return super().folder_reader(folder)
        with TemporaryDirectory() as temp, patch("socket.create_connection",side_effect=AssertionError):
            worker=Worker(Path(temp),{A:Adapter(),B:Broken()},pilot=True,folder_bound=True)
            worker.last_scan_ok[B]=True
            def watcher(folder,stop,wake,status,account,*,get_token):
                status(folder,True); stop.wait()
            report,closed=listener_trial(worker,.5,watcher=watcher)
            self.assertTrue(closed)
            self.assertTrue(report[A]["graph"])
            self.assertFalse(report[B]["graph"])
            self.assertFalse(trial_ready(report))

    def test_real_second_process_refused_while_timed_out_thread_remains(self):
        with TemporaryDirectory() as temp, patch("socket.create_connection",side_effect=AssertionError):
            root=Path(temp); lock=root/"machine-shared.lock"
            worker=Worker(root,{a:Adapter() for a in (A,B)},pilot=True,folder_bound=True)
            stopping=Event(); release=Event(); timed_out=Event(); outcome=[]
            def watcher(folder,stop,wake,status,account,*,get_token):
                status(folder,True); stop.wait()
                if (account,folder)==(A,"Inbox"):
                    stopping.set(); release.wait(15)
            def run():
                from windows_pilot import ListenerSupervisor
                original=ListenerSupervisor.close
                def checking(supervisor,timeout=45):
                    result=original(supervisor,timeout)
                    if timeout is not None and not result: timed_out.set()
                    return result
                with patch.object(ListenerSupervisor,"close",checking):
                    outcome.append(listener_trial(worker,1,watcher=watcher,lock_path=lock,shutdown_timeout=.01))
            def second_process():
                code=("import sys\nfrom pathlib import Path\nfrom worker_safety import exclusive_lock,LockBusy\n"
                      "try:\n with exclusive_lock(Path(sys.argv[1])): pass\n"
                      "except LockBusy: sys.exit(23)\n")
                return subprocess.run([sys.executable,"-B","-c",code,str(lock)],timeout=5,capture_output=True).returncode
            first=Thread(target=run)
            first.start()
            try:
                self.assertTrue(stopping.wait(5))
                self.assertTrue(timed_out.wait(5))
                self.assertTrue(first.is_alive())
                self.assertEqual(second_process(),23)
            finally:
                release.set()
                first.join(10)
            self.assertFalse(first.is_alive())
            self.assertEqual(len(outcome),1)
            self.assertFalse(outcome[0][1])
            self.assertEqual(second_process(),0)

    def test_external_termination_waits_for_process_handle_before_probe(self):
        with TemporaryDirectory() as temp:
            root=Path(temp); lock=root/"machine-shared.lock"; ready=root/"started"
            code=("import sys,time\nfrom pathlib import Path\nfrom worker_safety import exclusive_lock\n"
                  "with exclusive_lock(Path(sys.argv[1])):\n Path(sys.argv[2]).touch(); time.sleep(120)\n")
            p=subprocess.Popen([sys.executable,"-B","-c",code,str(lock),str(ready)])
            try:
                from time import monotonic, sleep
                deadline=monotonic()+5
                while not ready.exists() and monotonic()<deadline: sleep(.02)
                self.assertTrue(ready.exists())
                from worker_safety import LockBusy
                with self.assertRaises(LockBusy):
                    with exclusive_lock(lock): pass
                p.terminate()
                p.wait(timeout=5)
                self.assertIsNotNone(p.returncode)
                with exclusive_lock(lock): pass
            finally:
                if p.poll() is None:
                    p.kill()
                    p.wait(timeout=5)
