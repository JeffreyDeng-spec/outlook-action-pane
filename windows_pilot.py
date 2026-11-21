"""Source-only Windows pilot; default probe never opens state, credentials or sockets."""
from __future__ import annotations
import argparse
from contextlib import contextmanager
import os
from pathlib import Path
from threading import Event, Lock, Thread
from time import monotonic
from uuid import UUID

from imap_trigger import watch_folder
from mail_accounts import GRAPH_KEYCHAINS
from windows_credentials import DPAPIStore
from windows_pilot_adapters import MicrosoftReadOnly
from windows_graph_metadata import GraphMetadata
from windows_worker import Worker, locked_tls_server
from worker_safety import exclusive_lock, worker_state_root

class ListenerSupervisor:
    def __init__(self,worker,*,watcher=watch_folder,poll_seconds=None):
        self.worker,self.watcher=worker,watcher
        self.poll_seconds=poll_seconds
        if worker.graph_only and (poll_seconds is None or poll_seconds<=0):
            raise ValueError("Graph-only requires periodic polling")
        self.stop=Event()
        self.arrivals={a:Event() for a in worker.adapters}
        self.threads=[]
        self.started=False
        self._readiness_lock=Lock()
        self._readiness={a:{"inbox":False,"junkemail":False,"graph":False} for a in self.arrivals}
        self._done=[]

    def readiness(self):
        with self._readiness_lock:
            return {a:dict(flags) for a,flags in self._readiness.items()}

    def _ready(self,account,path,ok):
        with self._readiness_lock:
            self._readiness[account][path]=ok is True

    def _set_connection(self,account,path,ok):
        self.worker.connectivity(account,path,ok)
        self._ready(account,path,ok)

    @staticmethod
    def _run_job(target,args,done):
        try:
            target(*args)
        finally:
            done.set()

    def wake(self,account):
        if not self.stop.is_set():
            self.arrivals[account].set()

    def sync_loop(self,account):
        pending=self.arrivals[account]
        delay=1
        while not self.stop.is_set():
            if not pending.wait(.2):
                continue
            if self.stop.is_set():
                break
            pending.clear()
            try:
                self._ready(account,"graph",self.worker.scan(account) is True)
                delay=1
            except Exception:
                self._ready(account,"graph",False)
                if self.stop.wait(delay):
                    break
                pending.set()
                delay=min(30,delay*2)

    def poll(self):
        while not self.stop.wait(self.poll_seconds):
            for account in self.arrivals:
                self.wake(account)

    def listen(self,account,folder):
        name="inbox" if folder=="Inbox" else "junkemail"
        try:
            self.watcher(folder,self.stop,lambda: self.wake(account),
                         lambda _,ok: self._set_connection(account,name,ok),account,
                         get_token=self.worker.adapters[account].imap_token)
        finally:
            self._set_connection(account,name,False)

    def start(self):
        if self.started or self.stop.is_set():
            raise RuntimeError("Supervisor cannot start twice")
        self.started=True
        jobs=[]
        for account in self.arrivals:
            jobs.append((self.sync_loop,(account,)))
            if not self.worker.graph_only:
                jobs.extend((self.listen,(account,folder)) for folder in ("Inbox","Junk"))
        if self.worker.graph_only:
            jobs.append((self.poll,()))
        for target,args in jobs:
            done=Event()
            t=Thread(target=self._run_job,args=(target,args,done),daemon=False)
            t.start()
            self.threads.append(t)
            self._done.append(done)
        for account in self.arrivals:
            self.wake(account)
        return self

    def close(self,timeout=45):
        self.stop.set()
        for a in self.arrivals.values(): a.set()
        until=None if timeout is None else monotonic()+timeout
        for t,done in zip(self.threads,self._done):
            def remaining():
                return None if until is None else max(0,until-monotonic())
            if not done.wait(remaining()):
                return False
            t.join(remaining())
            if t.is_alive():
                return False
        return True

    def close_before_unlock(self,timeout=45):
        # A stop deadline is a result/reporting deadline, not permission to
        # unlock. If necessary an external supervisor can terminate the whole
        # process and wait for its handle before allowing a restart.
        try:
            if self.close(timeout):
                return True
        finally:
            while True:
                try:
                    if self.close(None):
                        break
                except KeyboardInterrupt:
                    continue
        return False

def build_pilot(args,*,provider=DPAPIStore,adapter=MicrosoftReadOnly,graph_adapter=GraphMetadata):
    if not args.confirm_read_only_pilot:
        raise ValueError("Explicit pilot enablement required")
    root=worker_state_root(str(args.state_dir))
    graph_only=getattr(args,"mode",None)=="graph-metadata"
    marker="graph-basic-polling-only" if graph_only else "pilot-metadata-only"
    if not root.is_dir() or (root/".outlook-triage-pilot-only").read_text(encoding="utf-8").strip()!=marker:
        raise ValueError("Dedicated isolated state directory is required")
    if not args.credential_dir.is_absolute():
        raise ValueError("Dedicated credential directory is required")
    secrets_root=args.credential_dir.resolve(strict=True)
    if root==secrets_root or root in secrets_root.parents or secrets_root in root.parents:
        raise ValueError("Credentials and pilot state must not overlap")
    UUID(args.graph_client_id)
    if graph_only:
        if args.imap_client_id or not 30<=args.poll_seconds<=3600:
            raise ValueError("Graph-only mode rejects IMAP configuration and requires 30–3600s polling")
    else:
        UUID(args.imap_client_id)
        if args.graph_client_id==args.imap_client_id:
            raise ValueError("Graph and IMAP pilot app registrations must be separate")
    once=getattr(args,"check_once",False)
    if once and not graph_only:
        raise ValueError("One-shot verification only works for Graph metadata")
    if not once and not all(p.is_absolute() and p.is_file() for p in (args.tls_cert,args.tls_key)):
        raise ValueError("Existing absolute TLS files required")
    store=provider(secrets_root)
    bearer=None if once else store.get("graph-basic-bridge-bearer" if graph_only else "bridge-bearer",None)
    if graph_only:
        adapters={a:graph_adapter(a,store,args.graph_client_id) for a in GRAPH_KEYCHAINS}
    else:
        adapters={a:adapter(a,store,args.graph_client_id,args.imap_client_id) for a in GRAPH_KEYCHAINS}
    return Worker(root,adapters,pilot=True,graph_only=graph_only),bearer

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("--mode",choices=("probe","pilot","graph-metadata"),default="probe")
    p.add_argument("--confirm-read-only-pilot",action="store_true")
    p.add_argument("--check-once",action="store_true")
    for name in ("state-dir","credential-dir","tls-cert","tls-key"):
        p.add_argument("--"+name,type=Path)
    p.add_argument("--graph-client-id")
    p.add_argument("--imap-client-id")
    p.add_argument("--port",type=int,default=18765)
    p.add_argument("--poll-seconds",type=int,default=120)
    args=p.parse_args(argv)
    if args.mode=="probe":
        print("Source-only probe: no listener, login, data, or credential access")
        return 0
    if os.name!="nt":
        p.error("Pilot requires Windows")
    if (not all(getattr(args,k) for k in ("state_dir","credential_dir","graph_client_id")) or
            (not args.check_once and not all((args.tls_cert,args.tls_key))) or
            (args.mode=="pilot" and not args.imap_client_id)):
        p.error("Pilot requires explicit state, credential, app registrations and TLS paths")
    try:
        w,bearer=build_pilot(args)
    except (OSError,ValueError,TypeError):
        p.error("Read-only pilot configuration not accepted")
    if args.check_once:
        result=0
        with exclusive_lock(w.root/"worker.lock"):
            for account in w.adapters:
                try:
                    w.scan(account)
                except Exception as exc:
                    print(f"{account}: Graph identity/folder check failed ({type(exc).__name__})")
                    result=1
                else:
                    print(f"{account}: Graph account identity, Inbox and Junk verified; queued metadata IDs: {w.metadata_counts[account]}")
        return result
    supervisor=ListenerSupervisor(w,poll_seconds=args.poll_seconds)
    # Pane scan is queued. It never invokes a second worker or changes apply mode.
    w.submit=supervisor.wake
    with locked_tls_server(w,bearer,args.tls_cert,args.tls_key,args.port) as server:
        try:
            supervisor.start()
            server.serve_forever()
        except KeyboardInterrupt:
            pass
        finally:
            if not supervisor.close_before_unlock():
                raise SystemExit("Pilot listener shutdown deadline exceeded")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
