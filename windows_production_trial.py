"""Bounded read-only production-consent validation; default is offline probe."""
from __future__ import annotations
import argparse,json,os
from pathlib import Path
from time import monotonic,sleep
from uuid import UUID

from imap_trigger import watch_folder
from mail_accounts import GRAPH_KEYCHAINS
from windows_pilot import ListenerSupervisor
from windows_production_grants import ProductionTrialAdapter,production_roots,production_store
from windows_worker import Worker
from worker_safety import exclusive_lock

def worker_candidate(profile,client,*,store_factory=production_store,adapter=ProductionTrialAdapter):
    _,state=production_roots(profile)
    if (state.is_symlink() or state.parent.is_symlink() or
        (state/".outlook-triage-production-trial").read_text(encoding="utf-8").strip()!="metadata-trial-no-moves"):
        raise ValueError("Existing isolated production-trial state directory required")
    store=store_factory(profile)
    return Worker(state,{a:adapter(a,store,client) for a in GRAPH_KEYCHAINS},
                  pilot=True,apply=False,folder_bound=True)

def trial_ready(reports):
    expected={"inbox","junkemail","graph"}
    return (set(reports)==set(GRAPH_KEYCHAINS) and
            all(set(flags)==expected and all(value is True for value in flags.values())
                for flags in reports.values()))

def listener_trial(worker,seconds,*,watcher=watch_folder,lock_path=None,shutdown_timeout=45):
    required=set(GRAPH_KEYCHAINS)
    if set(worker.adapters)!=required or worker.apply or not worker.pilot or not worker.folder_bound or worker.graph_only:
        raise ValueError("Production listener trial requires independently blocked moves and folder-bound Graph")
    deadline=monotonic()+seconds
    s=ListenerSupervisor(worker,watcher=watcher)
    reports=s.readiness()
    with exclusive_lock(lock_path if lock_path is not None else worker.root/"worker.lock"):
        try:
            s.start()
            while monotonic()<deadline:
                reports=s.readiness()
                if trial_ready(reports):
                    break
                sleep(min(.25,max(0,deadline-monotonic())))
        finally:
            stopped=s.close_before_unlock(shutdown_timeout)
    return reports,stopped

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("--mode",choices=("probe","listener-trial"),default="probe")
    p.add_argument("--client-id")
    p.add_argument("--seconds",type=int,default=90)
    p.add_argument("--confirm-live-read-only-listener-trial",action="store_true")
    args=p.parse_args(argv)
    if args.mode=="probe":
        print("Safe source probe: no Microsoft session or material state access")
        return 0
    if not (os.name=="nt" and args.client_id and 30<=args.seconds<=300
            and args.confirm_live_read_only_listener_trial):
        p.error("Read-only pilot additionally needs explicit approval and configured app")
    try:
        worker=worker_candidate(os.environ["LOCALAPPDATA"],str(UUID(args.client_id)))
        lock_root=Path("C:/ProgramData/OutlookTriage/state")
        # One known host-wide path, never per login session or pilot profile.
        if not lock_root.is_dir() or lock_root.resolve(strict=True)!=lock_root:
            raise ValueError("Trusted ProgramData production state root required")
        lock_path=lock_root/"worker.lock"
        if lock_path.is_symlink() or (lock_path.exists() and lock_path.resolve(strict=True)!=lock_path):
            raise ValueError("Production lock cannot be redirected")
        reports,stopped=listener_trial(worker,args.seconds,lock_path=lock_path)
    except (OSError,RuntimeError,ValueError):
        p.exit(1,"Read-only production-grant listener trial did not start.\n")
    print(json.dumps({"accounts":reports,"stopped":stopped,"Jev":False,"moves":False},separators=(",",":")))
    return 0 if stopped and trial_ready(reports) else 1

if __name__=="__main__":
    raise SystemExit(main())
