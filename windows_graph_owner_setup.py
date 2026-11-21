"""Owner launch helper for two Graph-only consents and a no-move verification."""
from __future__ import annotations
import argparse, json, os
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from uuid import UUID
from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import atomic_report
from windows_credentials import DPAPIStore
from windows_graph_enroll import enroll
from windows_pilot import build_pilot
from worker_safety import exclusive_lock, LockBusy

PROFILE_APP_DIR="OutlookTriageGraphPilot"

def roots(profile):
    base=Path(profile)/PROFILE_APP_DIR
    if not base.is_absolute():
        raise ValueError("The current user's LocalAppData must be absolute")
    return base/"credentials",base/"metadata-state"

def provision_folders(auth,state):
    for path,name,expected in (
        (auth,".outlook-triage-dpapi",""),
        (state,".outlook-triage-pilot-only","graph-basic-polling-only"),
    ):
        if path.is_symlink() or path.parent.is_symlink():
            raise ValueError("Pilot paths cannot be symbolic links")
        exists=path.exists()
        path.mkdir(parents=True,exist_ok=True)
        marker=path/name
        if exists and (not marker.is_file() or marker.is_symlink() or
                marker.read_text(encoding="utf-8").strip()!=expected):
            raise ValueError("Existing directory is not an isolated Graph pilot")
        if not exists:
            with marker.open("x",encoding="utf-8") as handle:
                handle.write(expected)

def run(client_id,profile,*,enroll_account=enroll,provider=DPAPIStore,builder=build_pilot):
    auth,state=roots(profile)
    provision_folders(auth,state)
    summary=state/"verification.json"
    record={a:{"authorized":False,"identity_inbox_junk_delta":False} for a in GRAPH_KEYCHAINS}
    def save():
        atomic_report(summary,json.dumps({"mode":"graph-basic-polling-only",
            "timestamp_utc":datetime.now(timezone.utc).isoformat(),"accounts":record},indent=2))
    with exclusive_lock(state/"worker.lock"):
        save()
        store=provider(auth)
        for account in record:
            print(f"Microsoft browser consent: {account}",flush=True)
            try:
                enroll_account(client_id,account,store)
            except Exception as err:
                record[account]["error"]=type(err).__name__
                save()
                print(f"Not authorized for this mailbox ({type(err).__name__}).",flush=True)
                continue
            record[account]["authorized"]=True
            save()
        settings=SimpleNamespace(mode="graph-metadata",state_dir=state,credential_dir=auth,
            graph_client_id=client_id,imap_client_id=None,poll_seconds=120,check_once=True,
            confirm_read_only_pilot=True,tls_cert=None,tls_key=None)
        if any(r["authorized"] for r in record.values()):
            w,bearer=builder(settings,provider=provider)
            for account in record:
                if not record[account]["authorized"]:
                    continue
                try:
                    w.scan(account)
                except Exception as err:
                    record[account]["error"]=type(err).__name__
                    print(f"{account}: Graph verification did not complete ({type(err).__name__}).",flush=True)
                else:
                    record[account]["identity_inbox_junk_delta"]=True
                    print(f"{account}: Graph identity and both folder deltas verified.",flush=True)
                save()
    return 0 if all(v["identity_inbox_junk_delta"] for v in record.values()) else 1

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("--client-id")
    p.add_argument("--dry-run",action="store_true",help="Show candidate paths without auth, file writes or networking")
    p.add_argument("--pause",action="store_true",help="Keep the interactive console open after completion")
    args=p.parse_args(argv)
    profile=os.environ.get("LOCALAPPDATA")
    if os.name!="nt" or not profile:
        p.error("Run in the intended interactive Windows profile")
    if args.dry_run:
        auth,state=roots(profile)
        print("Dry run only. No files, credentials or browser requests opened.")
        print("DPAPI credential directory:",auth)
        print("Isolated metadata-only test directory:",state)
        return 0
    if not args.client_id:
        p.error("A verified dedicated Application (client) ID is required")
    try:
        cid=str(UUID(args.client_id))
        return run(cid,profile)
    except (OSError, ValueError, LockBusy) as err:
        print(f"Could not start read-only verification ({type(err).__name__}).")
        return 1
    finally:
        if args.pause:
            try:
                input("Press Enter to close this verification window.")
            except EOFError:
                pass

if __name__=="__main__":
    raise SystemExit(main())
