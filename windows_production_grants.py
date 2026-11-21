"""Separate Windows production-consent candidate; defaults to offline probe.

This harness cannot move mail or invoke classifiers. Explicit owner actions
to enroll or verify grants remain separate and disabled unless invoked.
"""
from __future__ import annotations
import argparse,json,os,re,ssl
from pathlib import Path
from threading import Lock
from time import monotonic
from uuid import UUID

from imap_trigger import IdleClient, SERVER, SCOPE as IMAP_SCOPE, junk_name
from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import AUTH, GRAPH, PreviewError, fetch_json
from windows_credentials import DPAPIStore
from windows_graph_enroll import make_pkce, receive_code
from windows_graph_metadata import GraphMetadata
from worker_safety import exclusive_lock

PROD_APP_DIR="OutlookTriageProductionTrial"
GRAPH_SCOPE="https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite"
SCOPES={"graph-write":GRAPH_SCOPE,"imap-examine":IMAP_SCOPE}
ACCEPT={"graph-write":{"user.read","mail.readwrite"},"imap-examine":{"imap.accessasuser.all"}}
OAUTH_NAMES={"offline_access","openid","profile","email"}

def production_roots(profile):
    root=Path(profile)/PROD_APP_DIR
    if not root.is_absolute():
        raise ValueError("Production trial profile path must be absolute")
    return root/"credentials",root/"metadata-trial"

def production_store(profile):
    auth,_=production_roots(profile)
    if auth.is_symlink() or auth.parent.is_symlink() or not (auth/".outlook-triage-production-trial").is_file():
        raise ValueError("Expected a separately provisioned production trial credential directory")
    return DPAPIStore(auth)

def scope(resource):
    return SCOPES[resource]+" offline_access"

def verify_result(resource,response):
    reported=response.get("scope")
    raw_scopes=reported.split() if isinstance(reported,str) else []
    s={k.rsplit("/",1)[-1].lower() for k in raw_scopes}
    token=response.get("access_token")
    expected=ACCEPT[resource]
    # Consumer Outlook can report both requested resources together.
    # Expected permissions for the requested resource must still be present.
    # Any permission outside the two approved resource grants is rejected.
    approved=ACCEPT["graph-write"]|ACCEPT["imap-examine"]|OAUTH_NAMES
    if not isinstance(token,str) or not token or not expected<=s or not s<=approved:
        safe_names=sorted(name for name in s if re.fullmatch(r"[a-z][a-z0-9_.-]{0,47}",name))
        summary=", ".join(safe_names[:12]) if safe_names else "none reported"
        raise PreviewError(f"Microsoft grant not accepted ({resource}); reported permission names: {summary}")
    return token


def verify_graph_account(account,token,http):
    who=http(f"{GRAPH}/me?%24select=mail,userPrincipalName","Graph trial account identity",token=token)
    if account.lower() not in {str(who.get(k) or "").lower() for k in ("mail","userPrincipalName")}:
        raise PreviewError("Microsoft signed in a different Graph account")

def verify_imap_account(account,token,*,connect=IdleClient):
    conn=connect(SERVER,ssl_context=ssl.create_default_context(),timeout=35)
    try:
        responses=iter((f"user={account}\x01auth=Bearer {token}\x01\x01".encode(),b""))
        conn.authenticate("XOAUTH2",lambda _:next(responses,b""))
        answer,mailboxes=conn.list()
        if answer!="OK":
            raise PreviewError("IMAP trial could not identify the folder list")
        # imaplib select(readonly=True) issues EXAMINE, not SELECT.
        for folder in ("INBOX",junk_name(mailboxes)):
            typ,_=conn.select(folder,readonly=True)
            if typ!="OK":
                raise PreviewError("IMAP read-only EXAMINE could not open both folders")
    finally:
        try:
            conn.shutdown()
        except OSError:
            pass

def enroll_production(client_id,account,resource,store,*,http=fetch_json,
                      callback=receive_code,imap_identity=verify_imap_account):
    client_id=str(UUID(client_id))
    if account not in GRAPH_KEYCHAINS or resource not in SCOPES:
        raise ValueError("Unexpected account or resource")
    verifier,challenge,state=make_pkce()
    code,uri=callback(client_id,account,challenge,state,scope=scope(resource),timeout=300)
    reply=http(AUTH,"Microsoft production trial authorization",form={"client_id":client_id,
        "grant_type":"authorization_code","code":code,"redirect_uri":uri,"code_verifier":verifier,
        "scope":scope(resource)})
    token=verify_result(resource,reply)
    refresh=reply.get("refresh_token")
    if not isinstance(refresh,str) or not refresh:
        raise PreviewError("Microsoft did not return an offline refresh grant")
    if resource=="graph-write":
        verify_graph_account(account,token,http)
    else:
        imap_identity(account,token)
    store.set(resource,account,refresh)

class ProductionTrialAdapter(GraphMetadata):
    """The consent is broad; the validation application is always read-only."""
    def __init__(self,account,store,client_id,*,http=fetch_json):
        super().__init__(account,store,client_id,http=http)
        self.resources={r:Lock() for r in SCOPES}
        self.tokens={}
    def _refresh(self,resource):
        with self.resources[resource]:
            entry=self.tokens.get(resource)
            if entry and entry[0]>monotonic():
                return entry[1]
            refresh=self.store.get(resource,self.account)
            if not refresh:
                raise PreviewError("Owner grant for this resource and account is absent")
            reply=self.http(AUTH,"Microsoft production trial authorization",form={"client_id":self.client_id,
                    "grant_type":"refresh_token","refresh_token":refresh,"scope":scope(resource)})
            token=verify_result(resource,reply)
            if resource=="graph-write":
                verify_graph_account(self.account,token,self.http)
                ids={}
                for folder in ("inbox","junkemail"):
                    result=self.http(f"{GRAPH}/me/mailFolders/{folder}?%24select=id",
                                     "Graph trial folder identity",token=token)
                    if not isinstance(result.get("id"),str) or not result["id"]:
                        raise PreviewError("Graph folder identity is missing")
                    ids[folder]=result["id"]
                if len(set(ids.values()))!=2:
                    raise PreviewError("Graph trial returned ambiguous folder identities")
                self.folder_ids=ids
            rotated=reply.get("refresh_token")
            if isinstance(rotated,str) and rotated and rotated!=refresh:
                self.store.set(resource,self.account,rotated)
            try: age=max(1,min(3600,int(reply.get("expires_in",0)))-60)
            except (ValueError,TypeError): age=1
            self.tokens[resource]=(monotonic()+age,token)
            return token
    def graph_token(self):
        return self._refresh("graph-write")
    def imap_token(self,account):
        if account!=self.account:
            raise PreviewError("IMAP account mismatch")
        return self._refresh("imap-examine")
    def write(self,*a,**kw):
        raise RuntimeError("Production pilot forbids moves")
    def classify(self,*a,**kw):
        raise RuntimeError("Production pilot forbids Jev")

def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode",choices=("probe","enroll"),default="probe")
    parser.add_argument("--client-id")
    parser.add_argument("--account",choices=GRAPH_KEYCHAINS)
    parser.add_argument("--resource",choices=SCOPES)
    parser.add_argument("--confirm-production-owner-consent",action="store_true")
    args=parser.parse_args(argv)
    if args.mode=="probe":
        print("Production preparation only: no account or socket access")
        return 0
    if (os.name!="nt" or not args.confirm_production_owner_consent or
            not (args.client_id and args.account and args.resource)):
        parser.error("This operation requires explicit account, resource and owner confirmation")
    try:
        store=production_store(os.environ["LOCALAPPDATA"])
        with exclusive_lock(store.root/"owner-enrollment.lock"):
            enroll_production(args.client_id,args.account,args.resource,store)
    except (OSError,ValueError,PreviewError) as e:
        parser.exit(1,f"Production consent not saved ({type(e).__name__}).\n")
    print("Selected production trial grant verified and encrypted. No mail was changed.")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
