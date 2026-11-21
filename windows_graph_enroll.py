"""Owner-run OAuth code + PKCE for the Windows Graph metadata pilot."""
from __future__ import annotations
import argparse, base64, hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
import secrets
from time import monotonic
from urllib.parse import parse_qs, urlencode, urlsplit
from uuid import UUID
import webbrowser
from mail_accounts import GRAPH_KEYCHAINS
from outlook_triage import AUTH, GRAPH, PreviewError, fetch_json
from windows_credentials import DPAPIStore

AUTHZ="https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize"
SCOPE="https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadBasic offline_access"
PATH="/outlook-triage-callback"

def make_pkce():
    verifier=secrets.token_urlsafe(64)
    digest=hashlib.sha256(verifier.encode("ascii")).digest()
    return verifier,base64.urlsafe_b64encode(digest).rstrip(b"=").decode(),secrets.token_urlsafe(32)

def authorize_url(client_id,account,uri,challenge,state,*,scope=SCOPE):
    return AUTHZ+"?"+urlencode({"client_id":str(UUID(client_id)),"response_type":"code",
        "redirect_uri":uri,"response_mode":"query","scope":scope,"code_challenge":challenge,
        "code_challenge_method":"S256","state":state,"prompt":"select_account","login_hint":account})

def parse_callback(target,received_host,expected_host,state):
    try:
        p=urlsplit(target)
        q=parse_qs(p.query,keep_blank_values=True,max_num_fields=16)
    except ValueError:
        return None
    if (received_host!=expected_host or p.scheme or p.netloc or p.fragment or p.path!=PATH or
            len(q.get("state",[]))!=1 or not q["state"][0].isascii() or
            not secrets.compare_digest(q["state"][0],state)):
        return None
    if q.get("error"):
        return {"denied":True}
    codes=q.get("code",[])
    if not (len(codes)==1 and codes[0] and len(codes[0])<=4096):
        return None
    return {"code":codes[0]}

def receive_code(client_id,account,challenge,state,*,launch=webbrowser.open,timeout=180,scope=SCOPE):
    outcome={}
    class Server(HTTPServer):
        def get_request(self):
            sock,addr=super().get_request()
            sock.settimeout(3)
            return sock,addr

    class Callback(BaseHTTPRequestHandler):
        def log_message(self,*_):
            pass
        def do_GET(self):
            got=parse_callback(self.path,self.headers.get("Host"),host,state)
            if got:
                outcome.update(got)
            message=(b"Return to PowerShell for the authorization result." if got else
                     b"This request did not match the active authorization.")
            self.send_response(200 if got else 400)
            self.send_header("Content-Type","text/plain; charset=ascii")
            self.send_header("Cache-Control","no-store")
            self.send_header("Referrer-Policy","no-referrer")
            self.send_header("Content-Length",str(len(message)))
            self.end_headers()
            self.wfile.write(message)

    with Server(("127.0.0.1",0),Callback) as server:
        host=f"localhost:{server.server_port}"
        uri=f"http://{host}{PATH}"
        server.timeout=.5
        if not launch(authorize_url(client_id,account,uri,challenge,state,scope=scope)):
            raise PreviewError("Could not launch system browser")
        end=monotonic()+timeout
        while not outcome and monotonic()<end:
            server.handle_request()
    if "code" not in outcome:
        raise PreviewError("Authorization was declined or the local callback timed out")
    return outcome["code"],uri

def enroll(client_id,account,store,*,callback=receive_code,http=fetch_json):
    if account not in GRAPH_KEYCHAINS:
        raise ValueError("Unknown requested Outlook account")
    client_id=str(UUID(client_id))
    verifier,challenge,state=make_pkce()
    code,uri=callback(client_id,account,challenge,state)
    response=http(AUTH,"Microsoft authorization",form={"client_id":client_id,
        "grant_type":"authorization_code","code":code,"redirect_uri":uri,"code_verifier":verifier,
        "scope":SCOPE})
    token,refresh=response.get("access_token"),response.get("refresh_token")
    scopes={s.rsplit("/",1)[-1].lower() for s in response.get("scope","").split()}
    required={"user.read","mail.readbasic"}
    if not token or not refresh or not required<=scopes or not scopes<=required|{"openid","profile","email","offline_access"}:
        raise PreviewError("Microsoft did not return the approved basic read-only grant")
    who=http(f"{GRAPH}/me?%24select=mail,userPrincipalName","Microsoft identity",token=token)
    if account.lower() not in {str(who.get(k) or "").lower() for k in ("mail","userPrincipalName")}:
        raise PreviewError("Microsoft signed in a different account; nothing was saved")
    store.set("graph-read-basic",account,refresh)

def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("--client-id",required=True)
    p.add_argument("--account",choices=GRAPH_KEYCHAINS,required=True)
    p.add_argument("--credential-dir",type=Path,required=True)
    args=p.parse_args(argv)
    try:
        enroll(args.client_id,args.account,DPAPIStore(args.credential_dir))
    except (OSError,PreviewError,ValueError,TypeError) as e:
        p.exit(1,f"Authorization not installed: {type(e).__name__}\n")
    print("Graph metadata grant saved for the requested mailbox in the current Windows profile.")
    return 0

if __name__=="__main__":
    raise SystemExit(main())
