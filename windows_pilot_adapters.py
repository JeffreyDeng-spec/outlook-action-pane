"""Microsoft credential/GET adapters that are never used by synthetic tests by default."""
from threading import Lock
from time import monotonic
from urllib.parse import urlsplit

from imap_trigger import SCOPE
from outlook_triage import AUTH, GRAPH, PreviewError, fetch_json

class MicrosoftReadOnly:
    def __init__(self, account, store, graph_client_id, imap_client_id, *, http=fetch_json):
        self.account,self.store,self.http=account,store,http
        self.clients={"graph-read":graph_client_id,"imap-idle":imap_client_id}
        self.lock={k:Lock() for k in self.clients}
        self.cache={}

    def read(self,url,label,*,token=None,headers=None):
        p=urlsplit(url)
        if p.scheme!="https" or p.netloc!="graph.microsoft.com" or not p.path.startswith("/v1.0/me"):
            raise PreviewError("Unexpected Graph read target")
        return self.http(url,label,token=token,headers=headers)

    def _access(self,resource):
        with self.lock[resource]:
            old=self.cache.get(resource)
            if old and old[0]>monotonic():
                return old[1]
            refresh=self.store.get(resource,self.account)
            if not refresh:
                raise PreviewError("Owner authorization is missing")
            scope="https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.Read" if resource=="graph-read" else SCOPE
            response=self.http(AUTH,"Microsoft authorization",form={
                "client_id":self.clients[resource],"grant_type":"refresh_token",
                "refresh_token":refresh,"scope":scope+" offline_access",
            })
            token=response.get("access_token")
            scopes={s.rsplit("/",1)[-1].lower() for s in response.get("scope","").split()}
            if not token:
                raise PreviewError("Microsoft did not grant access")
            if resource=="graph-read":
                if not {"user.read","mail.read"}.issubset(scopes) or scopes.intersection({"mail.readwrite","mail.readwrite.shared","mail.send","mail.send.shared"}):
                    raise PreviewError("Pilot refused a Graph token outside the read-only grant")
                who=self.read(f"{GRAPH}/me?%24select=mail,userPrincipalName","Microsoft identity",token=token)
                aliases={str(who.get(k) or "").lower() for k in ("mail","userPrincipalName")}
                if self.account.lower() not in aliases:
                    raise PreviewError("Graph identity mismatched the selected account")
            elif "imap.accessasuser.all" not in scopes:
                raise PreviewError("Microsoft did not grant IMAP access")
            rotated=response.get("refresh_token")
            if rotated and rotated!=refresh:
                self.store.set(resource,self.account,rotated)
            try:
                valid_for=max(1,min(3600,int(response.get("expires_in",0))) - 60)
            except (TypeError,ValueError):
                valid_for=1
            self.cache[resource]=(monotonic()+valid_for,token)
            return token

    def graph_token(self):
        return self._access("graph-read")

    def imap_token(self,account):
        if account!=self.account:
            raise PreviewError("IMAP identity mismatched the selected account")
        return self._access("imap-idle")

    def write(self,*args,**kwargs):
        raise RuntimeError("All mailbox moves are disabled during the pilot")

    def classify(self,*args,**kwargs):
        raise RuntimeError("Jev calls are disabled during the read-only pilot")
