"""Graph-only pilot adapter. No IMAP scope, mail-item reads or mutation."""
from threading import Lock
from time import monotonic
from urllib.parse import parse_qs, unquote, urlsplit

from outlook_triage import AUTH, GRAPH, PreviewError, fetch_json

class GraphMetadata:
    def __init__(self,account,store,client_id,*,http=fetch_json):
        self.account,self.store,self.client_id,self.http=account,store,client_id,http
        self.lock=Lock()
        self.cached=None
        self.folder_ids={}

    def graph_token(self):
        with self.lock:
            if self.cached and self.cached[0]>monotonic():
                return self.cached[1]
            refresh=self.store.get("graph-read-basic",self.account)
            if not refresh:
                raise PreviewError("Owner authorization is missing")
            response=self.http(AUTH,"Microsoft authorization",form={
                "client_id":self.client_id,"grant_type":"refresh_token","refresh_token":refresh,
                "scope":"https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadBasic offline_access",
            })
            token=response.get("access_token")
            scopes={s.rsplit("/",1)[-1].lower() for s in response.get("scope","").split()}
            required={"user.read","mail.readbasic"}
            if not token or not required<=scopes or not scopes<=required|{"offline_access","openid","profile","email"}:
                raise PreviewError("Graph pilot requires the dedicated basic read-only grant")
            who=self.http(f"{GRAPH}/me?%24select=mail,userPrincipalName","Microsoft identity",token=token)
            if self.account.lower() not in {str(who.get(k) or "").lower() for k in ("mail","userPrincipalName")}:
                raise PreviewError("Graph identity mismatched the selected account")
            for folder in ("inbox","junkemail"):
                result=self.http(f"{GRAPH}/me/mailFolders/{folder}?%24select=id","Microsoft folder identity",token=token)
                if not isinstance(result.get("id"),str) or not result["id"]:
                    raise PreviewError("Graph folder identity is missing")
                self.folder_ids[folder]=result["id"]
            if len(set(self.folder_ids.values()))!=2:
                raise PreviewError("Graph returned an ambiguous folder identity")
            rotated=response.get("refresh_token")
            if rotated and rotated!=refresh:
                self.store.set("graph-read-basic",self.account,rotated)
            try:
                max_age=max(1,min(3600,int(response.get("expires_in",0)))-60)
            except (TypeError,ValueError):
                max_age=1
            self.cached=(monotonic()+max_age,token)
            return token

    def folder_reader(self,folder):
        allowed={"id","receivedDateTime"}
        def read(url,label,*,token=None,headers=None):
            p=urlsplit(url)
            parts=[unquote(x) for x in p.path.split("/")[1:]]
            allowed_folders={folder,self.folder_ids[folder]}
            # Outlook returns opaque delta tokens with OData quoted-key paths,
            # while the initial Graph URL uses the slash-separated key form.
            # Convert only exact keys for this account and folder.
            if len(parts)==5 and parts[:2]==["v1.0","me"] and parts[3:]==["messages","delta"]:
                keys={f"mailFolders('{k}')":k for k in allowed_folders if "'" not in k}
                if parts[2] in keys:
                    parts=["v1.0","me","mailFolders",keys[parts[2]],"messages","delta"]
            params=parse_qs(p.query,keep_blank_values=True)
            if (p.scheme!="https" or p.netloc!="graph.microsoft.com" or p.fragment or
                parts[:3]!=["v1.0","me","mailFolders"] or
                len(parts) not in (5,6) or
                parts[3] not in allowed_folders or parts[4]!="messages" or
                (len(parts)==6 and parts[5]!="delta")):
                raise PreviewError("Graph metadata continuation escaped the selected folder")
            if set(params)-{"$select","$filter","$top","$skip","$skiptoken","$deltatoken"}:
                raise PreviewError("Unexpected Graph metadata query")
            selected=params.get("$select",[])
            continuation=any(k in params for k in ("$skiptoken","$deltatoken"))
            if ((selected and (len(selected)!=1 or set(selected[0].split(","))!=allowed)) or
                (not selected and not continuation)):
                raise PreviewError("Graph metadata projection must request only ID and receive time")
            result=self.http(url,label,token=token,headers=headers)
            rows=result.get("value")
            if isinstance(rows,list):
                result=dict(result)
                result["value"]=[{k:r[k] for k in ("id","receivedDateTime","@removed") if k in r}
                                  for r in rows if isinstance(r,dict)]
            return result
        return read
