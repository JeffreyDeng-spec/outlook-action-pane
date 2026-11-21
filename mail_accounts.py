"""Mailbox configuration and stable per-account state directories."""
from hashlib import sha256
import os

accounts = tuple(address.strip().lower() for address in
                 os.getenv("OUTLOOK_TRIAGE_ACCOUNTS", "alpha@example.test,beta@example.test").split(","))
if len(accounts) != 2 or len(set(accounts)) != 2 or any(a.count("@") != 1 for a in accounts):
    raise ValueError("Set OUTLOOK_TRIAGE_ACCOUNTS to two distinct mailbox addresses")
DEFAULT_ACCOUNT = accounts[0]
GRAPH_KEYCHAINS = {a: f"outlook-triage-graph-{i}" for i, a in enumerate(accounts, 1)}


def account_state(root, account):
    return root / sha256(account.lower().encode()).hexdigest()[:12]
