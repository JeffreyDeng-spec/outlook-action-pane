# Outlook Action Pane

Find what an Outlook message asks the recipient to do. The project builds on event-driven sorting for two personal Outlook mailboxes. The Windows worker listens for Inbox and Junk changes, reads Graph deltas and stores per-mailbox decisions in SQLite. A Mac pane can connect over an authenticated HTTPS bridge.

## Run the offline checks

Use Python 3.11 or newer on Windows. Run the commands at the repository root. The suite uses synthetic mailbox identities and temporary state directories.

```powershell
python -B -m unittest discover -p "test_*.py"
python tools/check_public_tree.py
```

For frontend acceptance with Windows Edge installed, run the following checks. They serve the actual pane and bridge locally with synthetic messages, then exercise task details, completion, undo, dismissal and disconnected states.

```powershell
python -m pip install playwright==1.58.0
python -B verify_action_pane.py --channel msedge
```

## Configuration and private state

[.env.example](.env.example) lists environment variable names. The code reads the process environment. Set variables for the Windows process and the Mac bridge as applicable. The example accounts only exercise offline fixtures. Both real mailboxes need to grant consent to the operator's own Microsoft application.

The production worker uses Windows user-scoped DPAPI for OAuth and provider credentials. Keep local SQLite databases, certificate material, enrollment artifacts and mailbox previews outside the repository. Provision a private bridge and the localhost add-in manifest for your own installation. Existing installations need their own account bindings and secure deployment configuration. This repository contains application source and synthetic tests. It does not provision an Entra application or deploy to a signed-in mailbox.

A single worker owns production mail moves. During migration, pause the previous writer, transfer consistent SQLite backups, verify both accounts' delta cursors and only then grant move authority. Preserve the latest state before any rollback.

## Contributions

Use the [contribution guide](CONTRIBUTING.md) for behavior changes and the [security policy](SECURITY.md) for private reports.
