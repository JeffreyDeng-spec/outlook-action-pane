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

For sparse or inconclusive previews, the Windows worker checks the same message's Graph plain-text body and supplies Jev with up to 1,200 characters of locally screened and cleaned evidence. Jev distinguishes a program-wide membership announcement from a change to a specific booking, invoice or member tier. Both accounts use this policy across senders. The previous sender mapping is no longer used. The worker preserves folder placement when it cannot verify the evidence, and retries temporary Graph failures.

### Update an existing Mac pane

Use a reviewed release checkout on the Mac after the Windows worker serves the action API. Pass the private project containing the installed add-in assets. The utility preserves the existing manifest, access path, TLS certificate, SSH tunnel configuration and handoff script. It saves the previous five service and asset files in a private backup, restarts the running pane LaunchAgent, then checks the live assets and both mailbox views over authenticated localhost HTTPS. A failed health check restores the previous files and restarts the pane. The existing Outlook ribbon label comes from the installed manifest.

```sh
python3 -B tools/install_mac_pane.py --project "/path/to/existing/outlook-triage"
python3 -B tools/install_mac_pane.py --project "/path/to/existing/outlook-triage" --verify-only
```

Supply `--state` when the existing Mac installation uses a custom state directory. For an existing Windows installation, the equivalent owner-scoped host is `windows_pane.py`. Its private `windows-pane/manifest.xml` is generated at install time and contains the local access path.

## Contributions

Use the [contribution guide](CONTRIBUTING.md) for behavior changes and the [security policy](SECURITY.md) for private reports.
