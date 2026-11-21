# Contributing

Open a focused pull request against the default branch. Give it a title such as `feat(pane): show actionable invitations`. State what changes for a mailbox owner, what behavior stays compatible, how failure recovers and what checks passed. Include synthetic evidence for both an action and an ordinary notification.

Changes that write to Graph must specify how retries avoid duplicate effects. Use a local temporary SQLite database for tests. Check both account views independently and confirm the user owns final sending or submission. Keep real account setup and deployment credentials outside issue reports, logs and fixtures.
