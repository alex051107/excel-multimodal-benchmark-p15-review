# Data, Disclosure, and License Notices

This repository is publicly accessible and must be treated as a development-review snapshot, not a sealed benchmark.

The task portfolio combines public data extracts, project-authored professional scenarios, and project-authored synthetic document fixtures. The source mode and authenticity boundary are recorded in each task's `metadata/source_manifest.json` and `metadata/construction_manifest.json`.

Most task instructions are benchmark-authored professional scenarios. They are not direct transcripts of customer requests. Synthetic invoices, receipts, quotes, purchase orders, and statements are not real commercial or financial records.

## Disclosed benchmark material

Gold workbooks, solver scripts, task rubrics, evaluator code, and CONFIRM/private-labeled development assets were committed to this public repository in earlier releases. They must now be treated as disclosed. They cannot support future unseen-test, private-holdout, anti-contamination, or leakage-resistant claims.

Removing files from the current branch would not restore secrecy because the content remains in Git history. Any formal future evaluation must use rotated, materially distinct private families whose Gold, Oracle, evaluator truth, and held-out artifacts are stored outside this repository. A public release should expose only tasks and evidence that are intentionally designated for development.

## Prohibited material

Do not commit API keys, access tokens, cookies, account identifiers, token fingerprints, credential-bearing provider endpoints, approval files, request logs, raw provider responses, local absolute paths, Agent sessions, trajectories, or raw campaign runtime directories.

The current tracked-tree scan found no API credential, Bearer token, ZCloud endpoint, token fingerprint, or local `/Users/...` path. This statement applies to the inspected release tree; it is not a substitute for rotating credentials that may have appeared elsewhere.

Access to this repository does not grant permission to republish third-party source material beyond its original license. Consult each task's source manifest and [`references/THIRD_PARTY_NOTICES.md`](references/THIRD_PARTY_NOTICES.md) before reuse.
