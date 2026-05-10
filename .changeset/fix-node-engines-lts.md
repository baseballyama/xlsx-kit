---
'xlsx-kit': patch
---

Relax `engines.node` from `>=24.15.0` back to `>=22.0.0` so the published
package installs on every active Node LTS line (22.x, 24.x) plus current
(26.x), matching the CI matrix. 0.3.0 inadvertently shipped a Node 24+
floor that excluded the still-supported 22.x LTS; this restores broader
LTS coverage. The library does not rely on any Node 24-only API.
