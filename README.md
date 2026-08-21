# CRE OneSource — Post-Onboard Customer Success Automation

Two Google Apps Script web apps that automate post-onboarding task creation and bulk contact pulling for CRE OneSource's Customer Success team.

- **LUPE** — Individual Onboarding. Pulls a single HubSpot contact, resolves the onboarder/salesperson from Fathom call recordings, and creates the correct set of SOP-driven HubSpot tasks (15 for Suite, 16 for Prospects, 9 for Deal View) with correct due dates, assignees, and notes.
- **RENE** — Bulk Onboarding. Pulls all contacts associated with a company's OSBO Pipeline deal, grouped by team role.

Both tools live together in a single Apps Script project per environment (Test / Production), sharing one HubSpot private app token and one Fathom API key via Script Properties. RENE's backend (`ReneCode.gs`) deliberately reuses LUPE's `hsFetch()`/`getProp()` helpers from `Code.gs` rather than duplicating them — Apps Script has no real per-file isolation, so every RENE function is prefixed `rene_` to avoid ever silently colliding with a LUPE function name.

## Repo structure

```
lupe/
  test/
    Code.gs                — backend: HubSpot + Fathom resolution, task creation, tracking sheet writes
    Index.html              — frontend: contact search, pull, task review, email sending
    EmailAttachments.gs      — base64-embedded attachment templates for the post-onboard email
  production/
    Code.gs
    Index.html
    EmailAttachments.gs
rene/
  ReneCode.gs               — RENE-specific backend functions (rene_ prefixed)
  ReneIndex.html            — RENE frontend
```

## Known gap

`Landing.html` (the shared chooser page that routes `?app=lupe` / `?app=rene`) isn't included yet — not currently on hand for either environment. Test's `Code.gs` already has the `doGet()` router wired to expect it; Production's `doGet()` intentionally does not have this router and serves LUPE directly as a standalone deployment.

## Deploying a change

Each environment (`lupe/test/`, `lupe/production/`) maps 1:1 to its own Apps Script project. To deploy:

1. Copy the updated file(s) into the corresponding Apps Script project via the editor.
2. Save.
3. **Deploy → Manage deployments → (pencil icon on the active deployment) → New version → Deploy.**

Script Properties required in each project: `HUBSPOT_TOKEN`, `FATHOM_API_KEY`.
