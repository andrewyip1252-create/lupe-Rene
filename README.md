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

## Privacy / scrubbing notes

Before this was made public, the following were replaced with generic placeholders (real values live only in the actual deployed Apps Script projects):

- Real customer/prospect names and companies used as illustrative examples in bug-fix comments (e.g. "Contact A", "Company A") — the underlying logic and reasoning are unchanged, only the identifying names were swapped
- Google Sheet IDs (`YOUR_TEST_TRACKING_SHEET_ID_HERE` / `YOUR_PRODUCTION_TRACKING_SHEET_ID_HERE`)
- The HubSpot owner ID for the always-Thank-You-Card recipient (`YOUR_HUBSPOT_OWNER_ID_HERE`)
- The five base64-embedded attachment templates in `EmailAttachments.gs` (Contact Import, Properties Import, Outlook Integration, Deal Import Open/Closed) — the file structure and `getProspectsWelcomeAttachments()` logic are intact, but the actual template bytes are stripped. Re-populate these with `Utilities.base64Encode()` of your own real files if deploying this from scratch.
- The internal team roster (`VALID_REPS` / `REP_EMAILS`) — real employee names and `@creonesource.com` addresses replaced with generic placeholders (`Rep A` / `rep.a@example.com`, etc.), consistently across every file. The mapping logic itself is untouched; only the identities are placeholder'd.

No HubSpot token or Fathom API key was ever present in these files — both are pulled from Script Properties at runtime.

## Deploying a change

Each environment (`lupe/test/`, `lupe/production/`) maps 1:1 to its own Apps Script project. To deploy:

1. Copy the updated file(s) into the corresponding Apps Script project via the editor.
2. Save.
3. **Deploy → Manage deployments → (pencil icon on the active deployment) → New version → Deploy.**

Script Properties required in each project: `HUBSPOT_TOKEN`, `FATHOM_API_KEY`.
