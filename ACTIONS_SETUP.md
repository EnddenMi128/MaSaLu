# MaSaLu / MSL GitHub Actions

## What this adds

- Every push to main that changes website, GAS, Worker, tests or deployment config runs source checks and simulated Worker routing tests.
- Production deployment remains OFF until explicitly enabled. This protects the existing LINE Bot, Google Sheets and Apps Script deployment while credentials and source completeness are checked.
- Once enabled, Actions updates the existing MSL GAS Web App deployment, then the existing Cloudflare Worker. The current GitHub Pages branch-publishing setting continues to publish index.html automatically.

## Verified destinations

- Repository: EnddenMi128/MaSaLu
- GAS project (MSL) script ID: 1G9hsWk1qAm0G6fjccAAU1MhOxu2rpW7qytH_owosCzkHiUGRoFVKSCwv
- Cloudflare Worker name: leader-shift-line-webhook
- Worker health URL: https://leader-shift-line-webhook.edn869728.workers.dev/health

## Required GitHub Actions Secrets

In **Settings > Secrets and variables > Actions > Secrets**, add:

- `CLASPRC_JSON`: full clasp OAuth configuration for the Google account that can edit the MSL GAS project. Store only as an Actions Secret; never commit it.
- `MASALU_GAS_DEPLOYMENT_ID`: the existing MSL Web App deployment ID from **Deploy > Manage deployments**. This must refer to the existing /exec URL; it is NOT the script ID.
- `CLOUDFLARE_API_TOKEN`: Cloudflare API token with Worker edit permission for the account containing the existing Worker.
- `CLOUDFLARE_ACCOUNT_ID`: the Cloudflare account ID that owns that Worker.

Do not put the LINE channel secret, LINE access token or GAS shared secret into GitHub. They remain configured in Cloudflare Worker secrets and GAS Script Properties. The Wrangler configuration uses `keep_vars` to retain current dashboard variables; Cloudflare also preserves existing secrets on deploy.

## Required GitHub Actions Variables

In **Settings > Secrets and variables > Actions > Variables**, add:

1. `MASALU_GAS_SOURCE_CONFIRMED` = `true` **only after checking the remote MSL project contains no untracked Apps Script files**. `clasp push --force` is intentionally push-only and can overwrite remote scripts absent from GitHub.
2. `MASALU_AUTO_DEPLOY` = `true` **only after one successful manual deployment and live LINE testing**. Leave it unset or false until then.

### First release

1. Confirm MSL is the intended GAS project and that `GAS/程式碼.js`, `GAS/CalendarApi.js`, and `GAS/appsscript.json` are the complete authoritative source.
2. Inspect the existing Web App's deployment ID; do not generate a new deployment or URL.
3. Review Worker compatibility date and existing account bindings before the first Wrangler deployment.
4. Add all four secrets, and set `MASALU_GAS_SOURCE_CONFIRMED=true`.
5. Open GitHub **Actions > MaSaLu check and deploy > Run workflow**, check `Deploy MSL GAS and Cloudflare Worker to production`.
6. Verify `/health`, LIFF personal calendar, A01-A03 manager login, saveDay, LINE notifications, agree/disagree replies and authorization isolation.
7. Only after successful live checks, set `MASALU_AUTO_DEPLOY=true` for push-triggered releases.

The simulated tests DO NOT prove the production GAS Web App, LINE Messaging API or Google Sheets are functioning.

## Rollback

- Cloudflare: use Workers & Pages > leader-shift-line-webhook > Deployments > Rollback to the previous version.
- GAS: return to **Deploy > Manage deployments** and select the previous Apps Script version of the same deployment. Do not delete the deployment, or its /exec URL may change.
- GitHub Pages: revert the relevant index.html commit (not the entire repository).

## Security note

Before switching on automatic GAS deployment, verify the current CalendarApi.js does not expose complete employee rosters through a public JSONP endpoint and that non-managers see only their own schedule. A previously reviewed version still contained a legacy public read path. This is a production-release blocker if present.
