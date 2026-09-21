# Codex Handoff — Short-Lived Certificate Checkout for App-Only M365 Access

## Read first
Inspect the repo before writing anything. Match its existing layout, Terraform module conventions, naming, and CI workflows. Where this document conflicts with the repo's conventions, the repo wins. List any conflict in your summary.

## Objective
Let an authorized user "check out" an app-only identity for about one hour to run ad-hoc PowerShell against Graph, Exchange Online, and Teams. No long-lived credential exists anywhere. The user never receives a private key from the service.

## Current state
- Azure Automation account with a **system-assigned managed identity (SAMI)**.
- One **app registration** holds the Graph application permissions. Its service principal is being granted the directory roles (Exchange Administrator, Teams Reader, others).
- The SAMI is being granted `Application.ReadWrite.OwnedBy` and ownership of that app registration.
- Tenant-side grants are done by a human. See "Manual steps". Do not attempt them.
- Repo already targets GitHub-as-source-of-truth: Terraform manages runbooks, GitHub Actions validates PRs, OIDC federation, no stored secrets.

## Design

### Flow
1. User runs `Request-IdentityCheckout` locally.
2. Function creates a self-signed cert in `Cert:\CurrentUser\My`. Private key is **non-exportable**. Validity: 1 hour.
3. Function sends only the **public certificate** (base64 DER) to the runbook `Grant-IdentityCheckout` via `Start-AzAutomationRunbook`.
4. Runbook, running as the SAMI, validates the cert and appends it to the app registration's `keyCredentials` with a 1-hour end date.
5. Function waits for the job, then for propagation, then returns AppId, TenantId, Thumbprint.
6. User connects with `Connect-CheckedOutIdentity`.
7. Scheduled runbook `Remove-ExpiredCheckoutKeys` deletes expired checkout keys.

### Runbook: `Grant-IdentityCheckout`
Parameters: `PublicCertBase64` (string, mandatory), `Reason` (string, mandatory).

Must:
- `Connect-MgGraph -Identity -NoWelcome`.
- Parse the input as `X509Certificate2`. Reject if it fails to parse.
- Reject if: key is not RSA >= 2048 (or ECDSA P-256+), `NotAfter` is more than 65 minutes out, `NotBefore` is in the future by more than 5 minutes, or the input contains a private key.
- Set `endDateTime` = the earlier of cert `NotAfter` and now + 60 minutes (UTC).
- Set `displayName` = `checkout-<jobId>-<yyyyMMddTHHmmssZ>`. The `checkout-` prefix is how cleanup identifies these keys.
- Read current `keyCredentials`, append, write back the **full list**. The write replaces the collection. Dropping existing entries deletes them.
- Do not use `addKey`. It requires proof of possession of an existing key, which a managed identity does not have.
- Re-read after write and confirm both the new key and all prior keys are present. Retry up to 3 times with backoff if the new key is missing (concurrent job overwrote it). Fail loudly if a prior key went missing.
- Emit one structured JSON output record: jobId, keyId, thumbprint, endDateTime, reason. No cert bytes in logs.
- Target app object ID comes from an Automation variable or runbook parameter default set by Terraform. Not hardcoded.

### Runbook: `Remove-ExpiredCheckoutKeys`
- Schedule: every 15 minutes.
- Remove only entries where `displayName` starts with `checkout-` **and** `endDateTime` < now (UTC).
- Never touch keys without the prefix.
- Same read-modify-write and verify pattern as above.
- Log each removed keyId.

### Client module: `IdentityCheckout`
Functions:
- `Request-IdentityCheckout -Reason <string>` — steps 1–5 above. Poll the job with a timeout. After success, wait for propagation (start at 60 s, then retry token acquisition up to ~3 minutes).
- `Connect-CheckedOutIdentity -Service Graph|ExchangeOnline|Teams[]`:
  - Graph: `Connect-MgGraph -ClientId -TenantId -CertificateThumbprint`
  - Exchange: `Connect-ExchangeOnline -AppId -CertificateThumbprint -Organization <primary .onmicrosoft.com domain>`
  - Teams: `Connect-MicrosoftTeams -ApplicationId -CertificateThumbprint -TenantId`
- `Clear-IdentityCheckout` — disconnect sessions and delete the local cert.

Configuration (tenant ID, app client ID, organization domain, subscription, resource group, Automation account, runbook name) comes from a config file or parameters. No IDs hardcoded in function bodies.

**Microsoft Places:** `Connect-MicrosoftPlaces` has no documented app-only sign-in. Out of scope. Document the workaround in the README: Graph Places API (`Place.Read.All` / `Place.ReadWrite.All`) or `Get-Place` / `Set-Place` through the Exchange session.

### Terraform
- `azurerm_automation_runbook` for both runbooks, `content = file(...)`, PowerShell 7.2 runtime.
- Automation modules: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Applications`. Pin versions.
- `azurerm_automation_schedule` + job schedule link for the cleanup runbook.
- Automation variable for the target app object ID.
- Azure RBAC: `Automation Job Operator` for the checkout users group, scoped as narrowly as the provider allows (runbook scope preferred over account scope). Group object ID is a variable.
- Do not manage Entra directory roles, Graph app role assignments, or app ownership in Terraform unless the repo already does so with an approved provider. Default: document in `MANUAL-STEPS.md`.

### CI
- PSScriptAnalyzer on all `.ps1` / `.psm1`.
- Pester tests with Graph cmdlets mocked. Minimum cases:
  - Valid cert appended; existing keys preserved.
  - Rejected: malformed input, weak key, validity > 65 min, private key present.
  - Cleanup removes only expired `checkout-` keys; leaves unexpired and non-prefixed keys.
  - Verify-after-write retry path.
- `terraform fmt -check`, `terraform validate`, plan only.

## Guardrails
- No `terraform apply`. Plan only.
- No changes to the tenant, the app registration, or role assignments. No live Graph write calls during development or tests.
- No secrets, certificates, thumbprints, or tokens committed. No real tenant or object IDs in code; use variables and placeholders.
- Never write `keyCredentials` without first reading and preserving existing entries.
- Never log certificate bytes or tokens.
- Do not add `Application.ReadWrite.All` anywhere. `OwnedBy` only.
- Privileged steps go in `MANUAL-STEPS.md`, not in scripts that execute them.

## Manual steps (human only — write these into `MANUAL-STEPS.md`, status "verify")
1. Grant Graph app role `Application.ReadWrite.OwnedBy` to the SAMI (`New-MgServicePrincipalAppRoleAssignment`). Allow up to an hour for token cache.
2. Add the SAMI as **owner** of the app registration (`New-MgApplicationOwnerByRef`). Portal cannot do this.
   - Alternative to 1+2: custom directory role with `microsoft.directory/applications/credentials/update`, assigned at the app registration scope.
3. Assign directory roles to the **app registration's service principal** as Active assignments (Exchange Administrator, Teams Reader, others as required).
4. Add `Exchange.ManageAsApp` (Office 365 Exchange Online) to the app registration; grant admin consent.
5. Add Graph application permissions required by Teams app-only auth (`Organization.Read.All` baseline; others per cmdlet); grant admin consent.
6. Create the checkout users group; supply its object ID to Terraform.

## Known limitations — state these in the README
- During the window the user holds a working app-only credential and can do anything the app can, not just what a given script does.
- Access tokens obtained just before key expiry remain valid for roughly another hour.
- Entra sign-in logs show the app, not the person. Attribution = runbook output record (jobId, keyId) joined to the Azure Activity Log entry for the job start (caller identity). A `RequestedBy` runbook parameter would be spoofable; do not rely on one.
- Anything that can edit the Automation account's runbooks or role assignments is effectively as privileged as the app. Note this in the README's security section.
- Not all Teams cmdlets support app-only auth.

## Acceptance criteria
- Both runbooks and the client module exist, pass PSScriptAnalyzer and Pester.
- `terraform validate` and plan succeed with placeholder variables.
- `MANUAL-STEPS.md` and README written, including limitations and the Places workaround.
- No hardcoded IDs or credential material anywhere in the diff.

## Open questions — answer from the repo if possible, otherwise list them in your summary
1. Earlier planning assumed a user-assigned identity for runbooks; this design uses the Automation account's SAMI. Does the repo's Terraform already define one or the other?
2. Is there an existing module pattern for runbooks that these two should follow?
3. Is approval required before checkout (e.g. PIM for Groups on the checkout group), or is group membership sufficient?
4. Should runbook output records be forwarded to Log Analytics / Sentinel? If diagnostic settings exist in the repo, wire them; otherwise note it.
5. Maximum checkout duration: fixed at 60 minutes, or a bounded parameter?

## Deliverable summary expected from you
Files added/changed, test results, plan output summary, conflicts with repo conventions, and answers or remaining questions from the list above.
