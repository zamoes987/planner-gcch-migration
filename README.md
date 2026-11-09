# Migrate-PlannerToGccHigh.ps1

**One PowerShell script** that migrates Microsoft Planner basic plans and Planner premium plans
(Project Plan 3, formerly Project for the web) from a Microsoft 365 **Commercial** tenant to a Microsoft 365 **GCC High** tenant.
You only need `Migrate-PlannerToGccHigh.ps1`. It creates its own `config\`, `export\`, `state\` and `logs\` folders.
`tests\` is developer-only.

Every fact below is tied to a source ([S#], all checked 2026-11-09). Microsoft's own pages are used wherever one exists.
Anything the docs can't settle for *your* tenants is tested by `-Step Preflight`, not assumed.

---

## 1. What exists to migrate (Nov 2026)

| Data | Status | Handled by |
|---|---|---|
| **Project Online / PWA** | **Retired Sept 30, 2026.** "Projects and associated Project Online service data are no longer accessible through Project Online" [S1]. Microsoft's announcement says the same and lists "Planner – basic and premium capabilities" as unaffected [S2]. Message Center MC1478955 lists Worldwide, GCC, GCC High and DoD for that date [S3]. | Nothing to read via API. Pre-cutoff `.mpp` files only (§7). |
| **Premium plans** | Stored in **Dataverse**, written through the Project schedule APIs [S4]. **Microsoft's own sources conflict on whether they're available in GCC High** (see below). | `ExportPremium` → `ImportPremium`, **only if Preflight's premium checks PASS** |
| **Basic plans** | Graph Planner API: "Yes" for GCC-High in the Planner service description (updated Oct 15, 2026) [S7]. The create APIs are listed for US Gov L4 [S8][S9]. | `Export` → `ExportTaskChat` → `Import` → `Verify` |

**Premium plans in GCC High: unresolved documentation conflict (checked 2026-11-09)**

| Microsoft source | What it says |
|---|---|
| Planner FAQ, support.microsoft.com, undated [S23] | "premium capabilities are currently not available for customers in Government Cloud Communities (GCC), GCC High, and Department of Defense (DoD) tenants". Elsewhere on the same page: "*some* premium capabilities are currently not available". |
| "The new Planner for the web is now available to customers in GCC High", Microsoft Planner Blog, Aug 14, 2025, updated Dec 5, 2025 [S24] | "If you already have a Planner and Project premium license, you're already equipped with access to the new Planner's premium capabilities". Lists 7 GCC High exclusions: Copilot/Project Manager Agent, premium tasks in My Tasks, basic→premium conversion, premium task history, Portfolios, plan backgrounds, Task Publishing. |
| Teams GCC High rollout post, Microsoft Planner Blog, updated Dec 2025 [S5] | "premium plans are already in the new Planner" |
| Planner for admins, Learn, Oct 28, 2026 [S6] | No government exclusion either way |

Supporting, but **secondary**: a third-party repost of the Microsoft 365 roadmap lists ID 498496, raising the premium-plan task limit to 1,500, with GCC High among its clouds [S25]. Check that ID on the official roadmap before relying on it.

No document settles this. Your tenant does: `-Step Preflight` creates a test premium plan through `msdyn_CreateProjectV1` in your GCC High environment.
- **PASS:** run the premium steps.
- **FAIL:** premium plans can't be recreated in GCC High as premium plans. Keep `export\premium\` as the record. Options are rebuilding them as basic plans by hand, or a Microsoft support case.

Regular (basic) plans are unaffected either way.

## 2. What moves, and what can't (no tool can)

| Item | Basic plans | Premium plans |
|---|---|---|
| Plans, buckets (order kept), tasks, dates, priority, % complete, labels on tasks | ✅ | ✅ project, buckets, tasks (hierarchy, scheduled dates, effort), dependencies. ❌ **labels**: the API can't create them [S4]. They go to `export\premium\labels-manual.csv` |
| Assignments | ✅ for mapped users (DomainMap / users.csv) | ⚠️ Generic team member by default. `-TryNamedResources` uses the real user if they already have a bookable resource in GCC High (preflight checks) |
| Description, checklist, links | ✅ Links rewritten via `urlmap.csv`, encoded per Graph rules [S10] | ⚠️ Checklists best-effort, using metadata-discovered columns |
| **Old comments** (group conversation) | ⚠️ **As text** appended to the description (author + time) [S11][S12] | ❌ |
| **Task chat** (replaced comments for basic plans in early 2026 [S13]) | ⚠️ **As text.** It can only be *read* through Graph **beta**, Global cloud, signed-in user only [S14]. Writing chat is also Global-only, so chat can't be recreated as chat in GCC High [S15]. | ❌ |
| Created/completed by + dates | ❌ read-only; the service stamps the migration identity/time [S11] | ❌ |

**About task chat (`-Step ExportTaskChat`):**
- It's a one-time read through a **beta** API. Microsoft states beta APIs aren't supported for production use [S14].
- If that step fails or the API changes, task chat isn't migrated. Old comments still are.
- The signed-in account must be a **member of each source group**. Tasks it can't read are listed in `logs\taskchat-export-problems.csv`, so an admin may need temporary membership in the groups in `groups.csv`.
| Attachments (files) | ⚠️ Link only. Files must be migrated separately; `urlmap.csv` repoints the links | same |
| Card order inside a bucket, task history, Teams tabs | ❌ | ❌ (no premium task history in GCC High [S5]) |

## 3. The run (one script, step by step)

```powershell
.\Migrate-PlannerToGccHigh.ps1 -Step Init                     # creates config\settings.json, users.csv, urlmap.csv
#   edit config\settings.json: both TenantIds, DomainMap (e.g. contoso.com -> contoso.us),
#   PreflightTestGroupId (a TEST group in GCC High), DataverseUrl x2 (premium only)
.\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Source       # commercial admin: app + cert + permissions, writes IDs into settings
.\Migrate-PlannerToGccHigh.ps1 -Step Setup -Side Destination  # GCC High admin: same, in graph.microsoft.us
.\Migrate-PlannerToGccHigh.ps1 -Step Discover                 # -> config\groups.csv; fill in destGroupId per row
.\Migrate-PlannerToGccHigh.ps1 -Step Preflight                # GO / NO-GO. Fix every FAIL.
.\Migrate-PlannerToGccHigh.ps1 -Step Export                   # announce source freeze first; this JSON is also your backup
.\Migrate-PlannerToGccHigh.ps1 -Step ExportTaskChat           # sign in as a user who is a member of the source groups
.\Migrate-PlannerToGccHigh.ps1 -Step DryRun                   # resolves users/groups; writes nothing; read logs\basic-import-problems.csv
.\Migrate-PlannerToGccHigh.ps1 -Step Import -OnlyIds <planId> # pilot ONE plan; have its owner check it
.\Migrate-PlannerToGccHigh.ps1 -Step Import                   # everything; rerun after any interruption (resumes from state\)
.\Migrate-PlannerToGccHigh.ps1 -Step Verify                   # -> logs\verify-basic.csv
.\Migrate-PlannerToGccHigh.ps1 -Step ExportPremium            # premium: sign in as an account that can read all projects
.\Migrate-PlannerToGccHigh.ps1 -Step ImportPremium            # premium: sign in as a GCC High USER with a Project license
```

**One-day timeline:**

| Time | Steps |
|---|---|
| Hour 0–1 | Init, Setup ×2, Discover |
| Hour 1.5 | Preflight |
| Hour 2 | Export, ExportTaskChat, ExportPremium |
| Hour 2.5 | DryRun |
| Hour 3 | Pilot import of one plan |
| Hour 3.5–6 | Import, ImportPremium |
| Hour 6–7 | Verify, then the problem CSVs and manual items |
| Hour 7–8 | Communicate to users |

**Requirements:**
- PowerShell 7.4+ on Windows. Setup uses `New-SelfSignedCertificate` [S16].
- Modules: `Install-Module Microsoft.Graph.Authentication, Az.Accounts -Scope CurrentUser`. Az 11.1.0+ per Learn [S17].
- Destination users, groups and (for links) files already in GCC High.

**What Setup does:**
- Creates a self-signed certificate with `New-SelfSignedCertificate` [S16]. Unlike Microsoft's recipe, it uses `-KeyExportPolicy NonExportable`, so the private key never leaves this Windows profile.
  **Consequence: run every step from the same Windows account on the same machine.** Otherwise the certificate isn't there and sign-in fails.
- Registers the app with that certificate [S18].
- Looks up Microsoft Graph's permission IDs **at runtime** from the Graph service principal and grants them through `appRoleAssignedTo` [S19]:
  - Source: `Tasks.Read.All`, `Group.Read.All`, `User.Read.All`
  - Destination: `Tasks.ReadWrite.All`, `Group.Read.All`, `User.Read.All`
- Needs an admin who may grant Graph application permissions.

**Premium plans and the Project schedule APIs:** they only accept a **user with a Project license**. Application, system and integration users are refused [S4].

**If `Az.Accounts` can't get a GCC High Dataverse token:**
1. Set `Destination.DataverseAuth` to `DeviceCode`.
2. Set `Destination.DataverseClientId` to a public-client app with the Dynamics CRM *user_impersonation* permission [S20].
3. The script then signs in at `login.microsoftonline.us` [S21] using the documented device-code flow [S22].

If your tenant blocks device-code sign-in, Preflight's sign-in check will FAIL.

## 4. Pre-mortem: it's six months from now and this failed. Most likely reasons, in order

1. **Nobody ran Preflight or the pilot; they went straight to a full Import.**
   - A permission or consent gap in GCC High surfaced as hundreds of failures mid-run, under deadline.
   - Prevention: Preflight + one-plan pilot are mandatory; never skip them.
2. **Users found their comments "gone" and their tasks "re-authored by an app".**
   - Comments and task chat arrived as text in the description. Created-by and completed-by show the migration app.
   - This wasn't communicated, and trust in the migrated data collapsed.
   - Prevention: send users §2's table *before* cutover.
3. **People kept working in the commercial source after Export.**
   - Their changes never reached GCC High, and links pointed to files that hadn't moved.
   - Prevention: announce and enforce a source freeze before Export, and migrate files before running Import with `urlmap.csv`.

Close behind: premium plans built in GCC High but never shared with their groups, so nobody could see them (§6), and expired app certificates (`-CertMonths`, default 3).

## 5. If Preflight fails

| FAIL | Fix |
|---|---|
| Basic create/write | Rerun Setup for that side, wait a few minutes, rerun Preflight. Check `PreflightTestGroupId` is a Microsoft 365 group. |
| Premium `msdyn_CreateProjectV1` | Signed-in user lacks a Project license, or premium plans aren't provisioned in that environment (create one premium plan in Planner once). Basic plans still ship. |
| Premium sign-in | Switch to `DataverseAuth = DeviceCode` (§3). |

## 6. Manual follow-ups

- Share each premium plan with its group in Planner.
- Re-apply labels from `labels-manual.csv`.
- Replace generic team members if needed.
- Delete the empty default bucket that `msdyn_CreateProjectV1` creates [S4], and the `PREFLIGHT … delete me` plan.
- Re-add Planner tabs in Teams.
- Comments that were too long: `export\comments-not-imported\`.

## 7. Outside the script

- **`.mpp` files saved before the Project Online cut-off:** import them manually in Planner. That UI's availability in GCC High is **not verified**; test one file first.

## 8. How it was tested, and what remains unproven

**Offline, all passing:**
- 40 unit tests.
- 25 end-to-end checks for basic plans and 17 for premium plans. They run the real script against in-memory fakes of both clouds:
  - Setup, including a forced replication retry
  - task chat export
  - a forced 429 retry
  - a forced 412 that must not lose comments
  - an idempotent rerun
  - parent-before-child ordering across operation sets
  - device-code sign-in against the `.us` authority

**Not provable without your tenants:**
- That GCC High accepts app-only Planner writes.
- That premium plans and the Project schedule APIs are provisioned in *your* GCC High environment.
- Which Dataverse sign-in your Conditional Access allows.

`-Step Preflight` tests each of these in minutes, with throwaway objects, before any real data moves.

## Sources

- [S1] Prepare for Project Online retirement (Microsoft Learn, updated 2026-08-24): https://learn.microsoft.com/en-us/planner/prepare-for-project-online-retirement
- [S2] Microsoft Project Online is retiring: what you need to know (Microsoft Planner Blog, G. Bullock, updated 2026-08-24): https://techcommunity.microsoft.com/blog/plannerblog/microsoft-project-online-is-retiring-what-you-need-to-know/4450558
- [S3] MC1478955 "Final reminder: Microsoft Project Online retires on September 30, 2026" (2026-09-25). Verify in your Message Center; mirror: https://app.cloudscout.one/evergreen-item/mc1478955/
- [S4] Use Project schedule APIs (Learn, updated 2026-07-14): https://learn.microsoft.com/en-us/dynamics365/project-operations/project-management/schedule-api-preview. Web API payloads: https://learn.microsoft.com/en-us/dynamics365/project-operations/project-management/scheduling-apis-powerautomate
- [S5] The new Planner app in Teams is now available in GCC High (Microsoft Planner Blog, Oct 2025, updated Dec 2025): https://techcommunity.microsoft.com/blog/plannerblog/the-new-planner-app-in-microsoft-teams-is-now-available-to-customers-in-gcc-high/4458420
- [S6] Microsoft Planner for admins (Learn, 2026-10-28): https://learn.microsoft.com/en-us/planner/planner-for-admins
- [S7] Microsoft Planner service description (Learn, updated 2026-10-15): https://learn.microsoft.com/en-us/office365/servicedescriptions/project-online-service-description/microsoft-planner-service-description
- [S8] Create plannerTask (US Gov L4 ✅; Tasks.ReadWrite.All app-only): https://learn.microsoft.com/en-us/graph/api/planner-post-tasks?view=graph-rest-1.0
- [S9] Create plannerPlan: https://learn.microsoft.com/en-us/graph/api/planner-post-plans?view=graph-rest-1.0 · List group plans: https://learn.microsoft.com/en-us/graph/api/plannergroup-list-plans?view=graph-rest-1.0
- [S10] plannerExternalReferences: https://learn.microsoft.com/en-us/graph/api/resources/plannerexternalreferences?view=graph-rest-1.0 · Update task details (If-Match): https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update?view=graph-rest-1.0 · Order hints: https://learn.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0
- [S11] plannerTask resource: https://learn.microsoft.com/en-us/graph/api/resources/plannertask?view=graph-rest-1.0
- [S12] List conversation posts (US Gov L4 ✅): https://learn.microsoft.com/en-us/graph/api/conversationthread-list-posts?view=graph-rest-1.0
- [S13] MC1193421 "Retirement of several Microsoft Planner features in early 2026" (task chat replaces comments; published 2025-12-09, updated 2026-02-19). Verify in your Message Center; mirror: https://mc.merill.net/message/MC1193421
- [S14] List task chat messages (beta; Global only; delegated Tasks.Read only): https://learn.microsoft.com/en-us/graph/api/plannertask-list-messages?view=graph-rest-beta · resource: https://learn.microsoft.com/en-us/graph/api/resources/plannertaskchatmessage?view=graph-rest-beta
- [S15] Create task chat message (beta; Global only): https://learn.microsoft.com/en-us/graph/api/plannertask-post-messages?view=graph-rest-beta
- [S16] Create a self-signed public certificate (Learn): https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate
- [S17] Dataverse Web API with PowerShell (Az token pattern; 429 + Retry-After): https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/quick-start-ps · https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-ps-and-vscode-web-api
- [S18] App-only auth with Graph PowerShell (app + certificate): https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0 · `-Environment USGov`: https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0
- [S19] Grant API permissions programmatically (appRoleAssignedTo): https://learn.microsoft.com/en-us/graph/permissions-grant-via-msgraph
- [S20] OAuth with Dataverse (public clients use `<env>/user_impersonation`): https://learn.microsoft.com/en-us/power-apps/developer/data-platform/authenticate-oauth
- [S21] Microsoft Graph national cloud deployments (login.microsoftonline.us, graph.microsoft.us): https://learn.microsoft.com/en-us/graph/deployments
- [S22] OAuth 2.0 device authorization grant: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code
- [S23] Frequently asked questions about Microsoft Planner (support.microsoft.com, no date shown): https://support.microsoft.com/en-us/planner/frequently-asked-questions-about-microsoft-planner
- [S24] The new Planner for the web is now available to customers in GCC High (Microsoft Planner Blog, Aug 14, 2025, updated Dec 5, 2025): https://techcommunity.microsoft.com/blog/plannerblog/the-new-planner-for-the-web-is-now-available-to-customers-in-gcc-high/4443618
- [S25] (secondary) Planet Technologies repost of Microsoft 365 roadmap updates (Aug 28, 2025) listing ID 498496: https://go-planet.com/in-the-know/august-28-2025/
- Checked but silent on Planner: Office 365 GCC High and DoD service description (updated 2026-08-03): https://learn.microsoft.com/en-us/office365/servicedescriptions/office-365-platform-service-description/office-365-us-government/gcc-high-and-dod
- Also used: Power Apps US Government (`*.crm.microsoftdynamics.us`, admin URLs) https://learn.microsoft.com/en-us/power-platform/admin/powerapps-us-government · Graph throttling https://learn.microsoft.com/en-us/graph/throttling · Dataverse metadata queries https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-metadata-web-api · Azure Government PowerShell https://learn.microsoft.com/en-us/azure/azure-government/documentation-government-get-started-connect-with-ps
