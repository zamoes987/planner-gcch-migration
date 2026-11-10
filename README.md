# Migrate-PlannerToGccHigh.ps1

One PowerShell script that moves **Microsoft Planner** plans (basic and premium / Project Plan 3) from a
Microsoft 365 **Commercial** tenant to a Microsoft 365 **GCC High** tenant.

## How to run it: two commands

On a Windows machine with PowerShell 7.4+, in this folder:

```powershell
.\Migrate-PlannerToGccHigh.ps1 -Step Export    # commercial side
.\Migrate-PlannerToGccHigh.ps1 -Step Import    # GCC High side
```

Run both from **the same Windows account on the same machine**. The app certificate the script creates can't be moved to another machine.

There's nothing to edit. Browser sign-ins pop up when needed:

| Command | Sign-in | When |
|---|---|---|
| Export | Commercial **admin** | First run only. Registers the migration app. |
| Export | A commercial user who is a **member of the groups** | Reads task chat. |
| Export | Commercial account that can read premium plans | Only matters if you have premium plans. |
| Import | GCC High **admin** | First run only. Registers the migration app. |
| Import | GCC High user **with a Project license** | Only if premium plans were exported. |

Both commands are safe to rerun. They resume where they stopped and never duplicate anything.

## What happens automatically

- **Prerequisites:** missing PowerShell modules (`Microsoft.Graph.Authentication`, `Az.Accounts`) are installed for your user. Turn this off with `-NoModuleInstall`.
- **App setup:** an app registration and a non-exportable certificate are created in each tenant, with only these Graph application permissions:
  - Commercial: `Tasks.Read.All`, `Group.Read.All`, `User.Read.All`, `Organization.Read.All`
  - GCC High: `Tasks.ReadWrite.All`, `Group.Read.All`, `Group.Create`, `GroupMember.ReadWrite.All`, `User.Read.All`, `Organization.Read.All`
- **What gets exported:** every Microsoft 365 group that has plans is found, and every plan in it is exported.
- **Users are matched** by the same alias on the GCC High domains. Example: `jane@contoso.com` → `jane@contoso.us`, using both tenants' verified domains.
- **Groups are matched** by alias, then by name. **Groups that don't exist in GCC High are created** with the same name, alias, privacy and owners/members.
- **Assignees** who aren't members of their GCC High group are added to it, so their assignments work.
- **SharePoint links** are rewritten to the GCC High host. This is a best guess from the tenant names; see §4.
- **Premium plans:** Dataverse environments are found automatically.
- **The order of writes in GCC High:**
  1. Import previews the mapping without writing anything.
  2. It runs a throwaway write test (plan, bucket, task, details, then delete) in an **existing** matched group.
  3. Only if that passes does it create missing groups and import.
  4. If *no* group matched at all, it has to create the groups first so there's somewhere to test.
- **After importing,** Import checks every plan and task and writes reports to `logs\`.
- **Export ends with a summary** of what was exported and what was SKIPPED, with the reason and the fix. Import repeats any skipped premium export loudly.

**Import changes your GCC High tenant by default:** it creates missing groups and adds assignees as group members. To see exactly what would happen first, run `.\Migrate-PlannerToGccHigh.ps1 -Step DryRun` after Export, then open `logs\group-mapping.csv`, `logs\user-mapping.csv` and `logs\basic-import-problems.csv`. DryRun creates no groups, plans or memberships. On a machine's first run it does register the GCC High migration app (the same admin sign-in Import would ask for).

## What migrates

| | Basic plans | Premium plans |
|---|---|---|
| Plans, buckets (order kept), tasks, dates, priority, % complete | ✅ | ✅ including subtasks, effort, dependencies |
| Labels | ✅ | ❌ The API can't create them [S4]; listed in `export\premium\labels-manual.csv` |
| Assignments | ✅ for every matched user | ⚠️ Generic team member per person (`-TryNamedResources` uses the real user if possible) |
| Description, checklist, links | ✅ | ⚠️ Checklists best-effort |
| Comments and task chat | ⚠️ **As text** at the end of the task description, with author and time | ❌ |
| Attachment files | ⚠️ Only the link moves. Files must be migrated separately | same |
| Created by / completed by and their dates | ❌ Set by the service to the migration app and the import time [S11] | ❌ |
| Card order inside a bucket, task history, Teams tabs | ❌ | ❌ |
| Plans **not owned by a Microsoft 365 group** (e.g. personal plans) | ❌ Not found by discovery, which lists group plans only | n/a |

**Project Online / PWA can't be migrated.** It retired on Sept 30, 2026, and its data is no longer accessible [S1][S2].

**Task chat:** it replaced comments for basic plans in early 2026 [S13]. It can only be read through a Microsoft Graph **beta** API, and Microsoft doesn't support beta in production [S14]. If that step fails, Export logs a warning and continues; older comments are still exported.

**Premium plans in GCC High:** Microsoft's own pages disagree on whether they're available.
- Planner FAQ: "not available" [S23].
- GCC High launch post: available with a premium license [S24].

Import tests it in your tenant. If the premium test fails, basic plans are still imported, premium plans are skipped, and their data stays in `export\premium\`.

## After Import: check these files

| File | What to do |
|---|---|
| `logs\user-mapping.csv` | Rows with `NOT FOUND`: add them to `config\users.csv` (`sourceUpn,destUpn`) and rerun Import. |
| `logs\group-mapping.csv` | Each group shows `same alias`, `same name` or `created`. To force a different group, put its ID in `config\groups.csv` (`destGroupId`) and rerun. |
| `logs\verify-basic.csv`, `logs\verify-premium.csv` | Every FAIL row needs a look. |
| `logs\*problems.csv` | Anything that didn't carry. |

**Manual steps:**
- Share each premium plan with its group in Planner.
- Re-apply premium labels.
- Re-add Planner tabs in Teams.
- Delete the `PREFLIGHT … delete me` premium plan.

## 4. Optional overrides (`config\`)

- `users.csv`: `sourceUpn,destUpn` for anyone whose alias differs.
- `groups.csv`: written by Export. Put a GCC High group ID in `destGroupId` to choose the target yourself.
- `urlmap.csv`: `sourcePrefix,destPrefix` if your GCC High SharePoint address isn't `https://<tenant>.sharepoint.us/`. The automatic mapping is printed in the Import log.
- `settings.json`, created automatically:
  - `CreateMissingGroups: false` turns off group creation.
  - `AddMissingAssigneesToGroups: false` stops adding assignees to groups.
  - `DataverseUrl` pins a Dataverse environment.
  - `DataverseAuth: DeviceCode` with `DataverseClientId` is the fallback if Az sign-in is blocked.

## Pre-mortem: the 3 likeliest ways this goes wrong

1. **People keep working in the commercial tenant after Export.** Their changes never reach GCC High. Freeze the source before Export.
2. **Users think comments or authorship were lost.** They're text in the description, and the migration app shows as author. Tell users before cutover.
3. **Files weren't migrated before Import**, so the rewritten links point to nothing. Move the files first, or rerun Import later; it only fills in what's missing.

## How it was tested

The real script runs offline against simulated versions of both clouds:
- **44 unit tests**, all passing.
- **27 end-to-end checks** of the two-command flow, all passing, with no config files:
  - finding groups that have plans, plus the export summary
  - exporting task chat
  - a dry run on a fresh machine: registers the GCC High app, writes no Planner data
  - the write test running before any group is created
  - matching users by domain
  - matching one group by alias and creating another
  - adding an assignee after a forced replication delay
  - creating a plan in a new group after a forced not-ready response
  - a forced 429 retry
  - a forced 412 that must not lose comments
  - an idempotent rerun
  - verification, and idempotent Setup
- **20 premium checks**, all passing:
  - finding environments
  - parent-before-child ordering across operation sets
  - deleting the default bucket
  - an idempotent rerun with no extra sign-in
  - the device-code fallback

Microsoft's live services can only be tested in your tenants. That's what Import's built-in test does before it writes anything.

## Sources (checked 2026-11-09)

- [S1] Prepare for Project Online retirement (Learn): https://learn.microsoft.com/en-us/planner/prepare-for-project-online-retirement
- [S2] Microsoft Project Online is retiring (Microsoft Planner Blog, updated 2026-08-24): https://techcommunity.microsoft.com/blog/plannerblog/microsoft-project-online-is-retiring-what-you-need-to-know/4450558
- [S4] Use Project schedule APIs (license requirement, operation set limits, labels not creatable, default bucket): https://learn.microsoft.com/en-us/dynamics365/project-operations/project-management/schedule-api-preview
- [S11] plannerTask resource (read-only fields): https://learn.microsoft.com/en-us/graph/api/resources/plannertask?view=graph-rest-1.0
- [S13] MC1193421, task chat replaces comments, early 2026 (mirror): https://mc.merill.net/message/MC1193421
- [S14] List task chat messages (beta, Global only, delegated): https://learn.microsoft.com/en-us/graph/api/plannertask-list-messages?view=graph-rest-beta
- [S23] Planner FAQ (support.microsoft.com, undated): https://support.microsoft.com/en-us/planner/frequently-asked-questions-about-microsoft-planner
- [S24] The new Planner for the web is now available in GCC High (Microsoft Planner Blog, updated 2025-12-05): https://techcommunity.microsoft.com/blog/plannerblog/the-new-planner-for-the-web-is-now-available-to-customers-in-gcc-high/4443618
- Graph national clouds (graph.microsoft.us, login.microsoftonline.us): https://learn.microsoft.com/en-us/graph/deployments
- Create plannerPlan / plannerTask (US Gov L4, app-only): https://learn.microsoft.com/en-us/graph/api/planner-post-plans?view=graph-rest-1.0 · https://learn.microsoft.com/en-us/graph/api/planner-post-tasks?view=graph-rest-1.0
- Create group (Group.Create, ≤20 owners+members, owners required): https://learn.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0
- Add members (GroupMember.ReadWrite.All, replication-delay retry): https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0
- Get organization (verifiedDomains): https://learn.microsoft.com/en-us/graph/api/organization-get?view=graph-rest-1.0
- Grant permissions programmatically (appRoleAssignedTo): https://learn.microsoft.com/en-us/graph/permissions-grant-via-msgraph
- Self-signed certificate: https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate
- Dataverse Global Discovery Service (globaldisco.crm.microsoftdynamics.us): https://learn.microsoft.com/en-us/power-apps/developer/data-platform/discovery-service
- Dataverse Web API with PowerShell (Az token, 429/Retry-After): https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-ps-and-vscode-web-api
- OAuth with Dataverse (user_impersonation): https://learn.microsoft.com/en-us/power-apps/developer/data-platform/authenticate-oauth · Device code flow: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code
- Planner references, order hints, task details: https://learn.microsoft.com/en-us/graph/api/resources/plannerexternalreferences?view=graph-rest-1.0 · https://learn.microsoft.com/en-us/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0 · https://learn.microsoft.com/en-us/graph/api/plannertaskdetails-update?view=graph-rest-1.0
- Graph throttling: https://learn.microsoft.com/en-us/graph/throttling
