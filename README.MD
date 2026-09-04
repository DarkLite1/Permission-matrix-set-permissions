# 🛡️ Permission Matrix: Set Permissions

Welcome to **Permission Matrix**, a robust PowerShell tool designed to automate, apply, verify, and manage NTFS and SMB permissions at scale.

By reading a centralized Excel-based matrix and a simple JSON configuration file, this script takes the headache out of complex folder security, ensuring your environment stays secure and compliant across multiple remote computers.

## 🌟 Features

- **Action Modes**: Choose between `Check`, `Fix`, and `New` actions to either safely audit or actively apply security settings.
- **Multi-Threaded Performance**: High-speed execution through concurrent jobs across multiple computers and folders.
- **Comprehensive Reporting**: Automatically generates HTML overviews and detailed reports of your current permission state and any changes made.
- **Automated Notifications**: Built-in SMTP support (via MailKit/MimeKit) keeps your team updated via email.
- **ServiceNow & SharePoint Integration**: Syncs with ServiceNow for role management and optionally uploads HTML reports directly to a SharePoint Online document library.
- **Automated Access Reviews**: A standalone, read-only audit tool e-mails folder owners their current access for periodic sign-off—without changing a single permission.
- **Centralized Log Analysis**: Consolidate past run logs into a single, sortable Excel workbook to quickly spot errors, incorrect permissions, corrections, and processing bottlenecks.

## 🏷️ Check types

Every finding carries one of five types, which drive the pill, the colour, the counters and the mail subject.

| Type | Meaning | Colour |
| --- | --- | --- |
| `FatalError` | The run could not do its job. | Red |
| `Incorrect` | Disk does not match the matrix and nothing was changed (`Action` is `Check`). | Orange |
| `Warning` | An operational problem that is not permission drift, such as an ACL that could not be read. | Amber |
| `Fixed` | Drift that this run corrected (`Action` is `Fix` or `New`). | Dark green, deliberately darker than a clean matrix |
| `Information` | A notice only. Never colours a card and never promotes it in the overview. | Grey |

`Incorrect` and `Fixed` are the same finding under a different action, so the check `Name` is identical in both. That name is the grouping key of the issue report, which is also why runs from before 2026-09 - where drift was reported as a `Warning` - keep counting `Incorrect` together with `Warning` in its trend sheets.

## 🧩 Entrypoints

Everything begins in the **[`Scripts/Entrypoints`](Scripts/Entrypoints)**. There are three independent scripts you can run depending on your goal.

### 1. `PermissionMatrix.ps1` (The Main Engine)

- **What it does**: This is the core script. It reads your matrices and applies or checks your NTFS and SMB permissions. (This is the only script that actually modifies an ACL).
- **Driven by**: `-ConfigurationJsonFile`
- **Typical schedule**: Every few minutes, or nightly.

### 2. `PermissionMatrixAuditReport.ps1` (Access Review)

- **What it does**: A read-only audit tool. It e-mails each folder's designated owner their current access for sign-off, ensuring periodic governance.
- **Driven by**: `-ConfigurationJsonFile` (uses its own audit configuration file).
- **Typical schedule**: Monthly.

### 3. `PermissionMatrixIssueReport.ps1` (Log Analysis)

- **What it does**: A reporting tool that reads the log folders of previous runs and builds a single Excel overview of errors, incorrect permissions, corrections, and durations. Needs no configuration file and doesn't load the main module.
- **Driven by**: `-LogRoot` and `-OutputFile`
- **Typical schedule**: Ad hoc, or weekly.

**Technical Note**: The first two scripts are thin wrappers: they load the `PermissionMatrix` module and hand over to its public orchestrator. All behavior is configured in JSON rather than command-line parameters. All three scripts will surface a fatal error as a non-zero exit code, which smoothly hands off to Windows Task Scheduler.

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7.x+**: Built for modern PowerShell.
- **ImportExcel**: Required for .xlsx logging and reading the matrix.
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **Emailing Prerequisites**: Required for SMTP functionality.
  ```powershell
  Install-Package -Name 'MailKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```
- **ActiveDirectory (RSAT)**: Required by the audit report, which resolves folder owners to mailboxes via `Get-ADObject`, `Get-ADGroupMember` and `Get-ADUser`. The main permission run does not need it: it queries the directory through `System.DirectoryServices` instead.
  ```powershell
  Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'
  ```
- **Microsoft.Graph.Authentication (Optional)**: Only required if `SharePoint.SiteUrl` is set for SharePoint uploads.
  ```powershell
  Install-Module -Name 'Microsoft.Graph.Authentication' -Scope 'AllUsers'
  ```
  The upload authenticates as an Entra ID app registration using a certificate. The app needs an admin-consented `Sites.ReadWrite.All` role, or a `Sites.Selected` grant with the **write** role on the specific target site.
- **Permissions**: You will need Administrative rights on the target file systems and remote management permissions.
- **WinRM capacity**: Each concurrent job holds one WinRM shell on the remote computer. Check your `MaxShellsPerUser` limit on busy file servers before raising concurrency:
  ```powershell
  winrm get winrm/config/winrs
  ```

### Usage

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/DarkLite1/Permission-matrix-set-permissions.git
    cd Permission-matrix-set-permissions
    ```
2.  **Set up your files**: Review the **[`Examples`](Examples)** folder. Create your own JSON configuration file, your defaults file (`Defaults.xlsx`), and your matrix files (`Matrix.xlsx`). Point your configuration file to these new assets.

    Then point `Matrix.FolderPath` and `Matrix.DefaultsFile` in your configuration at the matrix folder and defaults file you created.

3.  **Run the entrypoint**:
    ```powershell
    pwsh.exe -Command "& '.\Scripts\Entrypoints\PermissionMatrix.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit `$LASTEXITCODE"
    ```

💡 **TIP**: Use the `Detailed` logging setting in the JSON configuration during initial testing to see granular ACL comparisons (current vs. expected) before running a `Fix` action.

## ⚙️ Configuration

The script is driven by a single JSON configuration file. There's no complex schema to memorize—simply copy **[`Example.json`](Examples/Example.json)** and adjust the values. It is completely self-documenting, with every option detailed inline.

💡 **TIP**: Most fields accept an environment-variable reference like `"ENV:SMTP_SERVER"`. You never need to hardcode secrets or credentials in the JSON file!

### Concurrency

Execution speed is heavily influenced by concurrency. Three settings control your limits:

| Setting            | Caps                                                                                                                      |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------- |
| `JobsTotal`        | Jobs running at the same time across **all** remote computers.                                                            |
| `JobsPerComputer`  | Jobs running at the same time on any **single** remote computer. (Keep this under the computer's MaxShellsPerUser limit). |
| `FoldersPerMatrix` | Folders processed at the same time **within one** job. Also caps how many matrix files are read in parallel at startup.   |

#### How to tune for performance:

1. Set `JobsPerComputer` based on your remote computers' `MaxShellsPerUser`, leaving a little headroom.

2. Set `JobsTotal` to roughly `JobsPerComputer` × the number of computers you want active at once.

3. If one long-running matrix file dominates the queue, raise `FoldersPerMatrix` to parallelize the folders inside it. Note that this setting does double duty: it also throttles how many matrix files are read from disk at startup, so raising it speeds up that phase too.

## 📄 Matrix Files

Your per-folder rules live in the Excel matrix files, with shared fallbacks in your `DefaultsFile`.

### Permission Codes

Use these shorthand codes to define access in the matrix:

| Code | Permission   | Description                                                     |
| ---- | ------------ | --------------------------------------------------------------- |
| L    | List         | View the contents of the folder.                                |
| R    | Read         | Read files and folder attributes.                               |
| W    | Write        | Create and modify files.                                        |
| F    | Full Control | All permissions, including changing permissions/ownership.      |
| I    | Ignore       | Skips this entry without making changes (valid in Matrix only). |

### Default Permissions

Each matrix's `Settings` sheet has an `ApplyDefaultPermissions` column (`True`/`False`):

- **True**: Merges the global fallback permissions from your `DefaultsFile` (e.g., automatically granting "IT Staff" access). **Note: Conflicts between the matrix and defaults will halt the script.**

- **False**: Ignores global defaults entirely. Security is managed strictly by the explicit entries in that specific matrix (ideal for highly sensitive HR or Finance folders).

## ⏱️ Scheduled Task Example

For automated, recurring executions via Windows Task Scheduler, use this setup:

- Program: `pwsh.exe`
- Arguments:

```
-Command "& 'C:\Tools\Permission-matrix-set-permissions\Scripts\Entrypoints\PermissionMatrix.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

## 🔍 Access Review (Audit Report)

Close the loop on access governance with `PermissionMatrixAuditReport.ps1`.

This tool e-mails the responsible person (found in the `MatrixResponsible` field) a copy of that folder's current access as an Excel attachment for review. It uses a separate configuration file (see [`ExampleAuditReport.json`](Examples/ExampleAuditReport.json).

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\PermissionMatrixAuditReport.ps1' -ConfigurationJsonFile 'C:\MyAuditConfig.json'; exit `$LASTEXITCODE"
```

💡 **TIP**: To safely test the review emails without spamming actual owners, set `Settings.SendMail.To` in your audit configuration. All emails will route to this test address instead.

## 📊 Logging & Monitoring

The script provides three layers of feedback (configured under `Settings`):

1. **File Logs**: Saved locally to your specified folder.
2. **Event Log**: Written to the Windows Event Log (e.g., `Scripts`) for system-level monitoring.
3. **Email Reports**: Summaries sent via SMTP.

### Diagnostics

Every run also writes volume and cost counters, so that a change in run time can be attributed rather than guessed at. Three files land in the dated run folder:

| File                       | Grain                | Answers                                                     |
| -------------------------- | -------------------- | ----------------------------------------------------------- |
| `Diagnostics.json`         | One row per Settings row | Which matrix got slower.                                 |
| `Diagnostics.Paths.json`   | One row per matrix folder | Where inside it. Joins to the above on `ID`.            |
| `Diagnostics.html`         | Both                 | Sortable page for eyeballing and comparing two runs.         |
| `Diagnostics.Fields.json`  | —                    | What every counter means, and the caveats.                   |

Each matrix folder also gets an `ID <guid> - Diagnostics.json` holding that one Settings row in full, linked from a **Diagnostics** chip on the execution report card (with a **?** next to it pointing at the field reference).

#### How to read them:

A duration on its own cannot tell you why a job got slower. Read it together with `ItemsWalked` (how much data there was) and `AclReadMsPerItem` (what each storage operation cost):

- **Duration up, `ItemsWalked` up in proportion** → the share grew. Expected, nothing to fix.
- **Duration up, `ItemsWalked` flat, `AclReadMsPerItem` up** → each operation got more expensive. Look at the file server (backup, anti-virus, deduplication, snapshot pressure), not at the matrix.
- **`AceCountMean` climbing run over run** → permissions are being appended rather than replaced. This compounds every night and is worth fixing before anything else.
- **`IncorrectItems` the same non-zero value every run** → the tree never converges; the same items are corrected every night.

To trend one path across every run:

```powershell
Get-ChildItem '\\SERVER\Log\Permission matrix\*\Diagnostics.json' | ForEach-Object {
    $run = $_.Directory.Name
    Get-Content $_ -Raw | ConvertFrom-Json |
    Where-Object Path -eq 'E:\DEPARTMENTS\STAFF\SCM' |
    Select-Object @{n='Run';e={$run}}, DurationSeconds, ItemsWalked, AclReadMsPerItem, AceCountMean
}
```

If a Settings row got slower but you cannot see why, filter `Diagnostics.Paths.json` on that row's `ID`. Its rows are already sorted by measured cost, so the folder responsible is at the top.

For eyeballing rather than scripting, open `Diagnostics.html` from two run folders side by side. Click any column to sort, type in the box to filter. Rows start in a stable order (computer, then path) so the two windows line up row for row — sort by cost only once you know which row you are chasing, since cost order differs between nights. The page embeds its own data, so it works from a log share with no network and needs no other file in the folder.

⚠️ **NOTE**: `AclReadMsPerItem` is a sampled mean, and carries a few percent of run-to-run noise even on identical work. Check `AclReadSamples` and `AclReadBasis` before trusting it. A single run moving 10% is not a signal; a trend across several runs is.

## 📈 Issue & Performance Report

To look across all runs at once, use `PermissionMatrixIssueReport.ps1`. Point it at your log folder to produce a master Excel workbook that answers what went wrong, how often, where, and how long it took.

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\PermissionMatrixIssueReport.ps1' -LogRoot '\\SERVER\Log\Permission matrix' -OutputFile 'C:\Temp\Matrix issues.xlsx'"
```

#### Key Parameters:

- `-LogRoot`: (Mandatory) The folder holding the run folders.
- `-OutputFile`: Path of the `.xlsx` to write. (Omit to simply print counts to the console).
- `-Parallel`: Reads several run folders at once (great for network shares).

The generated workbook contains fully filterable tables, hyperlinked logs, and detailed performance metrics breaking down processing time versus elapsed time.

## 🙏 Acknowledgements

We extend our sincere gratitude to the developers and maintainers of the following projects that make this tool possible:

- [MailKit](https://github.com/jstedfast/MailKit) & [MimeKit](https://mimekit.net/): For powerful and flexible email sending functionality.
- [ImportExcel](https://github.com/dfinke/ImportExcel): For seamless Excel data manipulation in PowerShell.

## 🤝 Contributing & Support

- **Contributing**: We welcome your ideas! Feel free to open issues or submit pull requests with suggestions or bug fixes.
- **Support**: If you encounter any problems, please open an issue on the GitHub repository.

## 📜 License

This software is provided for your use, free of charge, for both private and professional purposes.

**NO GUARANTEES ARE GIVEN**: This software is provided "as is," without any warranty of any kind.

**YOUR RESPONSIBILITY**: You are solely responsible for the code you deploy and its consequences. In no event shall the author be liable for any claim or damages arising from the use of this software.
