# Permission Matrix: Set Permissions

A robust PowerShell tool designed to **apply, verify, and manage NTFS and SMB permissions** at scale. By reading a centralized Excel-based matrix and a JSON configuration file, the script automates the complex task of ensuring folder security across multiple remote computers.

## 🌟 Features

- **Action Modes**: Supports `Check`, `Fix`, and `New` actions to audit or apply security settings.
- **Multi-Threaded Performance**: Configure concurrent jobs across multiple computers and folders for high-speed execution.
- **Comprehensive Reporting**: Generates HTML overviews and detailed reports of the current permission state and/or changes.
- **Automated Notifications**: Integrated SMTP support (via MailKit/MimeKit) to send status updates by mail.
- **Granular Logging**: Detailed logging to XLSX/JSON and the Windows Event Log.
- **ServiceNow Integration**: Syncs with ServiceNow tables to manage roles and form data.
- **SharePoint Publishing**: Optionally uploads the HTML overview to a SharePoint Online document library after each run, overwriting the previous version in place.
- **Access Reviews**: A separate, read-only audit entrypoint e-mails folder owners their current access for periodic sign-off — without changing any permissions.

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7.x or later**: This script is built for modern PowerShell.
- **ImportExcel**: Required for .xlsx logging and reading the matrix.
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **Emailing Prerequisites**: Required for email functionality. The installed DLL paths are referenced from `Settings.SendMail.AssemblyPath` in your configuration.
  ```powershell
  Install-Package -Name 'MailKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```
- **Microsoft.Graph.Authentication**: Only required when `SharePoint.SiteUrl` is set in the configuration. Without it the run applies permissions normally and reports the failed upload as a system error in the summary mail.
  ```powershell
  Install-Module -Name 'Microsoft.Graph.Authentication' -Scope 'AllUsers'
  ```
  The upload authenticates as an Entra ID app registration using a certificate. The app needs an admin-consented `Sites.ReadWrite.All` role, or a `Sites.Selected` grant with the **write** role on the specific target site.
- **Permissions**: Administrative rights on the target file systems and remote management permissions.
- **WinRM capacity**: Each concurrent job holds one WinRM shell on the remote computer for its lifetime, so `MaxConcurrent.JobsPerComputer` must stay within that computer's `MaxShellsPerUser` limit. Check it on the busy file servers before raising the concurrency settings:
  ```powershell
  winrm get winrm/config/winrs
  ```
  Exceeding the limit surfaces as ordinary-looking connection failures rather than a clear quota error, so it is worth confirming up front. See [Concurrency](#concurrency) below.

### Usage

1.  Clone this repository to your local machine:
    ```bash
    git clone https://github.com/DarkLite1/Permission-matrix-set-permissions.git
    cd Permission-matrix-set-permissions
    ```
2.  Review the samples in the **[`Examples`](Examples)** folder, then create your own:
    - **Configuration file** — based on [`Example.json`](Examples/Example.json).
    - **Defaults file** — based on `Defaults.xlsx`.
    - **One or more matrix files** — based on `Matrix.xlsx`.

    Then point `Matrix.FolderPath` and `Matrix.DefaultsFile` in your configuration at the matrix folder and defaults file you created.

3.  Run the entrypoint, pointing it at your configuration file:
    ```powershell
    pwsh.exe -Command "& '.\Scripts\Entrypoints\PermissionMatrix.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit `$LASTEXITCODE"
    ```

💡 **TIP**: Use the `Detailed` logging setting in the JSON configuration during initial testing to see granular ACL comparisons (current vs. expected) before running a `Fix` action.

## ⚙️ Configuration

The script is driven by a single JSON configuration file, passed with `-ConfigurationJsonFile`. There's no schema to memorize: copy **[`Example.json`](Examples/Example.json)** and adjust the values to suit your environment. It's a complete, self-documenting template — every option is described inline in its `?` block, including whether it is mandatory, its accepted values, and its default.

💡 **TIP**: Most fields accept an environment-variable reference such as `"ENV:SMTP_SERVER"` in place of a literal value, so secrets like SMTP or ServiceNow credentials never need to live in the file.

### Concurrency

Run time is usually dominated by how much work is allowed to happen at once. Three settings control it, and each caps a different thing:

```json
"MaxConcurrent": {
  "JobsTotal": 30,
  "JobsPerComputer": 3,
  "FoldersPerMatrix": 3
}
```

| Setting            | Caps                                                            |
| ------------------ | --------------------------------------------------------------- |
| `JobsTotal`        | Jobs running at the same time across **all** remote computers.    |
| `JobsPerComputer`  | Jobs running at the same time on any **single** remote computer.  |
| `FoldersPerMatrix` | Folders processed at the same time **within one** job.            |

One job means one matrix file being applied to one computer, and it holds exactly one WinRM shell from start to finish. So `JobsPerComputer` is also the most shells this script will ever open on a single machine — that's the number to keep under the remote computer's `MaxShellsPerUser`. It is enforced structurally rather than by the throttle, so no other setting can cause it to be exceeded.

`JobsTotal` is the overall budget and must be greater than or equal to `JobsPerComputer`; a per-computer cap above the total would be unreachable, and the configuration is rejected at validation.

#### Sizing them

Matrix files are grouped by the computer they target, and each computer's jobs are handed to up to `JobsPerComputer` workers that pull from a shared queue. Workers are started deepest-queue-first, so the busiest computer begins immediately rather than waiting behind lighter ones.

The consequence is that **an uneven spread of matrix files matters more than the raw numbers**. A run where most matrix files target the same file server is limited by that one computer, no matter how high `JobsTotal` goes:

- A run of 99 matrix files with 82 of them on a single server, at `JobsPerComputer: 1`, spent ~65 minutes working through that one queue while everything else sat idle.
- Raising `JobsPerComputer` to `3` divides that queue three ways.

Once the queue is divided, the floor becomes the **single slowest matrix file**, since one matrix cannot be split across workers. Beyond that point, raising `JobsPerComputer` buys nothing and `FoldersPerMatrix` is the setting that helps, because it parallelises the folders inside that one matrix.

A practical order to tune in:

1. Set `JobsPerComputer` from the remote computers' `MaxShellsPerUser`, leaving headroom.
2. Set `JobsTotal` to roughly `JobsPerComputer` × the number of computers you want busy at once.
3. If one long-running matrix still dominates, raise `FoldersPerMatrix`.

💡 **TIP**: The `Duration` column in the summary mail shows the time for each matrix file. If the total run time is far greater than the longest individual duration, work is being serialised somewhere and there is room to raise these numbers.

## 📄 Matrix Files

The per-folder rules live in the Excel matrix files, with shared fallbacks in the `DefaultsFile`. Two conventions are worth knowing.

### Permission Codes

Permissions in the matrix are defined using the following shorthand:

| Code | Permission   | Description                                                       |
| ---- | ------------ | ----------------------------------------------------------------- |
| L    | List         | Allows viewing the contents of the folder.                        |
| R    | Read         | Allows reading files and folder attributes.                       |
| W    | Write        | Allows creating and modifying files.                              |
| F    | Full Control | Grants all permissions, including changing permissions/ownership. |
| I    | Ignore       | The script will skip this entry without making changes.           |

ℹ️ **NOTE**: `I` (Ignore) is valid inside a matrix file but is **not** accepted in the `DefaultsFile`, where every default entry must be an explicit grant (`L`, `R`, `W`, or `F`).

### Default Permissions (`ApplyDefaultPermissions`)

Each matrix's `Settings` sheet has an `ApplyDefaultPermissions` column (`True`/`False`) that controls whether the shared defaults are applied to that folder:

- **True** — merges the global fallback permissions from your `DefaultsFile` into the folder (e.g., ensuring "IT Staff" automatically gets baseline access). To prevent accidental lockouts, if an AD object is defined in **both** the matrix and the defaults, the script halts and throws a conflict error.
- **False** — ignores the defaults entirely; the folder's security is managed strictly by the explicit entries in that matrix (ideal for sensitive HR or Finance folders where IT should not have default access).

## ⏱️ Scheduled Task Example

For automated, recurring executions (e.g., via Windows Task Scheduler), use the same invocation as above, wrapped for the scheduler to pass parameters directly:

- Program: `pwsh.exe`
- Arguments:

```
-Command "& 'C:\Tools\Permission-matrix-set-permissions\Scripts\Entrypoints\PermissionMatrix.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

## 🔍 Access Review (Audit Report)

Alongside the main run that *applies* permissions, the tool ships a second entrypoint that **reviews** them without changing anything. `PermissionMatrixAuditReport.ps1` reuses the read-and-resolve stage of the pipeline — importing the matrices and resolving Active Directory — but never runs the steps that modify ACLs, so it is safe to run at any time.

For each matrix it e-mails the responsible person (resolved from the matrix's `MatrixResponsible` field) a copy of that folder's current access as an Excel attachment, asking them to confirm it is still correct. This closes the loop on access governance: permissions are not only set automatically, they are also periodically signed off by the people who own them.

The audit is driven by its own JSON configuration, based on **[`ExampleAuditReport.json`](Examples/ExampleAuditReport.json)**, with its own `Settings.SendMail` block and subject/body templates. Administrators listed in `AuditReport.ScriptAdmin` are blind-copied on every review mail and receive summaries of any matrices that were skipped due to fatal errors or responsibles that could not be reached.

Run it the same way as the main script, typically on a slower schedule (e.g., monthly) via Task Scheduler:

```powershell
pwsh.exe -Command "& '.\Scripts\Entrypoints\PermissionMatrixAuditReport.ps1' -ConfigurationJsonFile 'C:\MyAuditConfig.json'; exit `$LASTEXITCODE"
```

💡 **TIP**: To preview the review mails before sending them to real owners, set `Settings.SendMail.To` in the audit configuration. When populated it acts as a test override — every mail goes to those address(es) only, the per-matrix responsible is ignored, and no blind copies are sent.

## 📊 Logging & Monitoring

The script provides three layers of feedback, all configured under `Settings`:

1. **File Logs**: Saved to the path defined in `Settings.SaveLogFiles.Where.Folder`.
2. **Event Log**: Writes to the Windows Event Log named in `Settings.SaveInEventLog.LogName` (e.g., `Scripts`) for system-level auditing.
3. **Email Reports**: Summaries sent via SMTP based on your `Settings.SendMail` settings.

## 🙏 Acknowledgements

This script relies on the excellent work of others in the PowerShell and .NET communities. We extend our sincere gratitude to the developers and maintainers of the following projects:

- [MailKit](https://github.com/jstedfast/MailKit) & [MimeKit](https://mimekit.net/): For powerful and flexible email sending functionality.
- [ImportExcel](https://github.com/dfinke/ImportExcel): For seamless importing and exporting of data to Excel files in PowerShell.

Their contributions make this script possible!

## 🤝 Contributing & Support

- **Contributing**: Feel free to open issues or submit pull requests if you have suggestions for improvements or bug fixes.
- **Support**: If you encounter any issues or have questions, please open an issue on the GitHub repository.

## 📜 License

This software is provided for your use, free of charge, for both private and professional purposes.

**NO GUARANTEES ARE GIVEN:** This software is provided "as is," without any warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement.

**YOUR RESPONSIBILITY:** In no event shall the author be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software. You are solely responsible for the code you deploy and its consequences.
