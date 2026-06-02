# Permission Matrix: Set Permissions

A robust PowerShell tool designed to **apply, verify, and manage NTFS and SMB permissions** at scale. By reading a centralized Excel-based matrix and a JSON configuration file, the script automates the complex task of ensuring folder security across multiple remote computers.

## 🌟 Features

- **Action Modes**: Supports `Check`, `Fix`, and `New` actions to audit or apply security settings.
- **Multi-Threaded Performance**: Configure concurrent jobs across multiple computers and folders for high-speed execution.
- **Comprehensive Reporting**: Generates HTML overviews and detailed reports of the current permission state and/or changes.
- **Automated Notifications**: Integrated SMTP support (via MailKit/MimeKit) to send status updates by mail.
- **Granular Logging**: Detailed logging to XLSX/JSON and the Windows Event Log.
- **ServiceNow Integration**: Syncs with ServiceNow tables to manage roles and form data.

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
- **Permissions**: Administrative rights on the target file systems and remote management permissions.

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
