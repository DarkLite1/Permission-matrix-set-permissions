# Permission Matrix: Set Permissions

A robust PowerShell tool designed to **apply, verify, and manage NTFS and SMB permissions** at scale. By reading a centralized Excel-based matrix and a JSON configuration file, the script automates the complex task of ensuring folder security across multiple remote computers.

## 🌟 Features

- **Configurable Transfers**: Define multiple SFTP tasks and actions in a single JSON configuration file.
  Action Modes: Supports Check, Fix, and New actions to audit or apply security settings.
- **Multi-Threaded Performance**: Configure concurrent jobs across multiple computers and folders for high-speed execution.
- **ServiceNow Integration**: Syncs with ServiceNow tables to manage roles and form data.
- **Comprehensive Reporting**: Generates HTML overviews and Excel reports of the current permission state.
- **Automated Notifications**: Integrated SMTP support (via MailKit/MimeKit) to send status updates.
- **Granular Logging**: Detailed logging to XLSX/JSON and the Windows Event Log.

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7.x or later**: This script is built for modern PowerShell.
- **ImportExcel**: Required for .xlsx logging and reading the matrix.
  ```powershell
  Install-Module -Name 'ImportExcel' -Scope 'AllUsers'
  ```
- **Emailing Prerequisites**: Required for email functionality.
  ```powershell
  Install-Package -Name 'MailKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  Install-Package -Name 'MimeKit' -Source 'https://www.nuget.org/api/v2' -Scope 'AllUsers' -SkipDependencies
  ```

  - **Permissions**: Administrative rights on the target file systems and remote management permissions.

## ⚙️ Configuration

The script is driven by a ConfigurationJsonFile. Below is a breakdown of the primary configuration blocks:

1. Concurrency (MaxConcurrent)
Controls the performance impact on the network and remote systems.Computers: Max unique computers to target simultaneously.FoldersPerMatrix: Max folders to iterate within a single job.
2. Matrix Settings (Matrix)
Defines the source of truth for permissions.FolderPath: Where the individual matrix Excel files are stored.DefaultsFile: General settings and default permissions applied to all matrices.ExcludedSamAccountName: AD accounts (like placeholders) to ignore during group membership checks.
3. Permission Codes
Permissions in the Excel matrix are defined using the following shorthand:
| Code | Permission | Description
| --- | --- | ---|
| L | List | Allows viewing the contents of the folder.|
| R | Read | Allows reading files and folder attributes.|
| W | Write | Allows creating and modifying files.|
| F | Full Control | Grants all permissions, including changing permissions/ownership.|
| I| Ignore | The script will skip this entry without making changes.| 

### 🛠 Usage

1.  Clone this repository to your local machine:
    ```bash
    git clone https://github.com/DarkLite1/Permission-matrix-set-permissions.git
    cd Permission-matrix-set-permissions
    ```
2.  Review `Example.json` to understand the configuration options.
3.  Create your own configuration file (e.g., `MyConfig.json`) based on `Example.json`.

[!TIP] Use the `Detailed` logging setting in the JSON configuration during initial testing to see granular ACL comparisons (current vs. expected).


### 📊 Logging & Monitoring

The script provides three layers of feedback:

1. **File Logs**: Saved to the path defined in `SaveLogFiles.Where.Folder`.
2. **Event Log**: Writes to the Windows Event Log (e.g., `Scripts`) for system-level auditing.
3. **Email Reports**: Summaries sent via SMTP based on your `SendMail` settings (Always, OnError, etc.).

### Scheduled Task Example ⏱️

For automated, recurring executions (e.g., via Windows Task Scheduler), you'll typically invoke the script using `pwsh.exe -Command`. This allows you to pass parameters directly.

- Program: `pwsh.exe`
- Arguments:

```
-Command  "& 'C:\Permission-matrix-set-permissions.ps1' -ConfigurationJsonFile 'C:\MyConfig.json'; exit $LASTEXITCODE"
```

## 🙏 Acknowledgements

This script relies on the excellent work of others in the PowerShell and .NET communities. We extend our sincere gratitude to the developers and maintainers of the following projects:

- [MailKit](https://github.com/jstedfast/MailKit) & [MimeKit](https://mimekit.net/): For powerful and flexible email sending functionality.
- [ImportExcel](https://github.com/dfinke/ImportExcel): For seamless importing and exporting of data to Excel files in PowerShell.

Their contributions make this script possible!

## 🤝 Contributing

Feel free to open issues or submit pull requests if you have suggestions for improvements or bug fixes.

## 📞 Support

If you encounter any issues or have questions, please open an issue on the GitHub repository.

## 📜 License

This software is provided for your use, free of charge, for both private and professional purposes.

**NO GUARANTEES ARE GIVEN:** This software is provided "as is," without any warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement.

**YOUR RESPONSIBILITY:** In no event shall the author be liable for any claim, damages, or other liability, whether in an action of contract, tort, or otherwise, arising from, out of, or in connection with the software or the use or other dealings in the software. You are solely responsible for the code you deploy and its consequences.
