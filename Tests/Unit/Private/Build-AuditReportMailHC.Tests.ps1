BeforeAll {
    $root = Resolve-Path "$PSScriptRoot\..\..\.."
    $moduleRoot = "$root\Modules\PermissionMatrix"

    Get-ChildItem "$moduleRoot\Private" -Filter '*.ps1' -File |
    ForEach-Object { . $_.FullName }

    $script:formData = [pscustomobject]@{
        MatrixFileName          = 'BEL-MTX-FIN'
        MatrixFilePath          = '\\srv\matrices\BEL-MTX-FIN.xlsx'
        MatrixCategoryName      = 'Finance'
        MatrixSubCategoryName   = 'Accounting'
        MatrixFolderPath        = '\\srv\data\Finance'
        MatrixFolderDisplayName = 'Finance share'
        MatrixResponsible       = 'alice@example.com, bob@example.com'
    }

    # Three rows: one direct user + one group with two members
    $script:accessList = @(
        [pscustomobject]@{
            Type = 'user'; SamAccountName = 'jdoe'; Name = 'John Doe'
            MemberSamAccountName = $null
        }
        [pscustomobject]@{
            Type = 'group'; SamAccountName = 'grp1'; Name = 'GRP-FIN'
            MemberSamAccountName = 'asmith'
        }
        [pscustomobject]@{
            Type = 'group'; SamAccountName = 'grp1'; Name = 'GRP-FIN'
            MemberSamAccountName = 'bjones'
        }
    )

    $script:mailSettings = [pscustomobject]@{
        From            = 'no-reply@example.com'
        FromDisplayName = 'Audit'
        Bcc             = @('admin@example.com')
        Subject         = '{{MatrixFileName}}: {{UniqueUserCount}} users, {{UniqueGroupCount}} groups'
        Body            = '<p>Matrix {{MatrixFileName}} owned by {{MatrixResponsible}}. Review at {{RequestTicketURL}}.</p>'
    }

    $script:commonParams = @{
        FormData         = $script:formData
        AccessList       = $script:accessList
        AttachmentPath   = 'C:\log\BEL-MTX-FIN.xlsx'
        MailSettings     = $script:mailSettings
        RequestTicketURL = 'https://portal/req'
    }
}

Describe 'Build-AuditReportMailHC' {
    Context 'recipients' {
        It 'resolves To from MatrixResponsible, split and trimmed' {
            $r = Build-AuditReportMailHC @commonParams
            $r.To | Should -Be @('alice@example.com', 'bob@example.com')
        }

        It 'merges configured Bcc with extra Bcc addresses uniquely' {
            $r = Build-AuditReportMailHC @commonParams `
                -Bcc @('admin@example.com', 'audit@example.com')

            $r.Bcc | Should -Contain 'admin@example.com'
            $r.Bcc | Should -Contain 'audit@example.com'
            @($r.Bcc | Where-Object { $_ -eq 'admin@example.com' }).Count | Should -Be 1
        }

        It 'carries From and FromDisplayName from the mail settings' {
            $r = Build-AuditReportMailHC @commonParams
            $r.From | Should -Be 'no-reply@example.com'
            $r.FromDisplayName | Should -Be 'Audit'
        }
    }

    Context 'counts' {
        It 'counts unique users across direct users and group members' {
            $r = Build-AuditReportMailHC @commonParams
            # jdoe + asmith + bjones = 3
            $r.Subject | Should -Be 'BEL-MTX-FIN: 3 users, 1 groups'
        }

        It 'returns zero counts for an empty AccessList' {
            $r = Build-AuditReportMailHC @commonParams -AccessList @()
            $r.Subject | Should -Be 'BEL-MTX-FIN: 0 users, 0 groups'
        }

        It 'counts each group only once' {
            $list = @(
                [pscustomobject]@{ Type = 'group'; Name = 'A'; MemberSamAccountName = 'm1' }
                [pscustomobject]@{ Type = 'group'; Name = 'A'; MemberSamAccountName = 'm2' }
                [pscustomobject]@{ Type = 'group'; Name = 'B'; MemberSamAccountName = 'm3' }
            )
            $r = Build-AuditReportMailHC @commonParams -AccessList $list
            $r.Subject | Should -Be 'BEL-MTX-FIN: 3 users, 2 groups'
        }
    }

    Context 'body template' {
        It 'replaces tokens in the body from the configured template' {
            $r = Build-AuditReportMailHC @commonParams

            $r.Body | Should -BeLike '*Matrix BEL-MTX-FIN owned by alice@example.com, bob@example.com*'
            $r.Body | Should -BeLike '*Review at https://portal/req.*'
        }

        It 'leaves no unresolved {{tokens}} in subject or body' {
            $r = Build-AuditReportMailHC @commonParams
            $r.Body | Should -Not -BeLike '*{{*'
            $r.Subject | Should -Not -BeLike '*{{*'
        }

        It 'does not treat a value containing $ as a regex substitution' {
            $fd = $script:formData.PSObject.Copy()
            $fd.MatrixFolderPath = '\\srv\share$\Finance'
            $settings = $script:mailSettings.PSObject.Copy()
            $settings.Body = 'Path: {{MatrixFolderPath}}'

            $r = Build-AuditReportMailHC @commonParams -FormData $fd -MailSettings $settings
            $r.Body | Should -Be 'Path: \\srv\share$\Finance'
        }
    }

    Context 'attachment' {
        It 'attaches the supplied Excel log file' {
            $r = Build-AuditReportMailHC @commonParams
            $r.Attachments | Should -Be 'C:\log\BEL-MTX-FIN.xlsx'
        }
    }
}
