function Resolve-ResponsibleEmailHC {
    <#
    .SYNOPSIS
        Resolve a 'MatrixResponsible' value to a flat list of e-mail addresses.

    .DESCRIPTION
        The 'MatrixResponsible' field can be a comma-separated mix of:

          - e-mail addresses        -> kept as-is,
          - AD user accounts        -> resolved to the account's 'mail' attribute,
          - AD groups               -> resolved to the e-mail address of every
                                       member, recursing nested groups down to
                                       their user leaves (Get-ADGroupMember
                                       -Recursive).

        Any entry (a responsible token, or an individual group member) that
        cannot be resolved to an e-mail address is returned in '.Unresolved' so
        the caller can skip it and report it to the admin. Resolved addresses
        are de-duplicated.

        Lookups use the AD cmdlets directly (Get-ADObject / Get-ADGroupMember /
        Get-ADUser), the same ActiveDirectory module the Begin stage relies on.

    .PARAMETER Responsible
        The raw 'MatrixResponsible' string from the matrix FormData worksheet.

    .PARAMETER ExcludeSamAccountName
        SamAccountNames of placeholder accounts (Matrix.AdGroupPlaceHolders).
        These are never returned as recipients: they are dropped whether they
        appear as a group member or as a directly listed responsible, and they
        are not reported as unresolved.

    .OUTPUTS
        [pscustomobject] with:
          .Emails     - sorted, unique list of resolved e-mail addresses
          .Unresolved - tokens / members that produced no e-mail address
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Responsible,
        [string[]]$ExcludeSamAccountName = @()
    )

    $exclude = @($ExcludeSamAccountName | Where-Object { $_ })
    $emails = [System.Collections.Generic.List[string]]::new()
    $unresolved = [System.Collections.Generic.List[string]]::new()

    $tokens = @(
        "$Responsible".Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    )

    foreach ($token in $tokens) {
        #region Already an e-mail address - keep as-is
        if ($token -match '@') {
            $emails.Add($token)
            continue
        }
        #endregion

        #region Skip a placeholder account listed directly by SamAccountName
        if ($token -in $exclude) { continue }
        #endregion

        #region Look the token up in AD
        # Escape LDAP filter metacharacters in the user-supplied value.
        $safe = [regex]::Replace($token, '[\\\(\)\*\x00]', { '\{0:x2}' -f [int][char]$args[0].Value })

        $adObject = $null
        try {
            $adObject = Get-ADObject `
                -LDAPFilter "(|(sAMAccountName=$safe)(name=$safe)(displayName=$safe)(mail=$safe))" `
                -Properties objectClass, mail, sAMAccountName -ErrorAction Stop |
                Select-Object -First 1
        }
        catch { $adObject = $null }

        if (-not $adObject) {
            $unresolved.Add($token)
            continue
        }
        #endregion

        #region Group - resolve to the e-mail address of every (nested) member
        if ($adObject.objectClass -eq 'group') {
            $members = @()
            try {
                $members = @(
                    Get-ADGroupMember -Identity $adObject.DistinguishedName -Recursive -ErrorAction Stop |
                    Where-Object {
                        $_.objectClass -eq 'user' -and ($_.SamAccountName -notin $exclude)
                    }
                )
            }
            catch { $members = @() }

            if (-not $members) {
                $unresolved.Add("$token (group has no user members)")
                continue
            }

            foreach ($member in $members) {
                $user = $null
                try {
                    $user = Get-ADUser -Identity $member.distinguishedName `
                        -Properties EmailAddress -ErrorAction Stop
                }
                catch { $user = $null }

                if ($user -and $user.EmailAddress) {
                    $emails.Add($user.EmailAddress)
                }
                else {
                    $unresolved.Add("$token > $($member.name) (no e-mail)")
                }
            }
            continue
        }
        #endregion

        #region User / contact - take its mail attribute
        if ($adObject.sAMAccountName -and ($adObject.sAMAccountName -in $exclude)) {
            continue
        }
        if ($adObject.mail) {
            $emails.Add($adObject.mail)
        }
        else {
            $unresolved.Add("$token (no e-mail)")
        }
        #endregion
    }

    [pscustomobject]@{
        Emails     = @($emails | Where-Object { $_ } | Sort-Object -Unique)
        Unresolved = @($unresolved | Sort-Object -Unique)
    }
}