<#
.SYNOPSIS
   Exports Active Directory user data to a specified file format.

.DESCRIPTION
   The `Export-ADUserData` function retrieves user information from Active Directory based on the provided search base, processes the data, and exports it to a file in CSV or TXT format. The function supports filtering users and customizing the output.

.PARAMETER OutputPath
   Specifies the file path for the exported data. Supported extensions are `.csv` and `.txt`.

.PARAMETER Format
   Specifies the file format of the output. Accepted values are `CSV` (default) and `TXT`.

.PARAMETER SkipUsers
   A list of usernames (supports regex patterns) to exclude from the export.

.PARAMETER SearchBase
   The distinguished name of the directory search base. Defaults to `"OU=Company,DC=company,DC=com"`.

.EXAMPLES
   # Export user data to a CSV file
   Export-ADUserData -OutputPath "C:\Reports\UserData.csv"

   # Export user data to a TXT file while excluding specific users
   Export-ADUserData -OutputPath "C:\Reports\UserData.txt" -Format TXT -SkipUsers @('jdoe.*', 'admin*')

   # Export user data from a specific organizational unit
   Export-ADUserData -SearchBase "OU=Users,OU=HQ,DC=company,DC=com" -OutputPath "C:\Reports\HQUserData.csv"

.NOTES
   This function requires the Active Directory module (`ActiveDirectory`) to be installed and imported.

   Make sure to provide a valid path and have necessary permissions to create or overwrite files in the specified directory.

.OUTPUTS
   None. The function writes the result to the specified output file.
#>
function Export-ADUserData {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateScript({
            $ext = [System.IO.Path]::GetExtension($_)
            if ($ext -notin '.csv', '.txt') { 
                throw "Invalid file extension. Supported formats: .csv, .txt" 
            }
            $true
        })]
        [string]$OutputPath,

        [Parameter(Position = 1)]
        [ValidateSet('CSV', 'TXT')]
        [string]$Format = 'CSV',

        [string[]]$SkipUsers = @(),

        [Parameter(HelpMessage = "Distinguished name for the search base")]
        [string]$SearchBase = "OU=Company,DC=company,DC=com"
    )

    begin {
        #requires -Module ActiveDirectory

        # Configuration constants
        $properties = @(
            'displayName', 'rank', 'mail', 'mobile', 'telephoneNumber',
            'sAMAccountName', 'description', 'extensionAttribute2',
            'msRTCSIP-PrimaryUserAddress'
        ) + (1..10 | ForEach-Object { "dxMidOU$_" })

        $patternsToRemove = [System.Collections.Generic.HashSet[string]]@(

            "BORSOD-ABAÚJ-ZEMPLÉN VÁRMEGYEI RENDŐR-FŐKAPITÁNYSÁG",
            "VMRFK HELYI BESOROLÁSÚ SZERVEI",
            "BRFK+19VMRFK"
        )
        
        # Create output directory if needed
        $directory = [System.IO.Path]::GetDirectoryName($OutputPath)
        if (-not (Test-Path -Path $directory)) {
            New-Item -Path $directory -ItemType Directory -Force | Out-Null
        }

        # Rank pattern configuration
        $script:rankReplacements = @{
            '^(c\.)?\s*(r\.)?\s*(ezredes)\b'           = 'ezr.'
            '^(c\.)?\s*(r\.)?\s*(alezredes)\b'         = 'alez.'
            '^(c\.)?\s*(r\.)?\s*(őrnagy)\b'            = 'őrgy.'
            '^(c\.)?\s*(r\.)?\s*(százados)\b'          = 'szds.'
            '^(c\.)?\s*(r\.)?\s*(főhadnagy)\b'         = 'fhdgy.'
            '^(c\.)?\s*(r\.)?\s*(hadnagy)\b'           = 'hdgy.'
            '^(c\.)?\s*(r\.)?\s*(főtörzszászlós)\b'    = 'ftzls.'
            '^(c\.)?\s*(r\.)?\s*(törzszászlós)\b'      = 'tzls.'
            '^(c\.)?\s*(r\.)?\s*(zászlós)\b'           = 'zls.'
            '^(c\.)?\s*(r\.)?\s*(főtörzsőrmester)\b'   = 'ftőrm.'
            '^(c\.)?\s*(r\.)?\s*(törzsőrmester)\b'     = 'túrm.'
            '^(c\.)?\s*(r\.)?\s*(őrmester)\b'          = 'őrm.'
            '^(rendvédelmi alkalmazott|ra)\b'          = 'ria.'
            '^(munkavállaló|mv)\b'                     = 'mv.'
        }
    }

    process {
        try {
            $params = @{
                LDAPFilter  = "(&(objectCategory=person)(objectClass=user)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
                SearchBase  = $SearchBase
                Properties  = $properties
                ErrorAction = 'Stop'
            }

            $users = Get-ADUser @params
            $results = Process-Users -Users $users -SkipUsers $SkipUsers
            Export-Results -Results $results -OutputPath $OutputPath -Format $Format
        }
        catch {
            Write-Error "Failed to export user data: $_"
            throw
        }
    }
}

function Get-ValueOrDefault {
    param(
        [Parameter(ValueFromPipeline)]
        $Value,
        $Default
    )
    
    process {
        if ([string]::IsNullOrWhiteSpace($Value)) { 
            $Default 
        } 
        else { 
            $Value 
        }
    }
}

function Process-Users {
    param(
        [array]$Users,
        [string[]]$SkipUsers
    )

    $skipRegex = if ($SkipUsers) {
        [regex]::new(($SkipUsers -join '|'), 'IgnoreCase')
    }

    $results = [System.Collections.Generic.List[PSObject]]::new($Users.Count)
    
    foreach ($user in $Users) {
        if ($skipRegex -and $skipRegex.IsMatch($user.sAMAccountName)) { 
            continue 
        }

        $displayName = Process-DisplayName -User $user
        $dxMidOU = Process-DxMidOU -User $user
        $skype = Process-Skype -User $user

        $results.Add([PSCustomObject]@{
            sAMAccountName  = $user.sAMAccountName
            Name            = $displayName
            Mobile          = Get-ValueOrDefault $user.mobile ''
            TelephoneNumber = Get-ValueOrDefault $user.telephoneNumber ''
            Mail            = Get-ValueOrDefault $user.mail ''
            Description     = (Get-ValueOrDefault $user.description '').ToLower()
            Skype           = $skype
            dxMidOU         = $dxMidOU
            IsVezeto        = $user.extensionAttribute2 -eq 'vezeto'
            IsHidden        = $false
        })
    }

    return $results
}

function Process-DisplayName {
    param($User)
    
    $displayName = Get-ValueOrDefault $user.displayName ''
    $rank = Get-ValueOrDefault $user.rank ''
    
    if (-not [string]::IsNullOrWhiteSpace($rank)) {
        $processedRank = Process-Rank -Rank $rank
        $displayName = "$displayName $processedRank".Trim()
    }
    
    return $displayName.Trim()
}

function Process-Rank {
    param([string]$Rank)

    # Remove any existing "r." prefix to avoid duplication
    $Rank = $Rank -replace '^r\.\s*', ''

    # Capture if there was a 'c.' in the original rank
    $Prefix = if ($Rank -match '^(c\.)\s*') { 'c. ' } else { '' }

    foreach ($pattern in $rankReplacements.Keys) {
        if ($Rank -match $pattern) {
            $processedRank = $rankReplacements[$pattern]
            
            # Add 'r.' unless it's 'ria' or 'mv'
            if ($processedRank -notmatch 'ria\.|mv\.') {
                $processedRank = 'r. ' + $processedRank.Trim()
            }
            
            # Include 'c.' if it was in the original rank
            return $Prefix + $processedRank
        }
    }

    return $Rank
}

function Process-DxMidOU {
    param($User)

    $replacementConfig = @{
        # 'Search' = 'Replacement'
        'RENDŐRŐRS MEZŐCSÁT \(OSZTÁLY JOGÁLLÁSÚ\)' = 'RENDŐRŐRS MEZŐCSÁT'
        'SZABÁLYSÉRTÉSI ELŐKÉSZÍTŐ CSOPORT \(ÖNÁLLÓ\)' = 'SZABÁLYSÉRTÉSI ELŐKÉSZÍTŐ CSOPORT'
    }

    $dxMidOUs = [System.Collections.Generic.List[string]]::new(10)
    foreach ($i in 1..10) {
        $propValue = $user."dxMidOU$i"
        if (-not [string]::IsNullOrWhiteSpace($propValue) -and -not $patternsToRemove.Contains($propValue)) {
            $cleanValue = $propValue -replace '[;,"]' -replace '\s+', ' '

            foreach ($pattern in $replacementConfig.Keys) {
                $cleanValue = $cleanValue -replace $pattern, $replacementConfig[$pattern]
            }

            $dxMidOUs.Add($cleanValue.Trim())
        }
    }

    if ($dxMidOUs.Count -gt 0) {
        return ($dxMidOUs -join ', ')
    } 
    else {
        return ''
    }
}

function Process-Skype {
    param($User)

    $skype = Get-ValueOrDefault $user.'msRTCSIP-PrimaryUserAddress' ''
    
    if (-not [string]::IsNullOrWhiteSpace($skype)) {
        return ($skype -replace '^sip:', '' -replace '[^a-zA-Z0-9@.]', '').Trim()
    }

    return ''
}

function Export-Results {
    param(
        [array]$Results,
        [string]$OutputPath,
        [string]$Format
    )

    if ($Results.Count -eq 0) {
        Write-Debug "No users to export."
        return
    }

    $sortedResults = $Results | Sort-Object -Property @(
        @{ Expression = 'IsVezeto'; Descending = $true }
        'sAMAccountName'
    ) | Select-Object -ExcludeProperty SortKey*

    if ($Format -eq 'CSV') {
        $sortedResults | Export-Csv -Path $OutputPath -Encoding UTF8 -Delimiter ';' -NoTypeInformation
    }
    else {
        $sortedResults | Format-Table -AutoSize | Out-File -FilePath $OutputPath -Encoding UTF8
    }
}