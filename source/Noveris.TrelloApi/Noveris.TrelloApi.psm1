
################
# Global settings
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2

<#
#>
Function New-TrelloSession
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Token,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl = "https://api.trello.com/1"
    )

    process
    {
        if ($PSCmdlet.ShouldProcess("Trello Session", "Create"))
        {
            [PSCustomObject]@{
                Key = (([System.Net.NetworkCredential]::new("", $Key)).SecurePassword)
                Token = (([System.Net.NetworkCredential]::new("", $Token)).SecurePassword)
                BaseUrl = $BaseUrl
            }
        }
    }
}

<#
#>
Function Test-TrelloValidSession
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session
    )

    process
    {
        if (($Session | Get-Member).Name -notcontains "Key" -or [string]::IsNullOrEmpty($Session.Key))
        {
            Write-Error "Missing Key on Trello session object"
        }

        if (($Session | Get-Member).Name -notcontains "Token" -or [string]::IsNullOrEmpty($Session.Token))
        {
            Write-Error "Missing Token on Trello session object"
        }

        if (($Session | Get-Member).Name -notcontains "BaseUrl" -or [string]::IsNullOrEmpty($Session.BaseUrl))
        {
            Write-Error "Missing BaseUrl on Trello session object"
        }
    }
}

<#
#>
Function Import-TrelloSession
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    process
    {
        # Deserialise from Path to session
        $session = Import-CliXml $Path

        # Make sure we have a valid session object
        Test-TrelloValidSession -Session $session

        # Return the session object
        $session
    }
}

<#
#>
Function Export-TrelloSession
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    process
    {
        # Make sure we have a valid session
        Test-TrelloValidSession -Session $Session

        # Output as serialised XML to Path
        $Session | Export-CliXml $Path
    }
}

<#
#>
Function Invoke-TrelloApi
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Method = "Get",

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Endpoint,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [HashTable]$Parameters = @{},

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [string]$Body
    )

    process
    {
        # Make sure we have a valid session first
        Test-TrelloValidSession -Session $Session

        # Get our own version of the parameters to work with
        $params = $Parameters.Clone()
        $sanitised = $Parameters.Clone()

        # Add key to parameters
        $byteStr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Session.Key)
        $params["key"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($byteStr)
        $sanitised["key"] = "XXXX"

        # Add Token to parameters
        $byteStr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Session.Token)
        $params["token"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($byteStr)
        $sanitised["token"] = "XXXX"

        # Build param string and sanitised param string
        $paramStr = ($params.GetEnumerator() | ForEach-Object { ("{0}={1}" -f $_.Key, $_.Value)}) -join "&"
        $sanitisedStr = ($sanitised.GetEnumerator() | ForEach-Object { ("{0}={1}" -f $_.Key, $_.Value)}) -join "&"

        # Build uri and sanitised URI
        $uri = ("{0}{1}?{2}" -f $Session.BaseUrl, $Endpoint, $paramStr)
        $sanitisedUri = ("{0}{1}?{2}" -f $Session.BaseUrl, $Endpoint, $sanitisedStr)

        # Request parameters for Invoke-WebRequest
        $request = @{
            Uri = $uri
            ContentType = "application/json"
            Method = $Method
            UseBasicParsing = $true
        }

        # Add a body to the request, if supplied
        if ($PSBoundParameters.Keys -contains "Body")
        {
            $request["Body"] = $Body
        }

        # Diagnostic information
        Write-Verbose ("Request URI: " + $sanitisedUri)
        Write-Verbose ("Request Method:" + $request.Method)

        # Make the actual request and return the data as a PS object
        $result = Invoke-WebRequest @request
        $result.Content | ConvertFrom-Json
    }
}

<#
#>
Function Get-TrelloMemberBoards
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Member = "me",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameMatch,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameRegex,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeClosed = $false
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Attempt to retrieve the boards
        $results = Invoke-TrelloApi -Session $Session -Endpoint ("/members/{0}/boards" -f $Member) |
            ForEach-Object { $_ }

        # Filter out closed objects
        if (!$IncludeClosed)
        {
            $results = $results | Where-Object {$false -eq $_.closed}
        }

        # Filter for the first match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameMatch")
        {
            $results = $results |
                Where-Object {$_.Name -eq $FilterFirstNameMatch} |
                Select-Object -First 1
        }

        # Filter for the first match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameRegex")
        {
            $results = $results |
                Where-Object {$_.Name -match $FilterFirstNameRegex} |
                Select-Object -First 1
        }

        $results
    }
}

<#
#>
Function Get-TrelloListCards
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ListId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameMatch,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameRegex,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeClosed = $false
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Change filter to return all cards, if closed objects requested
        $fmt = "/lists/{0}/cards"
        if ($IncludeClosed)
        {
            $fmt = "/lists/{0}/cards/all"
        }

        # Attempt to retrieve the lists
        $results = Invoke-TrelloApi -Session $Session -Endpoint ($fmt -f $ListId) |
            ForEach-Object { $_ }

        # Filter out closed objects
        if (!$IncludeClosed)
        {
            $results = $results | Where-Object {$false -eq $_.closed}
        }

        # Filter for the first match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameMatch")
        {
            $results = $results |
                Where-Object {$_.Name -eq $FilterFirstNameMatch} |
                Select-Object -First 1
        }

        # Filter for the first regex match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameRegex")
        {
            $results = $results |
                Where-Object {$_.Name -match $FilterFirstNameRegex} |
                Select-Object -First 1
        }

        $results
    }
}

<#
#>
Function Get-TrelloBoardCards
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameMatch,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameRegex,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeClosed = $false
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Change filter to return all cards, if closed objects requested
        $fmt = "/boards/{0}/cards"
        if ($IncludeClosed)
        {
            $fmt = "/boards/{0}/cards/all"
        }

        # Attempt to retrieve the lists
        $results = Invoke-TrelloApi -Session $Session -Endpoint ($fmt -f $BoardId) |
            ForEach-Object { $_ }

        # Filter out closed objects
        if (!$IncludeClosed)
        {
            $results = $results | Where-Object {$false -eq $_.closed}
        }

        # Filter for the first match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameMatch")
        {
            $results = $results |
                Where-Object {$_.Name -eq $FilterFirstNameMatch} |
                Select-Object -First 1
        }

        # Filter for the first regex match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameRegex")
        {
            $results = $results |
                Where-Object {$_.Name -match $FilterFirstNameRegex} |
                Select-Object -First 1
        }

        $results
    }
}

<#
#>
Function Get-TrelloLists
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameMatch,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$FilterFirstNameRegex,

        [Parameter(Mandatory=$false)]
        [switch]$IncludeClosed = $false
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Change filter to return all cards, if closed objects requested
        $fmt = "/boards/{0}/lists"
        if ($IncludeClosed)
        {
            $fmt = "/boards/{0}/lists/all"
        }

        # Attempt to retrieve the lists
        $results = Invoke-TrelloApi -Session $Session -Endpoint ($fmt -f $BoardId) |
            ForEach-Object { $_ }

        # Filter out closed objects
        if (!$IncludeClosed)
        {
            $results = $results | Where-Object {$false -eq $_.closed}
        }

        # Filter for the first match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameMatch")
        {
            $results = $results |
                Where-Object {$_.Name -eq $FilterFirstNameMatch} |
                Select-Object -First 1
        }

        # Filter for the first regex match against the filter name, if supplied
        if ($PSBoundParameters.Keys -contains "FilterFirstNameRegex")
        {
            $results = $results |
                Where-Object {$_.Name -match $FilterFirstNameRegex} |
                Select-Object -First 1
        }

        $results
    }
}

<#
#>
Function Add-TrelloList
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [Parameter(Mandatory=$true,ParameterSetName="Options")]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [Parameter(Mandatory=$true,ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId,

        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [ValidateNotNull()]
        [string]$Body,

        [Parameter(Mandatory=$true, ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false, ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$Position = "bottom"
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Processing per parameter set name
        switch ($PSCmdlet.ParameterSetName)
        {
            "Body" {
                break
            }
            "Options" {
                $Body = [PSCustomObject]@{
                    name = $Name
                    pos = "bottom"
                } | ConvertTo-Json

                break
            }
        }

        # Create the target list
        Invoke-TrelloApi -Session $Session -Method Post -Endpoint ("/boards/{0}/lists" -f $BoardId) -Body $Body
    }
}

<#
#>
Function Add-TrelloListCard
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [Parameter(Mandatory=$true,ParameterSetName="Options")]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [Parameter(Mandatory=$true,ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$ListId,

        [Parameter(Mandatory=$true,ParameterSetName="Body")]
        [ValidateNotNull()]
        [string]$Body,

        [Parameter(Mandatory=$true, ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter(Mandatory=$false, ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$Position = "bottom",

        [Parameter(Mandatory=$false, ParameterSetName="Options")]
        [ValidateNotNullOrEmpty()]
        [string]$Description = ""
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Processing per parameter set name
        switch ($PSCmdlet.ParameterSetName)
        {
            "Body" {
                break
            }
            "Options" {
                $Body = [PSCustomObject]@{
                    name = $Name
                    pos = $Position
                    desc = $Description
                } | ConvertTo-Json

                break
            }
        }

        # Create the target list
        Invoke-TrelloApi -Session $Session -Method Post -Endpoint ("/lists/{0}/cards" -f $ListId) -Body $Body
    }
}
