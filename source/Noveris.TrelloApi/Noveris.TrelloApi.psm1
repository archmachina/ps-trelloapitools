[CmdletBinding()]
param(
)

################
# Global settings
$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2

<#
#>
Function New-TrelloSession
{
    [CmdletBinding()]
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
        [PSCustomObject]@{
            Key = (([System.Net.NetworkCredential]::new("", $Key)).SecurePassword)
            Token = (([System.Net.NetworkCredential]::new("", $Token)).SecurePassword)
            BaseUrl = $BaseUrl
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
        $session = Import-CliXml $Path

        Test-TrelloValidSession -Session $session

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
        Test-TrelloValidSession -Session $Session

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
        Test-TrelloValidSession -Session $Session

        $params = $Parameters.Clone()

        $byteStr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Session.Key)
        $params["key"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($byteStr)

        $byteStr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Session.Token)
        $params["token"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($byteStr)

        $paramStr = ($params.GetEnumerator() | ForEach-Object { ("{0}={1}" -f $_.Key, $_.Value)}) -join "&"

        $request = @{
            Uri = ("{0}{1}?{2}" -f $Session.BaseUrl, $Endpoint, $paramStr)
            ContentType = "application/json"
            Method = $Method
            UseBasicParsing = $true
        }

        # Add a body to the request, if supplied
        if ($PSBoundParameters.Keys -contains "Body")
        {
            $request["Body"] = $Body
        }

        Write-Verbose ("Request URI: " + $request.Uri)
        Write-Verbose ("Request Method:" + $request.Method)

        $result = Invoke-WebRequest @request
        $result.Content | ConvertFrom-Json
    }
}

<#
#>
Function Get-TrelloBoard
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId
    )

    process
    {
        Invoke-TrelloApi -Session $Session -Endpoint ("/boards/{0}" -f $BoardId)
    }
}

<#
#>
Function Get-TrelloListCards
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ListId
    )

    process
    {
        Invoke-TrelloApi -Session $Session -Endpoint ("/lists/{0}/cards" -f $ListId)
    }
}

<#
#>
Function Get-TrelloBoardCards
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId
    )

    process
    {
        Invoke-TrelloApi -Session $Session -Endpoint ("/boards/{0}/cards" -f $BoardId)
    }
}

<#
#>
Function Get-TrelloBoardLists
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId
    )

    process
    {
        Invoke-TrelloApi -Session $Session -Endpoint ("/boards/{0}/lists" -f $BoardId)
    }
}