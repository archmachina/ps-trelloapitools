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
Function Get-TrelloList
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

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ListName
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Attempt to retrieve the list
        $targetList = Invoke-TrelloApi -Session $Session -Endpoint ("/boards/{0}/lists" -f $BoardId) |
            ForEach-Object { $_ } |
            Where-Object {$_.Name -eq $ListName -and $_.closed -eq $false} |
            Select-Object -First 1

        # Return the list, if we found one
        if (($targetList | Measure-Object).Count -eq 1)
        {
            $targetList
        }
    }
}

<#
#>
Function Add-TrelloList
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BoardId,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ListName,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Position = "bottom"
    )

    process
    {
        # Check for a valid session
        Test-TrelloValidSession -Session $Session

        # Create the target list
        $body = [PSCustomObject]@{
            name = $ListName
            pos = $Position
        } | ConvertTo-Json
        Write-Verbose "Creating list with properties: $body"
        $targetList = Invoke-TrelloApi -Session $Session -Method Post -Endpoint ("/boards/{0}/lists" -f $BoardId) -Body $body

        # Return a copy of the target lsit
        $targetList
    }
}
