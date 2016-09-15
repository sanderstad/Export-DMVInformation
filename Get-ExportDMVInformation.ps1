# Original file from https://github.com/psget/psget/
#
# Adjusted to import the Export-DMVInformation module

param (
  [string[]]$url = ("https://raw.githubusercontent.com/sanderstad/Export-DMVInformation/master/Export-DMVInformation.psd1", "https://raw.githubusercontent.com/sanderstad/Export-DMVInformation/master/Export-DMVInformation.psm1")
)

function Find-Proxy() {
    if ((Test-Path Env:HTTP_PROXY) -Or (Test-Path Env:HTTPS_PROXY)) {
        return $true
    }
    Else {
        return $false
    }
}

function Get-Proxy() {
    if (Test-Path Env:HTTP_PROXY) {
        return $Env:HTTP_PROXY
    }
    ElseIf (Test-Path Env:HTTPS_PROXY) {
        return $Env:HTTPS_PROXY
    }
}

function Get-File {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [String] $Url,

        [Parameter(Mandatory=$true)]
        [String] $SaveToLocation
    )
    $command = (Get-Command Invoke-WebRequest -ErrorAction SilentlyContinue)
    if($command -ne $null) {
        if (Find-Proxy) {
            $proxy = Get-Proxy
            Write-Host "Proxy detected"
            Write-Host "Using proxy address $proxy"
            Invoke-WebRequest -Uri $Url -OutFile $SaveToLocation -Proxy $proxy
        }
        else {
            Invoke-WebRequest -Uri $Url -OutFile $SaveToLocation
        }
    }
    else {
        $client = (New-Object Net.WebClient)
        $client.UseDefaultCredentials = $true
        if (Find-Proxy) {
            $proxy = Get-Proxy
            Write-Host "Proxy detected"
            Write-Host "Using proxy address $proxy"
            $webproxy = new-object System.Net.WebProxy
            $webproxy.Address = $proxy
            $client.proxy = $webproxy
        }
        $client.DownloadFile($Url, $SaveToLocation)
    }
}

function Install-ExportDMVInformation {
  
    param (
      [string[]]
      # URL to the respository to download Export-DMVInformation from
      $url
    )
  
    $ModulePaths = @($env:PSModulePath -split ';')
    # $Export-DMVInformationDestinationModulePath is mostly needed for testing purposes,
    if ((Test-Path -Path Variable:ModulePath) -and $ModulePath) {
        $Destination = $ModulePath
        if ($ModulePaths -notcontains $Destination) {
            Write-Warning 'Export-DMVInformation install destination is not included in the PSModulePath environment variable'
        }
    }
    else {
        $ExpectedUserModulePath = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules
        $Destination = $ModulePaths | Where-Object { $_ -eq $ExpectedUserModulePath }
        if (-not $Destination) {
            $Destination = $ModulePaths | Select-Object -Index 0
        }
    }
    New-Item -Path ($Destination + "\Export-DMVInformation\") -ItemType Directory -Force | Out-Null

    Write-Host ('Downloading Export-DMVInformation from {0}' -f $url[0])
    Get-File -Url $url[0] -SaveToLocation "$Destination\Export-DMVInformation\Export-DMVInformation.psm1"

    Write-Host ('Downloading Export-DMVInformation from {0}' -f $url[1])
    Get-File -Url $url[1] -SaveToLocation "$Destination\Export-DMVInformation\Export-DMVInformation.psd1"

    $executionPolicy = (Get-ExecutionPolicy)
    $executionRestricted = ($executionPolicy -eq "Restricted")
    if ($executionRestricted) {
        Write-Warning @"
Your execution policy is $executionPolicy, this means you will not be able import or use any scripts including modules.
To fix this change your execution policy to something like RemoteSigned.
        PS> Set-ExecutionPolicy RemoteSigned
For more information execute:
        PS> Get-Help about_execution_policies
"@
    }

    if (!$executionRestricted) {
        # ensure Export-DMVInformation is imported from the location it was just installed to
        Import-Module -Name $Destination\Export-DMVInformation
    }
    Write-Host "Export-DMVInformation is installed and ready to use" -Foreground Green

}

Install-ExportDMVInformation -Url $url