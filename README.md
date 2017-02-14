# Export-DMVInformation
Export the resuts from Glenn Berry's DMV queries directly to Excel

## EXamples

Below are several examples how the module can be executed

    Export-DMVInformation -instance 'SERVER1'

    Export-DMVInformation -instance 'SERVER1' -database 'DB1' -excludenstance

    Export-DMVInformation -instance 'SERVER1' -database 'DB1' -destination 'C:\Temp\dmv\results'

    'server1', 'server2' | Export-DMVInformation

    'server1', 'server2' | Export-DMVInformation -database 'ALL'


## How to install

The easiest method to install the module is by copying the code below and entering it in a PowerShell window:

(new-object Net.WebClient).DownloadString("https://raw.githubusercontent.com/sanderstad/Export-DMVInformation/master/Get-ExportDMVInformation.ps1") | iex

### Alternative installation method

Alternatively you can download the module from here.

Unzip the file.

Make a directory (if not already present) named "Export-DMVInformation" in one of the following standard PowerShell Module directories:

    $Home\Documents\WindowsPowerShell\Modules (%UserProfile%\Documents\WindowsPowerShell\Modules)
    $Env:ProgramFiles\WindowsPowerShell\Modules (%ProgramFiles%\ WindowsPowerShell\Modules)
    $Systemroot\System32\WindowsPowerShell\v1.0\Modules (%systemroot%\System32\ WindowsPowerShell\v1.0\Modules)

Place both the "psd1" and "psm1" files in the module directory created earlier.

Execute the following command in a PowerShell command screen:

Import-Module Export-DMVInformation

To check if the module is imported correctly execute the following command:

Get-Command -Module Export-DMVInformation or Get-Module -Name Export-DMVInformation

If you see a list with the functions than the module is installed successfully. If you see nothing happening something has gone wrong.

I hope you enjoy the module and that it can help you to get performance information fast.
