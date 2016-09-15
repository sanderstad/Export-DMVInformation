################################################################################
#  Written by Sander Stad, SQLStad.nl
# 
#  (c) 2016, SQLStad.nl. All rights reserved.
# 
#  For more scripts and sample code, check out http://www.SQLStad.nl
# 
#  You may alter this code for your own *non-commercial* purposes (e.g. in a
#  for-sale commercial tool). Use in your own environment is encouraged.
#  You may republish altered code as long as you include this copyright and
#  give due credit, but you must obtain prior permission before blogging
#  this code.
# 
#  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF
#  ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED
#  TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
#  PARTICULAR PURPOSE.
#
#  Changelog:
#  v1.0: Initial version
#
################################################################################


function Export-DMVInformation
{
    <# 
    .SYNOPSIS
        Parse the DMV query files made by Glen Berry and export the results to an Excel document
    
    .DESCRIPTION
        The script will parse a specific DMV query file made by Glen Berry.
        After parsing the queries it will loop through each of the queries and
        if needed execute it. 
        The script will write the results to an Excel file.
    
    .PARAMETER instance
        The instance to connect to
        
    .PARAMETER database 
        The database to query out

    .PARAMETER dmvLocation
        The location where to find the DMV query files

    .PARAMETER destination
        The destination where to write the results to

    .PARAMETER excludeinstance
        Flag to exclude the queries for the instance
    
    .PARAMETER username
        Username needed if SQL authentication is required
    
    .PARAMETER password
        Password needed if SQL authentication is required

    .PARAMETER queryTimout
        Timeout how long a query may take in seconds 

    .EXAMPLE
        Get-DMVInformation -instance 'SERVER1' 

    .EXAMPLE
        Get-DMVInformation -instance 'SERVER1' -database 'DB1' -includeInstance $false

    .EXAMPLE    
        Get-DMVInformation -instance 'SERVER1' -database 'DB1' -destination 'C:\Temp\dmv\results'

    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
        Module ImportExcel: https://github.com/dfinke/ImportExcel
        Glenn Berry's DMV site: http://www.sqlskills.com/blogs/glenn/category/dmv-queries/
    #>

    param(
        [Parameter(Mandatory=$true, Position=1)][ValidateNotNullOrEmpty()]
        [string]$instance,
        [Parameter(Mandatory=$false, Position=2)]
        [string]$database = $null,
        [Parameter(Mandatory=$false, Position=3)]
        [string]$username = $null,
        [Parameter(Mandatory=$false, Position=4)]
        [string]$password = $null,
        [Parameter(Mandatory=$false, Position=5)]
        [string]$dmvlocation = ([Environment]::GetFolderPath("MyDocuments") + "\dmv\queries"),
        [Parameter(Mandatory=$false, Position=6)]
        [string]$destination = ([Environment]::GetFolderPath("MyDocuments") + "\dmv\results"),
        [Parameter(Mandatory=$false, Position=7)]
        [bool]$excludeinstance = $false,
        [Parameter(Mandatory=$false, Position=8)]
        [int]$querytimeout = $null
    )

    # Check if The neccesary modules are installed
    if (Get-Module -Name 'ImportExcel') 
    {
            Write-Host 
"Starting DMV Information Retrieval:
- Instance:    $instance
- Database:    $database
- Destination: $destination
"

        # Test the destination
        if(!(Test-Path $destination))
        {
            Write-Host "Destination '$destination' doesn't exist. Creating..."
            New-Item -ItemType directory -Path $destination | Out-Null
        }

    
        # Check if assembly is loaded
        Load-Assembly -name 'Microsoft.SqlServer.Smo'

        # Create the SMO server object
        $srv = New-Object Microsoft.SqlServer.Management.Smo.Server $instance

        if($srv.VersionString -ne $null)
        {
            # Test if the database exists
            if((($database -ne $null) -or ($database -ne '')) -and ($srv.Databases.Name -notcontains $database))
            {
                Write-Host "Database '$database' doesn't exists on '$instance'. Setting database to 'master'." -ForegroundColor Yellow
                $database = 'master'
            }

            # Count the files to see if the script can continue
            if(Test-Path $dmvlocation)
            {       
                $dmvFiles = Get-ChildItem $dmvlocation | Where-Object {$_.Extension -eq ".sql"}

                # If there are no files
                if($dmvFiles.Count -lt 1)
                {
                    Write-Host "No queries found, downloading from Glen Berry's blog"

                    # Dowload the files
                    Download-DMVFiles -destination $dmvlocation

                    # Count the files to see if the script can continue       
                    $dmvFiles = Get-ChildItem $dmvlocation | Where-Object {$_.Extension -eq ".sql"}
                }
            }
            else
            {
                # Dowload the files
                Download-DMVFiles -destination $dmvlocation

                # Count the files to see if the script can continue       
                $dmvFiles = Get-ChildItem $dmvlocation | Where-Object {$_.Extension -eq ".sql"}
            }

            # Check the version string of the server and select the dmv file
            switch($srv.VersionString)
            {
                {$_ -like '9*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2005*'}}
                {$_ -like '10.0*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2008 D*'}}
                {$_ -like '10.5*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2008 R2*'}}
                {$_ -like '11*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2012*'}}
                {$_ -like '12*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2014*'}}
                {$_ -like '13*'} {$dmvFile = $dmvFiles | Where-Object {$_.Name -like 'SQL Server 2016*'}}
            }


            # Declare the variables
            [int]$queryNumber = 0                      # Number of the query
            [bool]$dbSpecific = $false                 # Flag to see if the queries are database specific
            [string]$queryTitle = ""                   # Title of the query, later used for naming the Excel tabs
            [string]$description = ""                  # Description of the query used for informational purposes
            [string]$query = ""                        # The actual query
            [bool]$captureQuery = $false               # Flag to see if the script needs to capture the lines to create the query

            # Create the time stamp
            $timestamp = Get-Date -Format yyyyMMddHHmmss

            # Create the array list to hold all the queries
            $queries = Parse-DMVFile -file $dmvFile.FullName

            # Check if the array contains any queries to execute
            if($queries.Count -ge 1)
            {
                # Loop through all the items
                Foreach($item in $queries)
                {

                    # Reset the result set
                    $result = $null

                    # Check if the query is meant for the instance
                    if(($item.DBSpecific -eq $false) -and ($excludeinstance -eq $false))
                    {
                        Write-Host "Executing Query " $item.QueryNr " - " $item.QueryTitle

                        # Execute the query
                        $result = Execute-Query -instance $instance -database $database -username $username -password $password -query $item.Query -queryTimeout $querytimeout
                        
                        # Check if any values returned and write to the Excel file
                        if($result -ne $null)
                        {
                            $result | Export-Excel -Path "$destination\$($instance.Replace('\', '$'))_$($timestamp).xlsx" -WorkSheetname $($item.QueryTitle) -TableName $("Table" + $item.QueryNr) -TableStyle Dark9 
                            
                        }
                        else
                        {
                            "No Data" | Export-Excel -Path "$destination\$($instance.Replace('\', '$'))_$($timestamp).xlsx" -WorkSheetname $($item.QueryTitle)
                        }

                    }

                    # Check if the query is database specific
                    if(($item.DBSpecific -eq $true) -and (($database -ne $null) -or ($database -ne '')))
                    {
                        Write-Host "Executing Query " $item.QueryNr " - " $item.QueryTitle

                        # Execute the query
                        $result = Execute-Query -instance $instance -database $database -username $username -password $password -query $item.Query -queryTimeout $querytimeout

                        # Check if any values returned and write to the Excel file
                        if($result -ne $null)
                        {
                            $result | Export-Excel -Path "$destination\$($instance.Replace('\', '$'))_$($database)_$($timestamp).xlsx" -WorkSheetname $($item.QueryTitle) -TableName $("Table" + $item.QueryNr) -TableStyle Dark9  
                        }
                        else
                        {
                            "No Data" | Export-Excel -Path "$destination\$($instance.Replace('\', '$'))_$($database)_$($timestamp).xlsx" -WorkSheetname $($item.QueryTitle)
                        }
                    }

                }

            }
        }
        else
        {
            Write-Host "Couldn't connect to instance $instance" -ForegroundColor Red 
        }
    } 
    else 
    {
        Write-Host "Module ImportExcel is not installed or is not imported." -ForegroundColor Red
    }

}

function Download-DMVFiles
{
    <# 
    .SYNOPSIS
        Function to download the DMV files
    
    .DESCRIPTION
        This script will download the DMV files by Glenn Berry to a specific location
    
    .PARAMETER destination
        Destination directory
    
    .EXAMPLE
        Download-DMVFiles -destination 'C:\Temp\dmv\queries' 

    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
    #>

    param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$destination
    )

    # Test the destination
    if(!(Test-Path $destination))
    {
        Write-Host "DMV destination '$destination' doesn't exist. Creating..."
        New-Item -ItemType directory -Path $destination | Out-Null
    }

    # Download the individual files
    $webClient = New-Object System.Net.WebClient

    Write-Host "Downloading DMV Files..."

    # Set the URL and download the files
    $url2005 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202005%20Diagnostic%20Information%20Queries%20%28CY%202016%29.sql'
    $url2008 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202008%20Diagnostic%20Information%20Queries%20%28CY%202016%29.sql'
    $url2008R2 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202008%20R2%20Diagnostic%20Information%20Queries%20%28CY%202016%29.sql'
    $url2012 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202012%20Diagnostic%20Information%20Queries%20%28September%202016%29.sql'
    $url2014 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202014%20Diagnostic%20Information%20Queries%20%28September%202016%29.sql'
    $url2016 = 'https://dl.dropboxusercontent.com/u/13748067/SQL%20Server%202016%20Diagnostic%20Information%20Queries%20%28September%202016%29.sql'
    
    $webClient.DownloadFile($url2005, "$destination\SQL Server 2005 Diagnostic Information Queries.sql")
    $webClient.DownloadFile($url2008, "$destination\SQL Server 2008 Diagnostic Information Queries.sql")
    $webClient.DownloadFile($url2008R2, "$destination\SQL Server 2008 R2 Diagnostic Information Queries.sql")
    $webClient.DownloadFile($url2012, "$destination\SQL Server 2012 Diagnostic Information Queries.sql")
    $webClient.DownloadFile($url2014, "$destination\SQL Server 2014 Diagnostic Information Queries.sql")
    $webClient.DownloadFile($url2016, "$destination\SQL Server 2016 Diagnostic Information Queries.sql")
}

#
function Parse-DMVFile
{
    <# 
    .SYNOPSIS
        Function to parse the DMV file
    
    .DESCRIPTION
        This function will parse the DMV file and put it into an array.
        It will designate each query with a title, description, if its database specific and the query itself
    
    .PARAMETER file
        DMV file to parse
    
    .EXAMPLE
        Parse-DMVFile -file 'C:\temp\queries\file.sql'

    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
    #>

    param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$file
    )

    # Create the result variable 
    $result = @()

    #Loop through each line
    ForEach($line in Get-Content $file)
    {

        # Check if the script is at the database specific queries
        if($line.Contains('Database specific queries'))
        {
            $dbSpecific = $true
        }

        # If the line starts with dashes and has the text for the query number in it
        if($line.StartsWith('--') -and ($line.IndexOf("(Query") -ne -1)) 
        {
            # Empty the query string t
            [string]$query = ""

            # Split the items in the line at the paranthesis
            $items = $line.Trim().Split('()', [System.StringSplitOptions]::RemoveEmptyEntries)

            # Set the variables
            $queryNumber = ($items[($items.Length - 3)]).Replace("Query ", "")
            $queryTitle = $items[($items.Length - 1)]
            $queryDescription = ($items[0].Replace("-", "")).Trim()

            # Set the flag to start capturing the query text
            $captureQuery = $true

        } 

        # Check if the line starts with a selectong elements and the flag is set
        if((($line -match 'SELECT') -or ($line -match 'WITH') -or ($line -match 'EXEC') -or ($line -match 'DBCC')) -and ($captureQuery -eq $false))
        {
            # Set the flag
            $captureQuery = $true

            # Reset the query variable
            $query = ""
        }

        # If the flag is true and the line does not contain any dashes
        if(($captureQuery -eq $true) -and ($line -notlike '--*'))
        {
            # Cleanup the line
            if($line.IndexOf('--') -gt 3)
            {
                $line = $line.Substring(0, $line.IndexOf('--'))
            }

            # Add the line
            $query += "$line "
        }

        # Check if the line is the end of the query
        if($line.StartsWith('------'))
        {
            # Set the flag to false
            $captureQuery = $false

            # Set up the properties
            $props = @{QueryNr=$queryNumber;DBSpecific=$dbSpecific;QueryTitle=$queryTitle;Description=$queryDescription;Query=$query}

            # Create a new object based on the properties
            $queryObject = New-Object psobject -Property $props

            # Add the object to the querie array
            $result += $queryObject
        }
    
    }

    return $result
}

function Execute-Query
{
    <# 
    .SYNOPSIS
        Execute a query
    
    .DESCRIPTION
        The function will create a connection to an instance and execute a query
    
    .PARAMETER instance
        The instance to connect to
        
    .PARAMETER database 
        The database to query out

    .PARAMETER includeInstance
        Flag to inlude the queries for the instance
    
    .PARAMETER username
        Username needed if SQL authentication is required
    
    .PARAMETER password
        Password needed if SQL authentication is required

    .PARAMETER queryTimout
        Timeout how long a query may take in seconds 

    .EXAMPLE
        Execute-Query -instance 'SERVER1' 

    .EXAMPLE
        Execute-Query -instance 'SERVER1' -database 'DB1'  

    .EXAMPLE    
        Execute-Query -instance 'SERVER1' -database 'DB1' -username 'user1' -password 'pass1'

    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
    #>
    param
    (
        [Parameter(Mandatory=$true, Position=1)]
        [string]$instance,
        [Parameter(Mandatory=$false, Position=2)]
        [string]$database = $null,
        [Parameter(Mandatory=$true, Position=3)]
        [string]$query,
        [Parameter(Mandatory=$false, Position=4)]
        [string]$username = $null,
        [Parameter(Mandatory=$false, Position=5)]
        [string]$password = $null,
        [Parameter(Mandatory=$false, Position=6)]
        [int]$querytimeout = 300
    )

    # Create the connection object
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection

    # Check if the database is set
    if($database -eq $null)
    {
        # Check if the sql authentication  or integrated security is needed
        if(($username.Length -ge 1) -and ($password.Length -ge 1))
        {
            # Setup the connection with sql authentication for master database
            $sqlConnection.ConnectionString = “Server=$instance;Database=master;User Id=$username;Password=$password”
        }
        else
        {
            # Setup the connection
            $sqlConnection.ConnectionString = “Server=$instance;Database=master;Integrated Security=True”
        }

    }
    else
    {
        # Check if the sql authentication  or integrated security is needed
        if(($username.Length -ge 1) -and ($password.Length -ge 1))
        {
            # Setup the connection with sql authentication for master database
            $sqlConnection.ConnectionString = “Server=$instance;Database=$database;User Id=$username;Password=$password”
        }
        else
        {
            # Setup the connection
            $sqlConnection.ConnectionString = “Server=$instance;Database=$database;Integrated Security=True”
        }
    }

    # Open the connection
    $sqlConnection.Open()

    # Setup the command
    $sqlCommand = $sqlConnection.CreateCommand()
    $sqlCommand.CommandText = $query
    $sqlCommand.CommandTimeout = $querytimeout

    # Setup the data adapter
    $dataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCommand

    # Setup the dataset
    $dataset = New-Object System.Data.Dataset
    $dataAdapter.Fill($dataset) | Out-Null

    # Execute the query
    $result = $dataset.Tables[0] | Select -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors

    # Close the connection
    $sqlConnection.Close()

    return $result
}

function Load-Assembly
{
    <# 
    .SYNOPSIS
        Check if a assembly is loaded and load it if neccesary
    .DESCRIPTION
        The script will check if an assembly is already loaded.
        If it isn't already loaded it will try to load the assembly
    .PARAMETER  name
        Full name of the assembly to be loaded
    .EXAMPLE
        Load-Assembly -name 'Microsoft.SqlServer.SMO'
    .INPUTS
    .OUTPUTS
    .NOTES
    .LINK
    #>
     [CmdletBinding()]
     param(
          [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
          [String] $name
     )
     
     if(([System.AppDomain]::Currentdomain.GetAssemblies() | where {$_ -match $name}) -eq $null)
     {
        try{
            [System.Reflection.Assembly]::LoadWithPartialName($name) | Out-Null
        } 
        catch [System.Exception]
        {
            Write-Host "Failed to load assembly!" -ForegroundColor Red
            Write-Host "$_.Exception.GetType().FullName, $_.Exception.Message" -ForegroundColor Red
        }
     }
}

Export-ModuleMember -Function Export-DMVInformation
