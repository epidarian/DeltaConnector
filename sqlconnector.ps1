#SQL Exchange Script 1.0.0

#Configuration File Path Parameter
$ConfigFilePath="C:\Users\Public\Downloads\SQL"

#Globalsac
$firstRun = $false 

# Logging setup
$scriptName = $MyInvocation.MyCommand.Name

# Function to log messages
function Log-Message ([string]$message, [string]$level = "Information", $eventID = 36317) {
    #Check if Application Source Exists, if not, create a new App source for script
    If ([System.Diagnostics.EventLog]::SourceExists("$scriptName") -eq $False) {
        New-EventLog -LogName Application -Source "$scriptName"
    }

    # Check if the specified log level is "Error"
    if ($level -eq "Error") {
        # Throw an exception with the provided message
        Throw $message
    }

    #Log to WinEventV
    # Acceptable $level Parameters: Error, Information, FailureAudit, SuccessAudit, Warning (Case-Sensitive)
    Write-EventLog -LogName Application -EventID $eventID -EntryType $level -Source "$scriptName" -Message "$message"

}

#Load config.json
try {
    $config = Get-Content "$ConfigFilePath\config.json" | ConvertFrom-Json 
    Log-Message -message "Loaded configuration from $ConfigFilePath\config.json"
} catch {
    $errorMessage = "Failed to load config file: $_"
    Log-Message -message $errorMessage -level "ERROR"
    Throw $errorMessage
}

#Handle error w/ filesend
Function Error-Out ($message) {
    $jsonOutput.Error += $message 
    Log-Message -message $errorMessage -level "Error"
}

#If DataChecks is 1, set the error message in the json file. Otherwise initialize the storage object blank.
if ($config.DataChecks -eq 1) {
    $jsonOutput = [PSCustomObject]@{ Error = @() }
} else {
    $jsonOutput = [PSCustomObject]@{}
}

#Verify that SQL Server Management Studio exists and the pairing library can be used.
try { 
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null 
} catch {
    $errorMessage = "Unable to load SQL Server Management Studio library: $_"
    Log-Message -message $errorMessage -level "ERROR"
    Error-Out $errorMessage
} 

#Save an open database connection
$serverInstance = 
try {
    New-Object ('Microsoft.SqlServer.Management.Smo.Server') "$($config.DatabaseConnectionString)" 
} catch {
    $errorMessage = "Could not connect to the database: $_"
    Log-Message -message $errorMessage -level "ERROR"
    Error-Out $errorMessage
}

#Prep the summary table that sends at the end. Only if verbose file is enabled.
if ($config.DataChecks -eq 1) {
    $sessionSummary = @{ rowcount = @() }
}

#Prepare the current and previous folders
mkdir "$($config.OutputFilePath)Current\" -ErrorAction Ignore | Out-Null
mkdir "$($config.OutputFilePath)Previous\" -ErrorAction Ignore | Out-Null
$MarkedForDeath = "$($config.OutputFilePath)Previous\"
$FilePath = "$($config.OutputFilePath)Current\"

#Check for first time load
if ( (Get-ChildItem "$($config.OutputFilePath)Previous\" -Force | Select-Object -First 1 | Measure-Object).Count -eq 1 ) { #Detect if there is something in the previous folder
    # Do nothing
} else {
    $jsonOutput.Error += "[$(Get-Date)]::First-Time Run Detected. This could indicate data corruption. Data has been loaded.`n" 
    $firstRun = $true
}

#Check file hashes for tampering
if (Test-Path "$($config.OutputFilePath)metadata.csv") {
    $filesHash = Import-Csv "$($config.OutputFilePath)metadata.csv"
    $comparisonHash = Get-ChildItem -Path $MarkedForDeath -Recurse -File | Get-FileHash -Algorithm SHA1 
    $compareHashResult = Compare-Object -ReferenceObject $filesHash -DifferenceObject $comparisonHash 
    if ( $compareHashResult -eq $null ) {
        #Do nothing if hashes match
    } else {
        $jsonOutput.Error += "[$(Get-Date)]::File Integrity Compromised. Updates can no longer be trusted."
    }
} else {
    $jsonOutput.Error += "[$(Get-Date)]::File Hash Record not found. First run?"
}

#Prep the json updates file
$jsonFilePath = "$FilePath`export_updates_$( (Get-Date -Format 'MMddyyyHHmm' ).ToString() ).json"

#Main loop: For each object in the tables array, run the sql query and dump it to csv, then compare it to the matching previous csv, then save the comparison result to the json object
Foreach ($table in $config.tables) {
    $filename = "$FilePath`export_$($table.name)_$( (Get-Date -Format 'MMddyyyHHmm' ).ToString() ).csv"
    
    #Try to execute the SQL
    try { 
        $results = $serverInstance.Databases[$config.DatabaseName].ExecuteWithResults($table.selectStatement) 
    } catch {
        $errorMessage = "There was an error querying table '$($table.name)' with '$($table.selectStatement)': $_"
        Log-Message -message $errorMessage -level "ERROR"
        $jsonOutput.Error += $errorMessage
        Continue 
    }

    #The first object returned by the executed SQL is the data table- the rest is metadata. 
    $return = $results.Tables[0]

    #If there is no data do nothing, otherwise...
    if ($return -ne $null) {
        #Save it to csv
        $return | Export-Csv $filename -NoTypeInformation

        $localCSV = Import-Csv $filename

        #Load the first-detected previous csv into memory if it exists
        if ( !(Test-Path "$MarkedForDeath`export_$($table.name)_*.csv") ) {
            $ChallengeCSV = ""
        } else {
            $imported = "$MarkedForDeath`export_$($table.name)_*.csv" | Select-Object -First 1 -ErrorAction SilentlyContinue
            $ChallengeCSV = Import-Csv $imported
        }

        #Run the comparison, convert the conversion type field into something readable/useful, and then add it as a child to the parent json object
        $differenceTable = Compare-Object $localCSV $ChallengeCSV -CaseSensitive -Property ($localCSV[0].psobject.Properties.Name) | ForEach-Object {
					#rename sideindicator attribute value
					$itnm = Switch ($_.SideIndicator) { '=>' {$prem = "Delete"}; '<=' {$prem = "Add"}; '==' {$prem = "None"} }
					#Delete sideindicator attribute
					$_.PSObject.Properties.Remove('SideIndicator')
					#add a new attribute called updatetype with the value
					$_ | Add-Member -Name UpdateType -Value $prem -MemberType NoteProperty -PassThru 
				}
        
        # If there are no updates do nothing. Otherwise...
        if ( ($differenceTable -ne $null) -and ($firstRun -eq $false) ) {
            #Add table to the json object
            $jsonOutput | Add-Member -Name "$($table.name)" -Value $differenceTable -MemberType NoteProperty
            #Add the table rowcount to the session summary object
            if ($config.DataChecks -eq 1) {
                $sessionSummary.rowcount += @{tablename = "$($table.name)"; rowcount = $( if ($differenceTable.Count -eq $null) { 1 } else { $differenceTable.Count } )}
            }
        }
    }
}

#Add the session summary object to its desired parent, making it the last item in the json object
if ($config.DataChecks -eq 1) {
    $jsonOutput | Add-Member -Name Summary -Value $sessionSummary.rowcount -MemberType NoteProperty 
}

#Write the json object to file
$jsonOutput | ConvertTo-Json -Depth 100 | Out-File -Encoding utf8 $jsonFilePath
Log-Message "JSON object written to $jsonFilePath"

#Remove the previous folder
rm -Recurse -Force "$MarkedForDeath"

#Remove previous metadata
rm "$($config.OutputFilePath)metadata.csv" -ErrorAction Ignore

try {
    #Send 
    & "$($config.PuTTyPath)\pscp.exe" -P $config.SFTPPort -i $config.KeyFile "$FilePath`*.json" $config.SFTPUploadDestination
} catch {
    $errorMessage = "SFTP Error `n$error"
    Log-Message -message $errorMessage -level "Error"
		Throw $_
}

#Rename current to previous (saves file r/w)
try {
    Rename-Item -Path $FilePath -NewName "Previous" -Force 
} catch {
    $errorMessage = "Error renaming folders after processing csv files. Danger to next run. `n $_"
    Log-Message -message $errorMessage -level "Error"
}

#Metadata filehash snapshot
Get-ChildItem -Path $MarkedForDeath -Recurse -File | Get-FileHash -Algorithm SHA1 | Export-Csv -Path "$($config.OutputFilePath)metadata.csv" -NoTypeInformation
