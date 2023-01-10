$ErrorActionPreference = 'Stop' #'Inquire' # Other valid values are 'Continue', 'SilentlyContinue', 'Stop', and 'Ignore'

<#
    .NOTES
        The two text files used by this script were generated via the following two lines of code on two different PC's:
            Set-Location C:\Users\Owner\Downloads
            reg query 'hklm' /v "*InduSoft*" /s > RegQuery_Indusoft.txt
    .LINK 
        See https://www.rfc-editor.org/ for the standards, specifically:
            https://www.rfc-editor.org/rfc/rfc5322
            https://www.rfc-editor.org/rfc/rfc822.html#section-4.1
        Per the Internet Message Format, messages are structured broadly as fields or field bodies. 
            See https://www.rfc-editor.org/rfc/rfc5322.html#section-2
        The anatomy of an email address is taken from addr-spec Section 3.4.1 of RFC 5322, specifically:
            https://www.rfc-editor.org/rfc/rfc5322.html#section-3.4.1
            also found at:
            https://datatracker.ietf.org/doc/html/rfc5322#section-3.4.1

    .LINK
        For date parsing/formatting, see: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-date
        The '%K' is the timezone offset
        Regex Greedy and Lazy Quantifiers: https://learn.microsoft.com/en-us/dotnet/standard/base-types/quantifiers-in-regular-expressions
#>
<#-------------------------------------------#>
<#           CREATING VARIABLES              #>
<#-------------------------------------------#>
Clear-Host # Just clear the console of anything previous...

# Create a variable list
# New-Variable -Name 'myVarList' -Scope Script -Verbose
<#
# Populate it with an array of variable names
$myVarList = 'line', 'globalHashTbl', 'totalAddresses';
# Traverse the array, creating each item into a variable
$myVarList | ForEach-Object {New-Variable -Name $_ -Scope Script -Verbose};
$totalAddresses = 0
<# # DEBUGGING: Show details about $myVarList
Get-Variable -Name 'myVarList' | Select-Object * -Verbose #>
# MAKE AN EMPTY HASH TABLE
# $globalHashTbl = @{};
#>
<# By using a hash table for vars and vals, you can set the value when the variable is defined.#>
$GlobalVarList = @{
    'line' = '';
    'globalHashTbl' = @{};
    'fieldTallyHash' = @{};
    'totalAddresses' = 0;
    'dateFormat' = 'yyyy-MM-ddTHH:mm:ss %KZ';
    'giantString' = '';
    'ProblemLog' = "$PSScriptRoot\ERROR_LOG.txt";
    'fileFailCount' = 0;
};
<# Use of Set-Variable instead of New-Variable allows for setting values if the var already exists.
    See https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/set-variable?view=powershell-7.3
#>
$GlobalVarList.Keys | ForEach-Object {Set-Variable -Name $_ -Value $GlobalVarList.$_ -Scope Script -Verbose};
Set-Content -Path $ProblemLog -Value '';
<#################################################################>
<##################### START FORMAT-JSON #########################>
function Format-Json {
    <#
    .SYNOPSIS
        Prettifies JSON output.
    .DESCRIPTION
        Reformats a JSON string so the output looks better than what ConvertTo-Json outputs.
        Post/Answer by user "Theo" at https://stackoverflow.com/questions/56322993/proper-formating-of-json-using-powershell
    .PARAMETER Json
        Required: [string] The JSON text to prettify.
    .PARAMETER Minify
        Optional: Returns the json string compressed.
    .PARAMETER Indentation
        Optional: The number of spaces (1..1024) to use for indentation. Defaults to 4.
    .PARAMETER AsArray
        Optional: If set, the output will be in the form of a string array, otherwise a single string is output.
    .EXAMPLE
        $json | ConvertTo-Json  | Format-Json -Indentation 2
    #>
    [CmdletBinding(DefaultParameterSetName = 'Prettify')]
    Param(
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [string]$Json,

        [Parameter(ParameterSetName = 'Minify')]
        [switch]$Minify,

        [Parameter(ParameterSetName = 'Prettify')]
        [ValidateRange(1, 1024)]
        [int]$Indentation = 4,

        [Parameter(ParameterSetName = 'Prettify')]
        [switch]$AsArray
    )

    if ($PSCmdlet.ParameterSetName -eq 'Minify') {
        return ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100 -Compress
    }

    # If the input JSON text has been created with ConvertTo-Json -Compress
    # then we first need to reconvert it without compression
    if ($Json -notmatch '\r?\n') {
        $Json = ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100
    }

    $indent = 0
    $regexUnlessQuoted = '(?=([^"]*"[^"]*")*[^"]*$)'

    $result = $Json -split '\r?\n' |
        ForEach-Object {
            # If the line contains a ] or } character, 
            # we need to decrement the indentation level unless it is inside quotes.
            if ($_ -match "[}\]]$regexUnlessQuoted") {
                $indent = [Math]::Max($indent - $Indentation, 0)
            }

            # Replace all colon-space combinations by ": " unless it is inside quotes.
            $line = (' ' * $indent) + ($_.TrimStart() -replace ":\s+$regexUnlessQuoted", ': ')

            # If the line contains a [ or { character, 
            # we need to increment the indentation level unless it is inside quotes.
            if ($_ -match "[\{\[]$regexUnlessQuoted") {
                $indent += $Indentation
            }

            $line
        }

    if ($AsArray) { return $result }
    return $result -Join [Environment]::NewLine
};
<##################### END FORMAT-JSON ###########################>
<#################################################################>

<#################################################################>
<##### START TO EXTRACT RELEVANT LINES FROM AN EML FILE ##########>
function Get-MatchingLinesFromEml {
    param (
        [Parameter(Mandatory, HelpMessage = 'Full path to file.')][string]$filePath,
        [Parameter(Mandatory = $false,HelpMessage = 'Delimiter for data within each line.')][string]$intralinearDelim
    )
    # LOAD FILE CONTENTS
    # $rawFileContent = Get-Content $filePath -Verbose # LOAD FILE CONTENTS
    
    $emailAddr = '';
    $emailAddr2 = '';
    # $emailFrom = ''
    $latestSubj = '';
    $emlDateStmp = '';
    $subjectLine = '';
    $currentFieldName = '';


    # Write-Output("Raw Line from file: $($rawFileContent[3])"); # DEBUGGING
    Get-Content -LiteralPath $filePath | ForEach-Object { # FOR EACH LINE...
        # EXTRACT FIRST INSTANCE OF A DATE FIELD IN THE .EML AND NO OTHERS.
        if($_ -match "^(?<fieldName>(Delivery-d|D)ate):\s*(?<dateStamp>.*)") {
            $currentFieldName = $Matches.fieldName;
            if($emlDateStmp -eq ''){$emlDateStmp = (Get-Date -Date ($Matches.dateStamp) -Format $dateFormat);} # Write-Output("Date Stamp: $($Matches.dateStamp)");
        }
        # EXTRACT FIRST INSTANCE OF A SUBJECT IN THE .EML. THIS SHOULD CONTAIN THE "DELAYED" OR "FAILED" LINE.
        elseif($_ -match "^(?<fieldName>Subject):\s.*(?<lineSuffix>delayed|failed).*") {
            # $Matches.lineSuffix either is "delayed" or "failed," so it can be used as one of those two keys.
            if($latestSubj -eq ''){$latestSubj = $Matches.lineSuffix;} # Write-Output("Latest Subj: $($Matches.lineSuffix)");
            $currentFieldName = $Matches.fieldName;
        }
        <# EXTRACT "FINAL-RECIPIENT" LINE, WHICH CONTAINS THE EMAIL ADDRESS TO WHICH THE EMAIL COULD NOT BE DELIVERED.
         # BELOW REGEX LINE ACCOUNTS FOR EITHER RFC 822 OR THE NEWER RFC 5322 
         # NOTE THAT GREEDY AND LAZY QUANTIFIERS MAKE A DIFFERENCE HERE. #>
        # elseif($_ -match "^(?<fieldName>(Final-Recipient|From)):.*?(?<EmailAddrSpec>(?<localPart>[a-zA-Z0-9_\.\-\+]{1,})@(?<domain>(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+))>{0,1}") {
        elseif($_ -match "^(?<fieldName>([A-Za-z0-9\-]+Recipients?|From)):.*?(?<EmailAddrSpec>(?<localPart>[a-zA-Z0-9_\.\-\+]{1,})@(?<domain>(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+))>{0,1}") {
            $currentFieldName = $Matches.fieldName;
            # Write-Output($emailAddr + "(localPart is $($Matches.localPart))  in $filePath"); # IF THE EMAIL ADDRESS DOESN'T ALREADY EXIST IN THE HASH TABLE, ADD IT.
            if($Matches.fieldName -eq 'From'){
                if($emailAddr -eq ''){
                    $emailAddr = $Matches.EmailAddrSpec; # By setting this as a function-level var, you can use it as a check elsewhere
                }
            } else {
                if($emailAddr2 -eq ''){$emailAddr2 = $Matches.EmailAddrSpec;}
            }
        } 
        # THE FIELD "Thread-Topic" CONTAINS THE SUBJECT LINE OF THE EMAIL THAT HAD ISSUES.
        elseif ($_ -match "(?<fieldName>Thread-Topic):\s*(?<lineSuffix>.*)") {
            $currentFieldName = $Matches.fieldName;
            $subjectLine = $Matches.lineSuffix; # Write-Output("Subject Line: $subjectLine");
        }
        elseif($_ -match "^(?<fieldName>[A-Za-z0-9\-]+):.*$"){$currentFieldName = $Matches.fieldName;}
        # elseif(($emailAddr -ne '') -and ($latestSubj)) 

        if($fieldTallyHash.Contains($currentFieldName)){$fieldTallyHash[$currentFieldName]++}
        else{$fieldTallyHash.Add($currentFieldName,1)}
    }; 
    # ALL LINES OF FILE HAVE NOW BEEN READ. DO FINAL WRITING TO OBJECT FROM FUNCTION-LEVEL VARS
    if($emailAddr2 -ne ''){$emailAddr = $emailAddr2}
    if(-not($globalHashTbl.Contains($emailAddr))){ 
        $emailAddr -match "(?<localPart>[a-zA-Z0-9_\.\-\+]{1,})@(?<domain>(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+)"
        $globalHashTbl.Add($emailAddr, [ordered]@{
            'localPart' = $Matches.localPart;
            'domain' = $Matches.domain; #$domain;
            'occurrences' = 1; 
            'delayed' = 0;
            'failed' = 0;
            'latestInstance' = '';
            'earliestInstance' = '';
            'affectedSites' = @{};
            # 'subjects' = [ordered]@{'delayed' = 0; 'failed' = 0};
        })
        
    }
    if($emailAddr -ne ''){
        if($subjectLine -ne ''){
            # In case no subject line in the document contains "delayed" or "failed"...
            $globalHashTbl.$emailAddr.$latestSubj++; # $globalHashTbl.$emailAddr.subjects.$latestSubj++;
            
            if($globalHashTbl.$emailAddr.affectedSites.Contains($subjectLine)){
                $globalHashTbl.$emailAddr.affectedSites.$subjectLine += 1;
                # Write-Output("Site Name: '$($subjectLine)' not found!")
            } 
            else {$globalHashTbl.$emailAddr.affectedSites.Add($subjectLine, 1);}
        }
        # IF THE EMAIL ADDRESS ALREADY EXISTS IN THE HASH TABLE, INCREMENT ITS "occurrences" PROPERTY. 
        else {
                $globalHashTbl.$emailAddr.occurrences++; #= $globalHashTbl.$emailAddr.occurrences += 1;
        }
        if($emlDateStmp -gt ''){
            # If object key 'latestInstance' is less than the current $emlDateStmp, update latestInstance key.
            if($globalHashTbl.$emailAddr.latestInstance -lt $emlDateStmp){$globalHashTbl.$emailAddr.latestInstance = $emlDateStmp;}
            if(($globalHashTbl.$emailAddr.earliestInstance -eq '') -or ($globalHashTbl.$emailAddr.earliestInstance -gt $emlDateStmp)){$globalHashTbl.$emailAddr.earliestInstance = $emlDateStmp;}
        }
    } else {
        $fileFailCount = $fileFailCount + 1;
        Add-Content -Path $ProblemLog -Value "File Extraction Error Count: $fileFailCount  Error on: '$filePath'";  
        Write-Output("`t" + 'CONTENT ERROR ON: ' + $filePath);
    } 
    [System.GC]::Collect()
};
<##### END FUNCTION EXTRACT RELEVANT LINES FROM AN EML FILE ######>
<#################################################################>

<#################################################################>
<########## EXTRACT AND COUNT FIELD NAMES FROM EML FILE ##########>
function Get-TallyOfFieldNames {
    param (
        [Parameter(Mandatory, HelpMessage = 'Full path to file.')][string]$filePath,
        [Parameter(Mandatory = $false,HelpMessage = 'Delimiter for data within each line.')][string]$intralinearDelim
    )
    # LOAD FILE CONTENTS
    # $rawFileContent = Get-Content $filePath -Verbose # LOAD FILE CONTENTS
    
    $emailAddr = '';
    # $emailFrom = ''
    $latestSubj = '';
    $emlDateStmp = '';
    $fieldName = '';

    # Write-Output("Raw Line from file: $($rawFileContent[3])"); # DEBUGGING
    Get-Content -LiteralPath $filePath | ForEach-Object { # FOR EACH LINE...
        if($_ -match "^(?<fieldName>[A-Za-z0-9\-]+):.*$") {
            if($fieldTallyHash.Contains($Matches.fieldName)){$fieldTallyHash["$($Matches.fieldName)"]++}
            else{$fieldTallyHash.Add(($Matches.fieldName),1)}
        }
    }
    # FILE FULLY-READ. NOW DO GC.
    [System.GC]::Collect()
};
<######## END EXTRACT AND COUNT FIELD NAMES FROM EML FILE ########>
<#################################################################>

<#################################################################>
<########################### MAIN START ##########################>
# Write-Output('Val of $myCtr after the function is complete: '+"$myCtr")
# $fileList = Get-ChildItem -Path "$PSScriptRoot\ThunderbirdExports"
Write-Output("$PSScriptRoot\ThunderbirdExports");
# EXECUTE CompareFiles ON EACH FILE IN THE SOURCE FOLDER
Get-ChildItem -Path "$($PSScriptRoot)\ThunderbirdExports" | ForEach-Object {
    # CompareFiles -filePath $_
    # Write-Output($_.BaseName);
    # Get-MatchingLinesFromEml -filePath $_.FullName; # $totalAddresses++;
    Get-TallyOfFieldNames -filePath $_.FullName; # DEBUGGING!
    $totalAddresses++;
    Write-Output("Files Completed:  $totalAddresses  ($($_.BaseName))")
};
<#
# DEBUGGING
$globalHashTbl.Keys | ForEach-Object {
    $totalAddresses = $totalAddresses + $globalHashTbl[$_].occurrences
};
Write-Output("Total Addresses Found: $($totalAddresses) (Should equal number of files in the folder)");
# END DEBUGGING
#>
# Write-Output('Now converting to JSON and saving files...')
# LINE FOR DEPLOYMENT
# $globalHashTbl | ConvertTo-Json -Depth 7 | Format-Json |  Set-Content "$PSScriptRoot\EmlAnalysis.json" -Encoding UTF8;
$fieldTallyHash | ConvertTo-Json -Depth 7 | Format-Json |  Set-Content "$PSScriptRoot\Field_List_and_Frequency.json" -Encoding UTF8;

# LINE FOR DEVELOPMENT
# $globalHashTbl | ConvertTo-Json -Depth 7 | Format-Json;

<# $globalHashTbl.Keys | ForEach-Object {
    if($globalHashTbl[$_].Ctr -gt 1){
        # Write-Output("$_")
        # Write-Output($globalHashTbl[$_])
        Write-Output("Ctr: $($globalHashTbl.$rowKey.Ctr)`tArrayVals_1: $($globalHashTbl.$rowKey.ArrayVals_1)`tValue: $_")
    }
} #>
<#-------------------------------------------#>
<#                CSV OUTPUT                 #>
<#-------------------------------------------#>
# $colHeaders = @('email', 'localPart', 'domain', 'occurrences', 'delayed', 'failed', 'latestInstance', 'earliestInstance', 'affectedSites');
# # $testingKey = 'somebody@domain.com';
# # # $globalHashTbl[$testingKey].keys | ForEach-Object {Write-Output("Value: $($globalHashTbl[$testingKey][$_])    Key: $_")};
# # $rowDat = $globalHashTbl[$testingKey];
# # $globalHashTbl[$testingKey] | ConvertTo-Json -Depth 7 | Format-Json;
# # Write-Output('---------------------------------------------------------------------' + "`n"); # Just a separator
# # Write-Output($rowDat.values);
# # Write-Output('---------------------------------------------------------------------' + "`n"); # Just a separator
# # # $rowDat.GetEnumerator() | ForEach-Object {
# # #     Write-Output("Count is '$($_.Count)'  The value of '$($_.Key)' is: $($_.Value)");
# # # }
# # $rowDat.keys | ForEach-Object {
# #     Write-Output("Count is '$($rowDat.$_.Count)'  The value of '$($_)' is: $($rowDat.$_)");
# # }

# # # Write-Output($rowDat[$colHeaders[-1]].Count);
# # Write-Output('Delayed Count: ' + $rowDat.delayed.Count);
# # Write-Output('Failed Count: ' + $rowDat.failed.Count);
# # Write-Output('latestInstance Count: ' + $rowDat.latestInstance.Count);
# # Write-Output('affectedSites Count: ' + $rowDat.affectedSites.Count);
# # Write-Output($rowDat.affectedSites.Keys -join "; ");
# # # Set-Variable -Name 'line' -Value 'bar?';
# $textHashTbl = [ordered]@{};
# $myCtr = 0;
# # Establish the header row
# $textHashTbl.Add($myCtr,('"' + $($colHeaders -join '", "') + '"'));
# # $myCtr++;
# $globalHashTbl.keys | Sort-Object | ForEach-Object {
#     $rowKey = "$($_)"; $rowStr = ''; $myCtr++;

#     # if($globalHashTbl.$rowKey.$_.Count -eq 1){$rowStr += (', "' + "$($globalHashTbl.$rowKey.$_)" + )}
#     $rowStr = ('"' + $rowKey + '",' + '"' + $globalHashTbl.$rowKey.localPart + '", ' + '"' + $globalHashTbl.$rowKey.domain + '", ' + '"' + 
#         $globalHashTbl.$rowKey.occurrences + '", ' + '"' + $globalHashTbl.$rowKey.delayed + '", ' + '"' + $globalHashTbl.$rowKey.failed + '", ' + '"' + 
#         $globalHashTbl.$rowKey.latestInstance + '", ' + '"' + $globalHashTbl.$rowKey.earliestInstance + '", ' + '"' + 
#         ($globalHashTbl.$rowKey.affectedSites.Keys -join "; ") + '"'
#     );

#     $textHashTbl.Add($myCtr, $rowStr);    
# }
# # Write-Output($textHashTbl.Values -join [Environment]::NewLine); # DEBUGGING
# <#################################################################>
# <########################## DEPLOYMENT ###########################>
# <#################################################################>
# ($textHashTbl.Values -join [Environment]::NewLine) | Set-Content "$PSScriptRoot\EmlAnalysis.csv" -Encoding UTF8;
# # # Set-Variable -Name 'giantString' -Value ('"' + $($colHeaders -join '", "') + '"' + "`n");


# # Write-Output($giantString);
# # $globalHashTbl.keys | Sort-Object; # SORTS BY KEY VALUE
# # $globalHashTbl.GetEnumerator() | Get-Member
# # $globalHashTbl | Sort-Object -Property @{Expression = "occurrences"; Descending = $true} | Get-Member


<#-------------------------------------------#>
<#                  CLEANUP                  #>
<#-------------------------------------------#>
# Traverse the array of variable names, deleting each
# $myVarList | ForEach-Object {Remove-Variable -Name $_ -Scope Script -Force -Verbose}
$GlobalVarList.Keys | ForEach-Object {Remove-Variable -Name $_ -Scope Script -Force -Verbose};
$GlobalVarList.Clear(); 
# Remove-Variable -Name $GlobalVarList -Scope Script -Force -Verbose;
# Delete the variable name array itself.
# Remove-Variable -Name 'myVarList' -Scope Script -Force -Verbose
# Trigger garbage cleanup.
[System.GC]::Collect()
<########################### MAIN END ############################>
<#################################################################>