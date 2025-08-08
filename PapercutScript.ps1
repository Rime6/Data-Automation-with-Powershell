#TOPdesk Integration
#set variables

function connectTD {

    #Select which environment to use by uncommenting the one desired
    #$customerurl= 'https://uottawa-test.topdesk.net'
    $customerurl= 'https://uottawa.topdesk.net'
#setup base64 string
#Ser username and password for API in TOPdesk
$Text = 'apiuser:****PASSWORD******'

##Code to build list of valid operator with email address (removes api / mailimport /admin and system operator)
$Bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
$EncodedAppPass =[Convert]::ToBase64String($Bytes)


$headers = @{
      'Authorization' = 'BASIC ' + $EncodedAppPass
      }
    return $headers
}

function queryUOAccessName {
    param (
        [string]$name,
        [string]$email,
        [int]$employeeID,
        $connect
    )
    if(!($employeeID)){
    $urlTDIntegration = "https://uottawa.topdesk.net/tas/api/persons?query=dynamicName=='$name';personExtraFieldA.name=='employee'"

    $query = Invoke-RestMethod -Uri $urlTDIntegration -Headers $connect

if(!$query){
    if($email -notcontains "."){
        $urlTDIntegration = "https://uottawa.topdesk.net/tas/api/persons?query=email=='$email';personExtraFieldA.name=='employee'"
        $query = Invoke-RestMethod -Uri $urlTDIntegration -Headers $connect

    }
}
    }else{
        $urlTDIntegration = "https://uottawa.topdesk.net/tas/api/persons?query=employeeNumber=='$employeeID';personExtraFieldA.name=='employee'"

    $query = Invoke-RestMethod -Uri $urlTDIntegration -Headers $connect

    }
    return $query
}


##*****************************************************************

#Update CSVPath with file received from CAO that has the user and their budget, should be filled using the provided template
#Update outputPath with path where to save on local computer and keep it as tsv
$CSVPath = "***LOCALPATH***"
$outputPath = "***LOCALPATH***.tsv"
$connection = connectTD
$rows = Import-Csv -Path $CSVPath
foreach ($item in $rows) {
    if($item.Name.Contains(",")){
        $parseName = $item.Name.split(", ")
        $name2 = $parseName[1] + " " + $parseName[0]
    }

    try {
        if(!($item.employeeID)){
                $item.uoaccess = (queryUOAccessName -name $name2 -email $item.email -connect $connection)[0].email.split("@")[0]
            }else{
                $item.uoaccess = (queryUOAccessName -employeeID $item.employeeID -connect $connection)[0].email.split("@")[0]
            }

    }
    catch {
    #DATADUMP to diagnose error if any
       Write-Output("Error : "+ $rows.IndexOf($item))
       Write-Output($rows.IndexOf($item))
       Write-Output($item.Name)
       Write-Output($parseName)
       Write-Output($name2)
    }
    
    
}
# Filter rows with Action = "Add" or "To be removed"
$addUsers = $rows | Where-Object {
    $_.Action -and ($_.Action.Trim().ToLower() -eq "add")
}

$removeUsers = $rows | Where-Object {
    $_.Action -and ($_.Action.Trim().ToLower() -eq "remove")
}

# Group by Program ID
$groupedAdd = $addUsers | Group-Object "Budget"
$groupedRemove = $removeUsers | Group-Object "Budget"

# Output list
$output = @()

# Make "Add" rows into Csv format 
foreach ($group in $groupedAdd) {
    $programId = $group.Name
    $uoaccessList = $group.Group | ForEach-Object { $_.uoaccess } | Where-Object { $_ } | Sort-Object -Unique
    $userField = "+" + ($uoaccessList -join "|")

    $faculty = $group.Group[0]."Faculty"

    $Limited = $group.Group[0]."Limited (Cannot go under 0$)"

    $enabled = $group.Group[0].Enabled

    $budget = $group.Group[0].Amount

    $line = [PSCustomObject]@{
        "Parent Account Name" = $faculty
        "Sub-account Name"    = $programId
        "Enabled"             = $enabled
        "Account PIN/Code"    = ""
        "Credit Balance"      = $budget
        "Restricted Status"   = $Limited
        "Users"               = $userField
    }

    $output += $line
}

# Make "To be removed" rows into Csv format
foreach ($group in $groupedRemove) {
    $programId = $group.Name

    $uoaccessList = $group.Group | ForEach-Object { $_.uoaccess } | Where-Object { $_ } | Sort-Object -Unique
    $userField = "-" + ($uoaccessList -join "|")

    $faculty = $group.Group[0]."Faculty"

    $line = [PSCustomObject]@{
        "Parent Account Name" = $faculty
        "Sub-account Name"    = $programId
        "Enabled"             = ""
        "Account PIN/Code"    = ""
        "Credit Balance"      = ""
        "Restricted Status"   = ""
        "Users"               = $userField
    }

    $output += $line
}

# Export to TSV without headers
$output | ForEach-Object {
    "$($_.'Parent Account Name')`t$($_.'Sub-account Name')`t$($_.Enabled)`t$($_.'Account PIN/Code')`t$($_.'Credit Balance')`t$($_.'Restricted Status')`t$($_.Users)"
} | Set-Content -Path $outputPath

Write-Output "TSV file successfully written to $outputPath"
