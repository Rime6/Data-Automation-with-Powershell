$CSVPath = "C:\Users\rnassere\dev\Coop_Scripts\PapercutScript\Papercut_budget_Arts (2).csv"
$outputPath = "C:\Users\rnassere\dev\Coop_Scripts\PapercutScript\papercut_import.tsv"

$rows = Import-Csv -Path $CSVPath

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
