###############################################################
# Create-ESGGroups.ps1
###############################################################

Clear-Host

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Functions.ps1"

Write-Log "========================================"
Write-Log "CREATE ESG GROUPS"
Write-Log "========================================"

try
{
    $Excel = Import-Excel `
        -Path $ExcelFile `
        -WorksheetName $Worksheet
}
catch
{
    throw "Unable to open $ExcelFile"
}

$Total = $Excel.Count
$Current = 0

foreach($Row in $Excel)
{
    $Current++

    $Mailbox = $Row.SharedMailbox.Trim()

    $Members = $Row.FullAccessUsers

    if([string]::IsNullOrWhiteSpace($Mailbox))
    {
        continue
    }

    New-Progress `
        -Current $Current `
        -Total $Total `
        -Mailbox $Mailbox

    Write-Log "----------------------------------------"
    Write-Log "Mailbox : $Mailbox"

    $GroupName = Get-GroupName $Mailbox

    $Alias = Get-Alias $Mailbox

    Write-Log "ESG Group : $GroupName"

    Write-Log "Alias     : $Alias"

    Create-Group `
        -Mailbox $Mailbox

    Add-GroupMembers `
        -GroupName $GroupName `
        -Members $Members

    Add-Result `
        -Mailbox $Mailbox `
        -Group $GroupName `
        -Users $Members `
        -Status "CREATED"
}

Write-Progress `
    -Activity "Create ESG Groups" `
    -Completed

Start-DeltaSync

Wait-ForSync

Export-Result

Finish-Script
