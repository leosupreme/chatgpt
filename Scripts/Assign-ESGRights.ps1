###############################################################
# Assign-ESGRights.ps1
###############################################################

Clear-Host

. "$PSScriptRoot\Config.ps1"
. "$PSScriptRoot\Functions.ps1"

Write-Log "========================================"
Write-Log "ASSIGN ESG RIGHTS"
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

    if([string]::IsNullOrWhiteSpace($Mailbox))
    {
        continue
    }

    $GroupName = Get-GroupName $Mailbox

    New-Progress `
        -Current $Current `
        -Total $Total `
        -Mailbox $Mailbox

    Write-Log "----------------------------------------"
    Write-Log "Mailbox : $Mailbox"

    Write-Log "Group   : $GroupName"

    if(!(Group-Exists $GroupName))
    {
        Write-Log "$GroupName not found." "ERROR"

        Add-Result `
            -Mailbox $Mailbox `
            -Group $GroupName `
            -Users "" `
            -Status "GROUP NOT FOUND"

        continue
    }

    ###########################################################
    # FULL ACCESS
    ###########################################################

    if($Simulation)
    {
        Write-Log "[SIMULATION] Add-MailboxPermission FullAccess" "INFO"
    }
    else
    {
        try
        {
            Add-MailboxPermission `
                -Identity $Mailbox `
                -User $GroupName `
                -AccessRights FullAccess `
                -InheritanceType All `
                -AutoMapping:$false `
                -ErrorAction Stop

            Write-Log "FullAccess assigned." "SUCCESS"
        }
        catch
        {
            Write-Log $_.Exception.Message "ERROR"
        }
    }

    ###########################################################
    # SEND AS
    ###########################################################

    if($Simulation)
    {
        Write-Log "[SIMULATION] Add-RecipientPermission SendAs" "INFO"
    }
    else
    {
        try
        {
            Add-RecipientPermission `
                -Identity $Mailbox `
                -Trustee $GroupName `
                -AccessRights SendAs `
                -Confirm:$false `
                -ErrorAction Stop

            Write-Log "SendAs assigned." "SUCCESS"
        }
        catch
        {
            Write-Log $_.Exception.Message "ERROR"
        }
    }

        ###########################################################
    # GRANT SEND ON BEHALF
    ###########################################################

    if($Simulation)
    {
        Write-Log "[SIMULATION] Set-Mailbox GrantSendOnBehalfTo" "INFO"
    }
    else
    {
        try
        {
            Set-Mailbox `
                -Identity $Mailbox `
                -GrantSendOnBehalfTo @{Add=$GroupName} `
                -MessageCopyForSendOnBehalfEnabled $true `
                -MessageCopyForSentAsEnabled $true `
                -ErrorAction Stop

            Write-Log "GrantSendOnBehalf assigned." "SUCCESS"
        }
        catch
        {
            Write-Log $_.Exception.Message "ERROR"
        }
    }

    ###########################################################
    # RESULT
    ###########################################################

    Add-Result `
        -Mailbox $Mailbox `
        -Group $GroupName `
        -Users "" `
        -Status "RIGHTS ASSIGNED"

}

Write-Progress `
    -Activity "Assign ESG Rights" `
    -Completed

Export-Result

Finish-Script
