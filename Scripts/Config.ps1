###############################################################
# CONFIGURATION
###############################################################

# Simulation
# $true  = aucune modification
# $false = exécution réelle
$Simulation = $true

###############################################################
# EXCEL
###############################################################

$ExcelFile = "C:\Temp\SharedMailbox_Usage_Audit.xlsx"
$Worksheet = "SharedMailboxes"

###############################################################
# EXPORTS
###############################################################

$ExportFolder = "C:\Temp"

$LogFile = Join-Path $ExportFolder "CreateESG.log"

$ResultFile = Join-Path $ExportFolder "CreateESG_Result.xlsx"

###############################################################
# ACTIVE DIRECTORY
###############################################################

$OrganizationalUnit = "skopos.de/staff/Sicherheitsgruppen/Mailboxen"

###############################################################
# ADSYNC
###############################################################

$RunADSync = $true

$WaitAfterSyncMinutes = 5

###############################################################
# EXCHANGE
###############################################################

$GroupPrefix = "esg_"

###############################################################
# COLORS
###############################################################

$ColorInfo = "Cyan"
$ColorSuccess = "Green"
$ColorWarning = "Yellow"
$ColorError = "Red"

###############################################################
# REPORT
###############################################################

$Report = [System.Collections.Generic.List[object]]::new()

###############################################################
# TIMING
###############################################################

$ScriptStart = Get-Date

###############################################################
# MODULES
###############################################################

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Import-Module ADSync -ErrorAction SilentlyContinue
Import-Module ImportExcel -ErrorAction Stop

###############################################################
# VERIFY EXCEL
###############################################################

if (!(Test-Path $ExcelFile))
{
    throw "Excel file not found : $ExcelFile"
}

###############################################################
# VERIFY EXPORT FOLDER
###############################################################

if (!(Test-Path $ExportFolder))
{
    New-Item `
        -ItemType Directory `
        -Path $ExportFolder `
        -Force | Out-Null
}
