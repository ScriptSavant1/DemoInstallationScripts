# ==============================================================================
# 03_Upgrade_DataProcessor.ps1
# LRE Data Processor Host Upgrade  |  25.1 -> 26.1
# ------------------------------------------------------------------------------
# Run this on each DATA PROCESSOR HOST machine.
# The Data Processor role/purpose in LRE Administration is preserved across
# upgrade - no role reassignment is needed after upgrade.
#
# Prerequisites:
#   - 01_Upgrade_LREServer.ps1 must have completed on the Server.
#   - Public Key must be available at: $InstallerShare\PublicKey.txt
#
# Usage (elevated PowerShell):
#   .\03_Upgrade_DataProcessor.ps1
# ==============================================================================
#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
. "$ScriptRoot\Common-Functions.ps1"
. "$ScriptRoot\..\config\upgrade_config.ps1"

$Global:LogFile = "$LogDir\DataProcessor_Upgrade_$(hostname)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Initialize-Log -LogFile $Global:LogFile

# Dot-source the shared host upgrade logic
. "$ScriptRoot\Shared-HostUpgrade.ps1"

Write-Banner "DATA PROCESSOR HOST UPGRADE  |  25.1 -> 26.1  |  $(hostname)  |  $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

# Invoke shared host upgrade
Invoke-HostUpgrade -ComponentType "DataProcessor"
