﻿Function Exit-PExMaintenanceMode
{
	
<#
.SYNOPSIS
	Exit Exchange Server Maintenance Mode.
.DESCRIPTION
	This function takes Exchange Server out of Maintenance Mode.
.PARAMETER Server
    Exchange server name or object representing Exchange server
.EXAMPLE
	PS C:\> Exit-PExMaintenanceMode
.EXAMPLE
	PS C:\> Exit-PExMaintenanceMode -Confirm:$false
.NOTES
	Author      :: @ps1code
	Version 1.0 :: 26-Dec-2021  :: [Release] :: Beta
	Version 1.1 :: 03-Aug-2022  :: [Improve] :: Progress bar and steps counter, Parameter free function, Disk status check
    Version 1.2 :: 27-Jun-2024  :: [Improve] :: Add ability to run from admin workstation -CoadMonkey
.LINK
	https://ps1code.com/2024/02/05/pexmm/
#>
	
	[CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess)]
	[Alias('Cancel-PExMaintenanceMode', 'Disable-PExMaintenanceMode', 'Exit-PExMM')]
	Param (
		[Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName, HelpMessage = 'Exchange server name or object representing Exchange server')]
		[Alias('Name')]
		[string]$Server = $($env:COMPUTERNAME)
	)
	
	Begin
	{
		$FunctionName = '{0}' -f $MyInvocation.MyCommand
		Write-Verbose "$FunctionName :: Started at [$(Get-Date)]" -Verbose:$true
        If ($Server -notin (Get-ExchangeServer).name) {
            throw "$Server is not an Exchange server. Use -Server to specify a different name."
        }
        $RunLocal = $False
        If ($env:COMPUTERNAME -eq $Server) { $RunLocal = $True }
        If ($RunLocal) {
            $DAG = (Get-Cluster).Name
        } Else {
            $DAG = Invoke-Command -ComputerName $Server -ScriptBlock {
                (Get-Cluster).Name
            }
        }
		$i = 0
		$TotalStep = 6
		$WarningPreference = 'SilentlyContinue'
	}
	Process
	{
		### Bring online all Offline disks & Cancel ReadOnly flag ###
		$i++
		if ($DegradedDisk = Get-Disk | Where-Object { $_.OperationalStatus -ne 'Online' -or $_.IsReadOnly })
		{
			$DegradedDiskCount = ($DegradedDisk | Measure-Object -Property Number -Line).Lines
			$disk = if ($DegradedDiskCount -gt 1) { 'disks' }
			else { 'disk' }
			if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Bring online $($DegradedDiskCount) degraded $($disk)"))
			{
				Write-Progress -Activity "$($FunctionName)" `
							   -Status "Exchange server: $($Server)" `
							   -CurrentOperation "Current operation: [Step $i of $TotalStep] Bring online $($DegradedDiskCount) degraded disks" `
							   -PercentComplete ($i/$($TotalStep) * 100)
				$DegradedDisk | Where-Object { $_.OperationalStatus -ne 'Online' } | Set-Disk -IsOffline:$false
				$DegradedDisk | Where-Object { $_.IsReadOnly } | Set-Disk -IsReadOnly:$false
			}
		}
		else
		{
			Write-Verbose "[Step $i of $TotalStep] All disks on the [$Server] are Online" -Verbose:$true
		}
		
		### Activate Exchange [ServerWideOffline] componenet ###
		$i++
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Activate Exchange server [ServerWideOffline] componenet"))
		{
			Write-Progress -Activity "$($FunctionName)" `
						   -Status "Exchange server: $($Server)" `
						   -CurrentOperation "Current operation: [Step $i of $TotalStep] Activate Exchange server [ServerWideOffline] componenet" `
						   -PercentComplete ($i/$($TotalStep) * 100)
			Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance -Confirm:$false
		}
		
		### Resume DAG Cluster Node ###
		$i++
        If ($RunLocal) {
            $ClusterNode = (Get-ClusterNode -Name $Server).State -eq 'Paused'
        } Else {
            $ClusterNode = Invoke-Command -ComputerName $Server -ScriptBlock {
                (Get-ClusterNode -Name $Using:Server).State -eq 'Paused'
            }
        }
		if ($ClusterNode)
		{
			if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Resume DAG Cluster Node"))
			{
				Write-Progress -Activity "$($FunctionName)" `
							   -Status "Exchange server: $($Server)" `
							   -CurrentOperation "Current operation: [Step $i of $TotalStep] Resume DAG Cluster Node" `
							   -PercentComplete ($i/$($TotalStep) * 100)
                If ($RunLocal) {
                    Resume-ClusterNode -Name $Server | Out-Null
                } Else {
                    Invoke-Command -ComputerName $Server -ScriptBlock {
                        Resume-ClusterNode -Name $Using:Server | Out-Null
                    }
                }
			}
		}
        If ($RunLocal) {
            Get-ClusterNode | Select-Object Name, Cluster, ID, State -Unique | Format-Table -AutoSize
        } Else {
            Invoke-Command -ComputerName $Server -ScriptBlock {
                Get-ClusterNode | Select-Object Name, Cluster, ID, State -Unique | Format-Table -AutoSize
            }
        }
		
		### Enable DB copy automatic activation ###
		$i++
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Enable DB copy automatic activation"))
		{
			Write-Progress -Activity "$($FunctionName)" `
						   -Status "Exchange server: $($Server)" `
						   -CurrentOperation "Current operation: [Step $i of $TotalStep] Enable DB copy automatic activation" `
						   -PercentComplete ($i/$($TotalStep) * 100)
			Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$false
			Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow:$false -Confirm:$false
		}
		
		### Take the server out of Maintenance Mode ###
		$i++
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Take the server out of Maintenance Mode"))
		{
			Write-Progress -Activity "$($FunctionName)" `
						   -Status "Exchange server: $($Server)" `
						   -CurrentOperation "Current operation: [Step $i of $TotalStep] Take the server out of Maintenance Mode" `
						   -PercentComplete ($i/$($TotalStep) * 100)
			Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance -Confirm:$false
		}
		
		### Rebalance DAG, this will return your active DB copies to their most preferred DAG member ###
		$i++
		$ServerStatusAfter = Get-PExMaintenanceMode -Server $Server
		if ($ServerStatusAfter.State -eq 'Connected')
		{
			$ServerStatusAfter
			if ($PSCmdlet.ShouldProcess("DAG [$($DAG)]", "[Step $i of $TotalStep] Rebalance the DAG databases"))
			{
				Write-Progress -Activity "$($FunctionName)" `
							   -Status "Exchange server: $($Server)" `
							   -CurrentOperation "Current operation: [Step $i of $TotalStep] Rebalance the DAG [$($DAG)] databases" `
							   -PercentComplete ($i/$($TotalStep) * 100)
				Start-Sleep -Seconds 300
				& "$($exscripts)\RedistributeActiveDatabases.ps1" -DagName $DAG -BalanceDbsByActivationPreference -SkipMoveSuppressionChecks -Confirm:$false -ErrorAction SilentlyContinue
			}
		}
	}
	End
	{
        If ($RunLocal) {
            Get-ClusterNode | Select-Object Name, Cluster, ID, State -Unique | Format-Table -AutoSize
        } Else {
            Invoke-Command -ComputerName $Server -ScriptBlock {
                Get-ClusterNode | Select-Object Name, Cluster, ID, State -Unique | Format-Table -AutoSize
            }
        }
		Write-Verbose "$FunctionName :: Finished at [$(Get-Date)]" -Verbose:$true
	}	
}