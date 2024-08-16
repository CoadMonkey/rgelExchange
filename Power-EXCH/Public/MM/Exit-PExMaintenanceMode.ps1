Function Exit-PExMaintenanceMode
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
    Version 1.3 :: 08-Aug-2024  :: [Improve] :: Verbage updates, more verbose messages, additional checks. -CoadMonkey
                                :: [Bugfix]  :: Errors if there are non-DAG databases -CoadMonkey
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
	    $VerbosePreference = "Continue"
		$FunctionName = '{0}' -f $MyInvocation.MyCommand
		Write-Verbose "$FunctionName :: Started at [$(Get-Date)] on $Server" -Verbose:$True
        If ($Server -notin (Get-ExchangeServer).name) {
            throw "$Server is not an Exchange server. Use -Server to specify a different name."
        }
        $RunLocal = $False
        If ($env:COMPUTERNAME -eq $Server) { $RunLocal = $True }
        $MailboxServer = Get-MailboxServer -Identity $Server
        If ( $MailboxServer.DatabaseAvailabilityGroup -eq $null ) {    # If Server is not a DAG member
            $TotalStep = 3
        } Else {
            If ($RunLocal) {
                $DAG = (Get-Cluster).Name
            } Else {
                $DAG = Invoke-Command -ComputerName $Server -ScriptBlock {
                    (Get-Cluster).Name
                }
            }
            $TotalStep = 5
        }        
		$i = 0
		$WarningPreference = 'SilentlyContinue'
	}
	Process
	{
		### Bring online all Offline disks & Cancel ReadOnly flag ###
		$i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Bring online $($DegradedDiskCount) degraded disks" `
						-PercentComplete ($i/$($TotalStep) * 100)
        If ($RunLocal) {
            $DegradedDisk = Get-Disk | Where-Object { $_.OperationalStatus -ne 'Online' -or $_.IsReadOnly }
        } Else {
            $DegradedDisk = Invoke-Command -ComputerName $Server -ScriptBlock {
                Get-Disk | Where-Object { $_.OperationalStatus -ne 'Online' -or $_.IsReadOnly }
            }
        }
		if ($DegradedDisk)
		{
			$DegradedDiskCount = ($DegradedDisk | Measure-Object -Property Number -Line).Lines
			$disk = if ($DegradedDiskCount -gt 1) { 'disks' }
			else { 'disk' }
			if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Bring online $($DegradedDiskCount) degraded $($disk)"))
			{
				$DegradedDisk | Where-Object { $_.OperationalStatus -ne 'Online' } | Set-Disk -IsOffline:$false
				$DegradedDisk | Where-Object { $_.IsReadOnly } | Set-Disk -IsReadOnly:$false
			}
            Write-Verbose "[Step $i of $TotalStep] $($DegradedDiskCount) degraded disks set to online and Read-Write"
		}
		else
		{
			Write-Verbose "[Step $i of $TotalStep] All disks are Online"
		}
		
		### Activate Exchange componenets [ServerWideOffline] ###
		$i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Activate Exchange server componenets" `
						-PercentComplete ($i/$($TotalStep) * 100)
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Activate Exchange server components"))
		{
			Set-ServerComponentState $Server -Component ServerWideOffline -State Active -Requester Maintenance -Confirm:$false
			Set-ServerComponentState $Server -Component HubTransport -State Active -Requester Maintenance -Confirm:$false
    		$Count = 0
			do
			{
				$Count += 10
                Write-Progress -Activity "Waiting for [Step $i]" `
								-Status "Waiting for Exchange server componenets..." `
								-PercentComplete ($Count) -Id 1
                $ServerStatusAfter = Get-PExMaintenanceMode -Server $Server
                Start-Sleep 1
			}
			while ($ServerStatusAfter.State -ne 'Connected' -and $Count -lt 100)
		    Write-Progress -Activity "Completed" -Completed -Id 1
            If ($Count -ge 100)
            {
                Write-Warning "[Step $i of $TotalStep] Timeout waiting for Exchange server componenets. Please wait and try again."
            }
            Write-Verbose "[Step $i of $TotalStep] Exchange server componenets activated"
		} Else {
            Write-Verbose "[Step $i of $TotalStep] Exchange server componenets skipped"
        }
		
		### Resume DAG Cluster Node ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
		    $i++
			Write-Progress -Activity "$($FunctionName)" `
							-Status "Exchange server: $($Server)" `
							-CurrentOperation "Current operation: [Step $i of $TotalStep] Resume DAG Cluster Node" `
							-PercentComplete ($i/$($TotalStep) * 100)
            Write-Verbose "Checking cluster state..."
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
                    If ($RunLocal) {
                        Resume-ClusterNode -Name $Server | Out-Null
                    } Else {
                        Invoke-Command -ComputerName $Server -ScriptBlock {
                            Resume-ClusterNode -Name $Using:Server | Out-Null
                        }
                    }
                    Write-Verbose "[Step $i of $TotalStep] Resumed DAG Cluster Node"
			    } Else {
                    Write-Verbose "[Step $i of $TotalStep] Resume DAG Cluster Node skipped"
                }
		    } Else {
                Write-Verbose "[Step $i of $TotalStep] DAG Cluster Node is already Up"
            }
		}

		### Enable DB copy automatic activation ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
		    $i++
		    Write-Progress -Activity "$($FunctionName)" `
						    -Status "Exchange server: $($Server)" `
						    -CurrentOperation "Current operation: [Step $i of $TotalStep] Enable DB copy automatic activation" `
						    -PercentComplete ($i/$($TotalStep) * 100)
		    if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Enable DB copy automatic activation"))
		    {
			    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Unrestricted -Confirm:$false
			    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow:$false -Confirm:$false
    		    $Count = 0
                $DatabaseCopyAutoActivationPolicy = (Get-MailboxServer $Server).DatabaseCopyAutoActivationPolicy
			    do
			    {
				    $Count++
                    Write-Progress -Activity "Waiting for [Step $i]" `
								    -Status "Waiting for DatabaseCopyAutoActivationPolicy..." `
								    -CurrentOperation "DatabaseCopyAutoActivationPolicy = $DatabaseCopyAutoActivationPolicy" `
								    -PercentComplete ($Count) -Id 1
                    $DatabaseCopyAutoActivationPolicy = (Get-MailboxServer $Server).DatabaseCopyAutoActivationPolicy
				    Start-Sleep -Seconds 1
			    }
			    while ($DatabaseCopyAutoActivationPolicy -ne "Unrestricted" -and $Count -lt 100)
		        Write-Progress -Activity "Completed" -Completed -Id 1
                If ($Count -ge 100)
                {
                    Write-Warning "[Step $i of $TotalStep] Timeout waiting for DatabaseCopyAutoActivationPolicy. Please wait and try again."
                }
                Write-Verbose "[Step $i of $TotalStep] Enabled DB copy automatic activation"
		    } Else {
                Write-Verbose "[Step $i of $TotalStep] DB copy automatic activation Skipped"
            }
        }

		### Rebalance DAG and/or mount databases ###
        $i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Rebalance DAG / Mount Databases" `
						-PercentComplete ($i/$($TotalStep) * 100 - 1)
        if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Rebalance DAG / Mount Databases"))
        {
            get-mailboxdatabase -Server $Server|? {$_.ReplicationType -eq "None"}|Set-MailboxDatabase -MountAtStartup:$True -Confirm:$False
            If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
                Move-ActiveMailboxDatabase -ActivatePreferredOnServer $Server
                Get-MailboxDatabase -Server $Server|Mount-Database
                Write-Verbose "[Step $i of $TotalStep] Rebalanced DAG / Mounted Database(s)"
		    }
            Else
            {
                Get-MailboxDatabase -Server $Server|Mount-Database
                Write-Verbose "[Step $i of $TotalStep] Mounted Databases"
            }
        }
        Else
        {
            Write-Verbose "[Step $i of $TotalStep] Rebalance DAG / Mount Databases Skipped"
        }
		Write-Progress -Activity "Completed" -Completed -Id 0
	}
	End
	{
        Return $ServerStatusAfter
		Write-Verbose "$FunctionName :: Finished at [$(Get-Date)] on $Server" -Verbose:$True
	}	
}
