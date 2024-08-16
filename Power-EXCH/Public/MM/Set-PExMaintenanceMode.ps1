Function Set-PExMaintenanceMode
{
	
<#
.SYNOPSIS
	Put Exchange Server in Maintenance Mode.
.DESCRIPTION
	This function puts Exchange Server into Maintenance Mode.
.PARAMETER Server
    Exchange server name or object representing Exchange server
.EXAMPLE
	PS C:\> Set-PExMaintenanceMode
.EXAMPLE
	PS C:\> Set-PExMaintenanceMode -Confirm:$false
	Silently put Exchange server into Maintenance Mode, restart server at the end.
.NOTES
	Author      :: @ps1code
	Dependency  :: Function     :: Get-PExMaintenanceMode
	Version 1.0 :: 26-Dec-2021  :: [Release] :: Beta
	Version 1.1 :: 03-Aug-2022  :: [Improve] :: Progress bar and steps counter, Parameter free function
	Version 1.2 :: 19-Apr-2023  :: [Bugfix]  :: Error thrown on empty queue
	Version 1.3 :: 27-Aug-2023  :: [Improve] :: Warn about active mailbox move requests
	Version 1.4 :: 27-Aug-2023  :: [Improve] :: Shutdown option
	Version 1.5 :: 03-Sep-2023  :: [Improve] :: Warn about additional servers in MM
    Version 1.6 :: 27-Jun-2024  :: [Improve] :: Add ability to run from admin workstation
                                :: [Bugfix]  :: Divide by 0 error if no mounted databases -CoadMonkey
    Version 1.7 :: 5-Jul-2024   :: [Bugfix]  :: Queue redirect not filtering out ShadowRedundancy queues -CoadMonkey
    Version 1.8 :: 8-Aug-2024   :: [Bugfix]  :: Infinite loop and other errors if there are non-DAG databases -CoadMonkey
                                :: [Improve] :: Reroute messages using AD Sites instead of DAG members. Add more messages. -CoadMonkey

.LINK
	https://ps1code.com/2024/02/05/pexmm/
#>
	
	[CmdletBinding(ConfirmImpact = 'High', SupportsShouldProcess)]
	[Alias('Enter-PExMaintenanceMode', 'Enable-PExMaintenanceMode', 'Enter-PExMM')]
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
        $ExchangeServers = Get-ExchangeServer
        If ($Server -notin ($ExchangeServers).name) {
            Write-Error "$Server is not an Exchange server. Use -Server to specify a different name."
            throw "$Server is not an Exchange server. Use -Server to specify a different name."
        }
        $RunLocal = $False
        If ($env:COMPUTERNAME -eq $Server) { $RunLocal = $True }
        $i = 0
        $MailboxServer = Get-MailboxServer -Identity $Server
        If ( $MailboxServer.DatabaseAvailabilityGroup -eq $null ) {    # If Server is not a DAG member
            $TotalStep = 5
        } Else {
            $TotalStep = 6
        }        
		$WarningPreference = 'SilentlyContinue'
		$Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
		$Fqdn = ([System.Net.Dns]::GetHostByName($env:COMPUTERNAME)).HostName
	}
	Process
	{

        ### Warn about in-progress mailbox migrations ###
		if ($InProgressReq = Get-MoveRequest | Get-MoveRequestStatistics | Where-Object { $_.Status.Value -eq 'InProgress' -and $_.SourceServer, $_.TargetServer -contains $Fqdn })
		{
			$CountReq = ($InProgressReq | Measure-Object -Line -Property DisplayName).Lines
			$InProgressReq | Out-Host
			Write-Warning "The [$Server] is participating in $($CountReq) ACTIVE mailbox migration jobs !!!"
		}
		
		### Warn about another servers in MM ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
            If ($RunLocal) {
                $CurrentMM = Get-ClusterNode | Where-Object { $_.PSComputerName -ne $Server -and $_.State -ne 'Up' }
            } Else {
                $CurrentMM = Invoke-Command -ComputerName $Server -ScriptBlock {
                    Get-ClusterNode | Where-Object { $_.PSComputerName -ne $Server -and $_.State -ne 'Up' }
                }
            }    
		    if ($CurrentMM)
		    {
			    $MMCount = ($CurrentMM | Measure-Object -Property Name -Line).Lines
			    if ($MMCount -eq 1) { Write-Warning "There is an additional server in Maintenance Mode" -Verbose:$true }
			    else { Write-Warning "There are $($MMCount) additional servers in Maintenance Mode" -Verbose:$true }
		    }
        }		

		### Set the Hub Transport service to draining. It will stop accepting any more messages ###
		$i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Set the Hub Transport service to draining" `
						-PercentComplete ($i/$($TotalStep) * 100) -Id 0
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Set the Hub Transport service to draining"))
		{
			Set-ServerComponentState -Identity $Server -Component HubTransport -State Draining -Requester Maintenance -Confirm:$false
		}
        Write-Progress -Activity "Completed" -Completed -Id 0
        Write-Verbose "[Step $i of $TotalStep] Hub Transport service set to draining" 


		### Redirect any queued messages to another server (try same site first) ###
		$i++
        Write-Progress -Activity "$($FunctionName)" `
			-Status "Exchange server: $($Server)" `
			-CurrentOperation "Current operation: [Step $i of $TotalStep] Redirect messages in queue" `
			-PercentComplete ($i/$($TotalStep) * 100) -Id 0

        # Can we skip this step?
        if ((Get-Queue -Server $Server | Measure-Object -Property MessageCount -Sum).Sum)
        {

            # Build a list of servers trying same AD site first
            [array]$ADSites = ($ExchangeServers | Where-Object { $_.name -eq $Server }).Site
            $ADSites += ($ExchangeServers).Site |Where-Object {$_ -ne $ADSites[0]}| Select-Object -Unique
            $EligibleServers = @()
            foreach ($Site in $ADSites )
            {
                $EligibleServers += $ExchangeServers |
                    Where-Object { $_.site -eq $Site -and $_.Name -ne $Server -and $_.name -notlike "*test*"} |
                    Sort-Object { Get-Random }
            }
            Write-Verbose "Eligible Servers: $EligibleServers"

            if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Redirect messages in queue"))
		    {


                # Pick an eligible server not in maintentance mode.
                Foreach ( $TargetHostname in $EligibleServers.name )
                {
                    Write-Verbose "Checking $TargetHostname for maintenance mode..."
				    Write-Progress -Activity "Waiting for [Step $i]" `
								    -Status "Checking $TargetHostname for maintenance mode..." `
								    -PercentComplete ($EligibleServers.name.IndexOf($TargetHostname)/$EligibleServers.count * 100) -Id 1
                    $TargetServerStatus = Get-PExMaintenanceMode -Server $TargetHostname
       		        if ($TargetServerStatus.State -notcontains 'Maintenance') { Break }
                }
                Write-Progress -Activity "Completed" -Completed -Id 1

       		    if ($TargetServerStatus.State -contains 'Maintenance')
                {
                    Get-Queue -Server $Server -ErrorAction SilentlyContinue|? {$_.DeliveryType -ne "ShadowRedundancy"} | Select-Object Identity, DeliveryType, Status, MessageCount
                    Write-Error "[Step $i of $TotalStep] There are no eligible servers for queue redirection! Wait for queue to empty and then re-try"
                    throw "[Step $i of $TotalStep] There are no eligible servers for queue redirection! Wait for queue to empty and then re-try"
                }

                $TargetFqdn = "$($TargetHostname).$($Domain)"
	            Write-Verbose "Redirecting messages to $TargetFqdn"
    		    
                Write-Progress -Activity "$($FunctionName)" `
							    -Status "Exchange server: $($Server)" `
							    -CurrentOperation "Current operation: [Step $i of $TotalStep] Redirecting queued messages to [$($TargetFqdn)]" `
							    -PercentComplete ($i/$($TotalStep) * 100) -Id 0
			    
                ### Save transport queue stats before redirecting messages ###
			    $QueueLength = (Get-Queue -Server $Server | Measure-Object -Property MessageCount -Sum).Sum
				
			    if ($QueueLength)
			    {
	    		    Redirect-Message -Server $Server -Target $TargetFqdn -Confirm:$false
    				
				    ### Wait the transport queue to empty ###
				    Write-Verbose "Waiting for the transport queue to empty..."
                    $Count = 0
				    do
				    {
					    $Count += 2
                        $Queue = Get-Queue -Server $Server -ErrorAction SilentlyContinue #|? {$_.DeliveryType -ne "ShadowRedundancy"}
					    #$Queue | Select-Object Identity, DeliveryType, Status, MessageCount
					    $QueueLengthNow = ($Queue | Measure-Object -Property MessageCount -Sum).Sum
					    $QueuePercent = $($QueueLength - $QueueLengthNow)/$($QueueLength)
					    Write-Progress -Activity "Waiting for [Step $i]" `
									    -Status "Moving $($QueueLengthNow) queued messages to other transport servers ..." `
									    -CurrentOperation "Currently queued: $($QueueLengthNow) messages" `
									    -PercentComplete ($QueuePercent * 100) -Id 1
					    Start-Sleep -Seconds 1
				    }
				    while ($QueueLengthNow -gt 0 -and $Count -lt 100 )
				    Write-Progress -Activity "Completed" -Completed -Id 1
                    If ($Count -ge 100)
                    {
                        Write-Warning "[Step $i of $TotalStep] Timeout waiting for queue to empty. There are $($QueueLengthNow) message outstanding. Please wait and try again."
                    }
			    }
            }
            Write-Verbose "[Step $i of $TotalStep] Queued messages redirected. The transport queue is empty."
        }
        Else
        {
            Write-Verbose "[Step $i of $TotalStep] Redirect messages skipped, the transport queue is empty."
        }

		### Suspend Server from the DAG ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
		    $i++
			Write-Progress -Activity "$($FunctionName)" `
						    -Status "Exchange server: $($Server)" `
						    -CurrentOperation "Current operation: [Step $i of $TotalStep] Suspend Server from Cluster node" `
						    -PercentComplete ($i/$($TotalStep) * 100) -Id 0
		    if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Suspend Server from the DAG"))
		    {
                If ($RunLocal) {
                    Suspend-ClusterNode $Server -Wait -Drain -Confirm:$false | Out-Null
                } Else {
                    Invoke-Command -ComputerName $Server -ScriptBlock {
                        Suspend-ClusterNode $Using:Server -Wait -Drain -Confirm:$false | Out-Null
                    }
                }    
                Start-Sleep 5
                Get-ClusterNode | Select-Object Name, Cluster, ID, State -Unique | Format-Table -AutoSize
		    }
            Write-Verbose "[Step $i of $TotalStep] Server suspended from the DAG"
        }

        ### Disable DB copy automatic activation and move any active DB copies to other DAG members ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -ne $null ) {    # Skip if Server is not a DAG member
        {
		    $i++
			Write-Progress -Activity "$($FunctionName)" `
						    -Status "Exchange server: $($Server)" `
						    -CurrentOperation "Current operation: [Step $i of $TotalStep] Disable DB copy automatic activation & Move any active DB copies to other DAG members" `
						    -PercentComplete ($i/$($TotalStep) * 100) -Id 0
		    if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Disable DB copy automatic activation & Move any active DB copies to other DAG members"))
		    {
			    ### Save mounted DB copies stats before moving ###
			    $dbMountedCount = (Get-MailboxDatabaseCopyStatus -Server $Server | Where-Object { $_.Status -eq "Mounted" } | Measure-Object -Property Name -Line).Lines
			
			    Set-MailboxServer $Server -DatabaseCopyActivationDisabledAndMoveNow:$true -Confirm:$false
			    Set-MailboxServer $Server -DatabaseCopyAutoActivationPolicy Blocked -Confirm:$false
                Move-ActiveMailboxDatabase -Server $Server
                # Dismount non-DAG databases too
                get-mailboxdatabase -Server $Server|? {$_.ReplicationType -eq "None"}|Dismount-Database -Confirm:$False
			
			    ### Wait for any database copies that are still mounted on the server ###
			    Write-Verbose "Waiting for any database copies that are still mounted on the server ..." -Verbose:$true
                $Count = 0
			    do
			    {
				    $Count += 2
                    $dbMounted = Get-MailboxDatabaseCopyStatus -Server $Server | Where-Object { $_.Status -eq "Mounted" }
				    $dbMounteNow = ($dbMounted | Measure-Object -Property Name -Line).Lines -replace '^$', '0'
				    Write-Progress -Activity "Waiting for [Step $i]" `
							       -Status "Moving $($dbMountedCount) DB copies to other DAG members ..." `
							       -CurrentOperation "Currently mounted: $($dbMounteNow) DB copies" `
							       -PercentComplete ($($dbMountedCount - [int]$dbMounteNow)/$($dbMountedCount) * 100) -Id 2
				    Start-Sleep -Seconds 1
			    }
			    while ($dbMounted -and $Count -lt 100)
			    Write-Progress -Activity "Completed" -Completed -Id 2
                If ($Count -ge 100)
                {
                    Write-Warning "[Step $i of $TotalStep] Timeout waiting for databases to dismount. Please wait and try again."
                }
                Else
                {
                    Write-Verbose "[Step $i of $TotalStep] All databases are moved or dismounted"
                }
		    }
            Else
            {
                Write-Verbose "[Step $i of $TotalStep] Disable & Move DB(s) Skipped"
            }
        }
 
 		### Dismount databases ###
        If ( $MailboxServer.DatabaseAvailabilityGroup -eq $null ) {    # Non-DAG servers
        {
            $i++
		    Write-Progress -Activity "$($FunctionName)" `
						    -Status "Exchange server: $($Server)" `
						    -CurrentOperation "Current operation: [Step $i of $TotalStep] Dismount databases" `
						    -PercentComplete ($i/$($TotalStep) * 100) -Id 0
			### Save mounted DB copies stats before moving ###
			$dbMountedCount = (Get-MailboxDatabaseCopyStatus -Server $Server | Where-Object { $_.Status -eq "Mounted" } | Measure-Object -Property Name -Line).Lines

    	    If ($dbMountedCount)
            {
	            if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Dismount databases"))
	            {
                    get-mailboxdatabase -Server $Server|? {$_.ReplicationType -eq "None"}|Dismount-Database -Confirm:$False
                    get-mailboxdatabase -Server $Server|? {$_.ReplicationType -eq "None"}|Set-MailboxDatabase -MountAtStartup:$False
			
                    Write-Verbose "Waiting for all databases to dismount ..."
                    $Count = 0
			        do
			        {
				        $Count += 2
                        $dbMounted = Get-MailboxDatabaseCopyStatus -Server $Server | Where-Object { $_.Status -eq "Mounted" }
				        $dbMounteNow = ($dbMounted | Measure-Object -Property Name -Line).Lines -replace '^$', '0'
				        Write-Progress -Activity "Waiting for [Step $i]" `
							            -Status "Moving $($dbMountedCount) Dismounting Database(s) ..." `
							            -CurrentOperation "Currently mounted: $($dbMounteNow) Database(s)" `
							            -PercentComplete ($($dbMountedCount - [int]$dbMounteNow)/$($dbMountedCount) * 100) -Id 2
				        Start-Sleep -Seconds 1
			        }
    			    while ($dbMounted -and $Count -lt 100)
        	        Write-Progress -Activity "Completed" -Completed -Id 2
                    If ($Count -ge 100)
                    {
                        Write-Warning "[Step $i of $TotalStep] Timeout waiting for databases to dismount. Please wait and try again."
                    }
                    Else
                    {
                        Write-Verbose "[Step $i of $TotalStep] All databases are dismounted"
                    }
                }
                Else
                {
                    Write-Verbose "[Step $i of $TotalStep] Dismount databases Skipped"
                }
            }
            Else
            {
                Write-Verbose "[Step $i of $TotalStep] No mounted databases"
            }
        }

		### Set Server component states to ServerWideOffline ###
		$i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Set Server component states to ServerWideOffline" `
						-PercentComplete ($i/$($TotalStep) * 100) -Id 0
		if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Set Server component states to ServerWideOffline"))
		{
			Set-ServerComponentState $Server -Component ServerWideOffline -State Inactive -Requester Maintenance -Confirm:$false
		}
        Write-Verbose "[Step $i of $TotalStep] Server component states set to ServerWideOffline"
		
		### Reboot/Shutdown the server ###
		$i++
		Write-Progress -Activity "$($FunctionName)" `
						-Status "Exchange server: $($Server)" `
						-CurrentOperation "Current operation: [Step $i of $TotalStep] Reboot or Shutdown the server OS" `
						-PercentComplete ($i/$($TotalStep) * 100) -Id 0
        Write-Verbose "[Step $i of $TotalStep] Checking $Server for Maintenance mode ..."
		$ServerStatusAfter = Get-PExMaintenanceMode -Server $Server
		if ($ServerStatusAfter.State -eq 'Maintenance')
		{
			if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Reboot the server"))
			{
                Write-Verbose "[Step $i of $TotalStep] Restarting $Server"
				Restart-Computer -ComputerName $Server -Force -Confirm:$false
			}
			else
			{
				if ($PSCmdlet.ShouldProcess("Server [$($Server)]", "[Step $i of $TotalStep] Shutdown the server"))
				{
                    Write-Verbose "[Step $i of $TotalStep] Shutting down $Server"
					Stop-Computer -ComputerName $Server -Force -Confirm:$false
                }
			}
		}
        Else
        {
            Write-Error "$($FunctionName) failed, please try again."
        }
		Write-Progress -Activity "Completed" -Completed -Id 0
	}
	End
    {
        Write-Verbose "$FunctionName :: Finished at [$(Get-Date)] on $Server" -Verbose:$True
       	Return $ServerStatusAfter
    }
}
