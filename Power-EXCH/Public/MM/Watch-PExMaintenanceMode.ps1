Function Watch-PExMaintenanceMode
{
	
<#
.SYNOPSIS
	Watch Maintenance mode items for all servers.
.DESCRIPTION
	This function displays Exchange Maintenance Mode items in summary.
.PARAMETER DelaySec
    Configurable delay between loop itterations.
.EXAMPLE
	PS C:\> Watch-PExMaintenanceMode
.NOTES
	Author      :: @ps1code
	Dependency  :: Function     :: Get-PExMaintenanceMode
    Version 1.0 :: 21-Aug-2024  :: [Release] :: Beta -CoadMonkey
    Version 2.0 :: 29-Aug-2025  :: [Improvement] :: Improve output and streamline. Added configurable delay.

.LINK

#>

    Param (

        # Sets how often the loop will repeat
        [Parameter( Mandatory = $False,
				    Position = 0
				    )]
            [int]
            $DelaySec = 10

    )

	
	Begin
	{
        $FunctionName = '{0}' -f $MyInvocation.MyCommand
	}

	Process
	{
        While ($true) {

    		Write-Verbose "$FunctionName :: Started at [$(Get-Date)] on $Server" -Verbose:$True

            Write-Verbose "Executing Get-ExchangeServer"
            $ExchangeServers = Get-ExchangeServer | sort Name
            If (!($ExchangeServers)) {
                Write-Error "Unable to get Exchange Servers" -Verbose:$True
                throw "Unable to get Exchange Servers"
            }


            ### Test Connection ###
            Write-Verbose "Executing Test-Connection"
            Write-Host -ForegroundColor Yellow "Ping Test"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.name) {
                $Online = Test-Connection $Server -Count 1 -Quiet
                $Object = New-Object Psobject -Property @{
                    Name = $Server
                    Online = $Online
                }
                $Obj_Arr += $Object
                If (!($Online)) {
                    $ExchangeServers = $ExchangeServers |? {$_.name -ne $Server}
                }
                    
            }
            $Obj_Arr|ft Name,Online


		    ### Hub Transport ###
            Write-Verbose "Executing Get-ServerComponentState"
            Write-Host -ForegroundColor Yellow "Hub Tansport status"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.name) {
                $Obj_Arr += Get-ServerComponentState -Identity $Server -Component HubTransport
            }
            $Obj_Arr|ft

        
		    ### Queue summary ###
            Write-Verbose "Executing Get-Queue"
            Write-Host -ForegroundColor Yellow "Mail Queue status"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.name) {
                $Object = New-Object Psobject -Property @{
                    Name = $Server
                    Queue = (Get-Queue -Server $Server | Measure-Object -Property MessageCount -Sum).Sum
                }
                $Obj_Arr += $Object
            }
            $Obj_Arr|ft Name,Queue


		    ### Cluster Nodes ###
            Write-Verbose "Executing Get-ClusterNode"
            Write-Host -ForegroundColor Yellow "Cluster Node status"
            If ($RunLocal)
            {
                Get-ClusterNode |ft Name,State
            }
            Else
            {
                Invoke-Command -ComputerName $ExchangeServers[0].name -ScriptBlock {
                    Get-ClusterNode |ft Name,State
                }
            }   


            ### Mailbox Database Copy Status ###
            Write-Verbose "Executing get-mailboxdatabasecopystatus"
            Write-Host -ForegroundColor Yellow "Mailbox Database Copy status"
            $Databases = get-mailboxdatabasecopystatus *|sort ActiveDatabaseCopy,Name
            foreach ($database in $databases) {
                if ($database.DatabaseSeedStatus) {
                    $database| Add-Member –MemberType NoteProperty –Name "Seed%" –Value $database.DatabaseSeedStatus.split(';').split(':')[1]
                }
            }
            $Databases|ft Name,@{l="Pref.";e={$_.ActivationPreference}},@{l="Active";e={$_.ActiveDatabaseCopy}},AutoActivationPolicy,@{l="Dis.&Move";e={$_.ActivationDisabledAndMoveNow}},Status,@{l="Index";e={$_.ContentIndexState}},@{l="Queue";e={$_.CopyQueueLength}},@{l="Disk%";e={$_.DiskFreeSpacePercent}},Seed% -auto

            # Warn if any are unhealthy
            If ($Databases|? {($_.status) -notlike "*Healthy*" -and ($_.status) -notlike "*Mounted*"}) {
                Write-Host -ForegroundColor Red "Problem database(s) were found!"
                Sound-Warning.ps1
            } Else {
                Write-Host -ForegroundColor Green "All Mailbox database copies are healthy!"
            }


<#		    ### Component States ###
            Write-Verbose "Executing Get-PExMaintenanceMode"
            Write-Host -ForegroundColor Yellow "PExMaintenceMode status"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.Name) {
                $Obj_Arr += Get-PExMaintenanceMode $Server
            }
            $Obj_Arr|ft
#>








# Output









            Write-Verbose "Executing Start-Sleep"
            Start-Sleep $DelaySec

            Write-Host "`n`n`n`n`n"
        }
    }
	End
    {
    }
}
