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
    Version 2.0 :: 28-Aug-2025  :: [Improvement] :: Improve output and streamline. Added configurable delay.
    Version 2.1 :: 29-Aug-2025  :: [Improvement] :: Improve output by doing all processing first and combining output objects.

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

    		Write-Verbose "$FunctionName :: Started at [$(Get-Date)]" -Verbose:$True

            ### Gather server information ###
            Write-Verbose "Executing Get-ExchangeServer"
            $ExchangeServers = Get-ExchangeServer | sort Name
            If (!($ExchangeServers)) {
                Write-Error "Unable to get Exchange Servers" -Verbose:$True
                throw "Unable to get Exchange Servers"
            }
            $RunLocal = $False
            If ($env:COMPUTERNAME -in $ExchangeServers.name) { $RunLocal = $True }


            ### Reset variables ###
            $Obj_Arr = @()
            Remove-Variable Object,Server,Online,HubTransport,Queue,MaintMode,ClusterNode -ErrorAction SilentlyContinue


            ### Server Checks ###
            foreach ($Server in $ExchangeServers.name)
            {


                ### Test Connection ###
                Write-Verbose "Executing Test-Connection"
                $Online = Test-Connection $Server -Count 1 -Quiet
                

                If ($Online)
                {


		            ### Hub Transport ###
                    Write-Verbose "Executing Get-ServerComponentState"
                    $HubTransport = (Get-ServerComponentState -Identity $Server -Component HubTransport).State

        
		            ### Queue totals ###
                    Write-Verbose "Executing Get-Queue"
                    $Queue = (Get-Queue -Server $Server | Measure-Object -Property MessageCount -Sum).Sum


		            ### Cluster Nodes ###
                    Write-Verbose "Executing Get-ClusterNode"
                    If (!($ClusterNodeArray))
                    {
                        If ($RunLocal)
                        {
                            $ClusterNodeArray = Get-ClusterNode
                        }
                        Else
                        {
                            $ClusterNodeArray = Invoke-Command -ComputerName $Server -ScriptBlock {
                                Get-ClusterNode
                            }
                        }   
                    }
                    $ClusterNode = ($ClusterNodeArray|Where-Object {$_.Name -eq $Server}).State


		            ### Maintenance Mode ###
                    Write-Verbose "Executing Get-PExMaintenanceMode"
                    $MaintMode = (Get-PExMaintenanceMode $Server).State

                }


                ### Build Output object ###
                Write-Verbose "Building output object"
                $Object = New-Object Psobject -Property @{
                    Name = $Server
                    Online = $Online
                    HubTransport = $HubTransport
                    Queue = $Queue
                    MaintMode = $MaintMode
                    ClusterNode = $ClusterNode
                }
                $Obj_Arr += $Object

            }


            ### Mailbox Database Copy Status ###
            Write-Verbose "Executing get-mailboxdatabasecopystatus"
            $Databases = get-mailboxdatabasecopystatus *|sort ActiveDatabaseCopy,Name
            foreach ($database in $databases)
            {
                if ($database.DatabaseSeedStatus)
                {
                    $database| Add-Member –MemberType NoteProperty –Name "Seed%" –Value $database.DatabaseSeedStatus.split(';').split(':')[1]
                }
            }


            ### Output ###
            $Obj_Arr|ft Name,Online,HubTransport,Queue,MaintMode,ClusterNode

            $Databases|ft Name,@{l="Active";e={$_.ActiveDatabaseCopy}},AutoActivationPolicy,@{l="Pref.";e={$_.ActivationPreference}},@{l="Dis.&Move";e={$_.ActivationDisabledAndMoveNow}},Status,@{l="Index";e={$_.ContentIndexState}},@{l="Queue";e={$_.CopyQueueLength}},@{l="Disk%";e={$_.DiskFreeSpacePercent}},Seed% -auto

            # Warn if any DBs are unhealthy
            If ($Databases|? {($_.status) -notlike "*Healthy*" -and ($_.status) -notlike "*Mounted*"})
            {
                Write-Host -ForegroundColor Red "Problem database(s) were found!"
            }
            Else
            {
                Write-Host -ForegroundColor Green "All Mailbox database copies are healthy!"
            }


            ### Sleepy time ###
            Write-Verbose "Executing Start-Sleep"
    		Write-Verbose "$FunctionName :: Sleeping $DelaySec seconds at [$(Get-Date)]" -Verbose:$True
            Start-Sleep $DelaySec

            Write-Host "`n`n`n`n`n"
        }
    }
	End
    {
    }
}
