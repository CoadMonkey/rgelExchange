Function Watch-PExMaintenanceMode
{
	
<#
.SYNOPSIS
	Watch Maintenance mode items for all servers.
.DESCRIPTION
	This function displays Exchange Maintenance Mode items in summary.
.PARAMETER Server
    Exchange server name or object representing Exchange server
.EXAMPLE
	PS C:\> Watch-PExMaintenanceMode
.NOTES
	Author      :: @ps1code
	Dependency  :: Function     :: Get-PExMaintenanceMode
    Version 1.0 :: 20-Aug-2024  :: [Release] :: Beta -CoadMonkey

.LINK

#>
	
	
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
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.name) {
                $Obj_Arr += Get-ServerComponentState -Identity $Server -Component HubTransport
            }
            $Obj_Arr|ft

        
		    ### Queue summary ###
            Write-Verbose "Executing Get-Queue"
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
            $Databases = get-mailboxdatabasecopystatus *|sort ActiveDatabaseCopy,Name
            foreach ($database in $databases) {
                if ($database.DatabaseSeedStatus) {
                    $database| Add-Member –MemberType NoteProperty –Name "Seed%" –Value $database.DatabaseSeedStatus.split(';').split(':')[1]
                }
            }
            $Databases|ft Name,@{l="Active";e={$_.ActiveDatabaseCopy}},Status,@{l="Index";e={$_.ContentIndexState}},@{l="Queue";e={$_.CopyQueueLength}},@{l="Disk%";e={$_.DiskFreeSpacePercent}},Seed% -auto


            ### DatabaseCopyActivationDisabledAndMoveNow,DatabaseCopyAutoActivationPolicy ###
            Write-Verbose "Executing Get-MailboxServer"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.Name) {
                $Obj_Arr += Get-MailboxServer $Server
            }
            $Obj_Arr|ft Name,DatabaseCopyActivationDisabledAndMoveNow,DatabaseCopyAutoActivationPolicy


		    ### Component States ###
            Write-Verbose "Executing Get-PExMaintenanceMode"
            $Obj_Arr = @()
            foreach ($Server in $ExchangeServers.Name) {
                $Obj_Arr += Get-PExMaintenanceMode $Server
            }
            $Obj_Arr|ft

            Write-Verbose "Executing Start-Sleep"
            Start-Sleep 10

            Write-Host "`n`n`n`n`n"
        }
    }
	End
    {
    }
}
