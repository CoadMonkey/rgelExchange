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
    Version 2.2 :: 01-Sep-2025  :: [Improvement] :: Small output improvements.
    Version 2.3 :: 04-Sep-2025  :: [Improvement] :: Output improvements (GitHub Issues #6,7,9)
    Version 2.4 :: 28-Oct-2025  :: [Improvement] :: Add Namespace DNS checks.

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
        Function Spin {
            New-Variable -Scope Global -Name SpinCounter -Description "This is the Global counter variable for Spin function." -ErrorAction SilentlyContinue
            $spin="/-\|"
            Write-Host "`b$($spin.Substring($Global:SpinCounter++%$spin.Length,1))" -nonewline
        } #End Function Spin

        $FunctionName = '{0}' -f $MyInvocation.MyCommand

	}

	Process
	{
        While ($true) {

    		Write-Verbose "$FunctionName :: Started at [$(Get-Date)]" -Verbose:$True


            ### Sanitize Variables ###
            Write-Verbose "Sanitizing Variables"
            $Obj_Arr = @()
            Remove-Variable ExchangeServers,Server,ResponseTime,RunLocal,OWAURLs,ClusterNodeArray,OLAnywhereHostnames,DNSNameSpaces,Object -ErrorAction SilentlyContinue
            IF ($Host.Name -notlike "*ISE*") {spin}


            ### Gather server information ###
            Write-Verbose "Executing Get-ExchangeServer"
            $ExchangeServers = Get-ExchangeServer | sort Name
            If (!($ExchangeServers)) {
                Write-Error "Unable to get Exchange Servers" -Verbose:$True
                throw "Unable to get Exchange Servers"
            }
            IF ($Host.Name -notlike "*ISE*") {spin}

            foreach ($Server in $ExchangeServers) {
                Write-Verbose "Executing Test-Connection $Server"
                $ResponseTime = (Test-Connection -ComputerName $Server.name -Count 1).ResponseTime
                If ($ResponseTime) {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value $ResponseTime -Force
                } Else {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value $False -Force
                }
                IF ($Host.Name -notlike "*ISE*") {spin}
            }

            $RunLocal = $False
            If ($env:COMPUTERNAME -in $ExchangeServers.name) { $RunLocal = $True }


            ### DNS Namespace(s) ###
            Write-Verbose "Executing Get-OutlookAnywhere"
            $OLAnywhereHostnames = Get-OutlookAnywhere -Server (($ExchangeServers| sort ResponseTime)[0]).name|select *ternalHostname
            $DNSNameSpaces = @()
            $Object = New-Object Psobject -Property @{
                DNSNameSpace = $OLAnywhereHostnames.ExternalHostname
                IPAddress = $Null
            }
            $DNSNameSpaces += $Object
            $Object = New-Object Psobject -Property @{
                DNSNameSpace = $OLAnywhereHostnames.InternalHostname
                IPAddress = $Null
            }
            $DNSNameSpaces += $Object
            [Array]$DNSNameSpaces = $DNSNameSpaces|sort DNSNameSpace -Unique
            IF ($Host.Name -notlike "*ISE*") {spin}
                        

		    ### Cluster Nodes ###
            Write-Verbose "Executing Get-ClusterNode"
            If ($RunLocal)
            {
                $ClusterNodeArray = Get-ClusterNode
            }
            Else
            {
                $ClusterNodeArray = Invoke-Command -ComputerName (($ExchangeServers| sort ResponseTime)[0]).name -ScriptBlock {
                    Get-ClusterNode
                }
            }   
            IF ($Host.Name -notlike "*ISE*") {spin}


            ### Server Checks ###
            foreach ($Server in $ExchangeServers | Where-Object {$_.ResponseTime -ne $False})
            {


                ### Sanitize Variables ###
                $OWAURLs = @()
                Remove-Variable Object,HubTransport,Queue,MaintMode,ResponseTime -ErrorAction SilentlyContinue


		        ### Hub Transport ###
                Write-Verbose "Executing Get-ServerComponentState"
                $HubTransport = (Get-ServerComponentState -Identity $Server.Name -Component HubTransport).State
                IF ($Host.Name -notlike "*ISE*") {spin}

        
		        ### Queue totals ###
                Write-Verbose "Executing Get-Queue"
                $Queue = (Get-Queue -Server $Server.Name | Measure-Object -Property MessageCount -Sum).Sum
                IF ($Host.Name -notlike "*ISE*") {spin}


		        ### Maintenance Mode ###
                Write-Verbose "Executing Get-PExMaintenanceMode"
                $MaintMode = Get-PExMaintenanceMode $Server.Name
                IF ($Host.Name -notlike "*ISE*") {spin}


                ### Build Output object ###
                Write-Verbose "Building output object"
                $Object = New-Object Psobject -Property @{
                    Name = $Server.Name
                    "Time(ms)" = $Server.ResponseTime
                    HubTransport = $HubTransport
                    Queue = $Queue
                    "MaintMode" = "$($MaintMode.state)($($MaintMode.TotalActiveComponent))"
                    Cluster = "Pending..."
                }
                IF ($Host.Name -notlike "*ISE*") {spin}


                ### Name Resolution ###
                foreach ($DNSNameSpace in $DNSNameSpaces) {
                    Write-Verbose "Executing Resolve-DnsName"
                    If ($RunLocal)
                    {
                        $DNSNameSpace.IPAddress = (Resolve-DnsName -Name $DNSNameSpace.DNSNameSpace -DnsOnly)[0]
                    }
                    Else
                    {
                        $DNSNameSpace.IPAddress = Invoke-Command -ComputerName $Server.Name -ScriptBlock {
                            ((Resolve-DnsName -Name $Using:DNSNameSpace.DNSNameSpace -DnsOnly)[0]).IPAddress
                        }
                    }
                    $Object | Add-Member -MemberType NoteProperty -Name $DNSNameSpace.DNSNameSpace -Value $DNSNameSpace.IPAddress
                    IF ($Host.Name -notlike "*ISE*") {spin}                
                }


                $Obj_Arr += $Object
            }

            
            ### Add ClusterNode info to output array ###
            foreach ($Server in $ExchangeServers.name)
            {
                ($Obj_Arr|Where-Object {$_.Name -eq $Server}).Cluster = ($ClusterNodeArray|Where-Object {$_.Name -eq $Server}).State
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
                IF ($Host.Name -notlike "*ISE*") {spin}
            }


            ### Output ###
            Write-Host "`b " -NoNewline     # Clear the spinner

            # Servers
            $Obj_Arr|ft Name,"Time(ms)",HubTransport,@{l="MsgQueue";e={$_.Queue}},Cluster,MaintMode,*.*
                                                                                
            # Databases
            $Databases|ft Name,@{l="Active";e={$_.ActiveDatabaseCopy}},@{l="ActivationPolicy";e={$_.AutoActivationPolicy}},@{l="Pref";e={$_.ActivationPreference}},@{l="Dis&Move";e={$_.ActivationDisabledAndMoveNow}},Status,@{l="IndexState";e={$_.ContentIndexState}},@{l="CpQueue";e={$_.CopyQueueLength}},@{l="Disk%";e={$_.DiskFreeSpacePercent}},Seed% -auto
            # Warn if any DBs are unhealthy
            If ($Databases|? {($_.status) -notlike "*Healthy*" -and ($_.status) -notlike "*Mounted*"})
            {
                Write-Host -ForegroundColor Red "Problem database(s) were found!"
            }


            ### Sleepy time ###
    		Write-Verbose "$FunctionName :: Sleeping $DelaySec seconds at [$(Get-Date)]" -Verbose:$True
            $a = (Get-Date).AddSeconds($DelaySec)
            While ((Get-Date) -le $a) { 
                spin
                sleep -Milliseconds 100
            }
            Write-Host "`b "

            Write-Host "`n`n`n`n`n"
        }
    }
	End
    {
    }
}
