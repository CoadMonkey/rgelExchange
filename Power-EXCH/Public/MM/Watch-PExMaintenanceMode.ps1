Function Watch-PExMaintenanceMode
{
	
<#
.SYNOPSIS
	Watch Maintenance mode items for all servers.
.DESCRIPTION
	This function displays Exchange Maintenance Mode items in summary.
.PARAMETER DelaySec
    Configurable delay between loop itterations. Default 10 sec.
.PARAMETER AdditionalDNSNamespaces
    Addtional Namespaces to test DNS resolution.
.EXAMPLE
	PS C:\> Watch-PExMaintenanceMode
.NOTES
	Author      :: @ps1code
	Dependency  :: Function     :: Get-PExMaintenanceMode
    Version 1.0 :: 21-Aug-2024  :: [Release]     :: Beta -CoadMonkey
    Version 2.0 :: 28-Aug-2025  :: [Improvement] :: Improve output and streamline. Added configurable delay.
    Version 2.1 :: 29-Aug-2025  :: [Improvement] :: Improve output by doing all processing first and combining output objects.
    Version 2.2 :: 01-Sep-2025  :: [Improvement] :: Small output improvements.
    Version 2.3 :: 04-Sep-2025  :: [Improvement] :: Output improvements (GitHub Issues #6,7,9)
    Version 2.4 :: 28-Oct-2025  :: [Improvement] :: Add Namespace DNS checks.
    Version 2.5 :: 20-Mar-2026  :: [Improvement] :: Added parameter for additional DNS Namespaces. Added support for multiple DAGs in the org.
    version 2.6 :: 24-Jun-2026  :: [Improvement] :: Better handling for offline servers.

.LINK

#>

    Param (

        # Sets how often the loop will repeat
        [Parameter( Mandatory = $False,
				    Position = 0
				    )]
            [int]
            $DelaySec = 10,

        # Addtional Namespaces to test DNS resolution
        [Parameter ( Mandatory = $False)]
            [string[]]
            $AdditionalDNSNamespaces

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

            ## Update Server ResponseTimes
            foreach ($Server in $ExchangeServers) {
                Write-Verbose "Executing Test-Connection $($Server.name)"
                $ResponseTime = (Test-Connection -ComputerName $Server.name -Count 1).ResponseTime
                If ($ResponseTime) {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value $ResponseTime -Force
                } Else {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value 999 -Force
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
            foreach ($DNSNameSpace in $AdditionalDNSNamespaces) {
                $Object = New-Object Psobject -Property @{
                    DNSNameSpace = $DNSNameSpace
                    IPAddress = $Null
                }
                $DNSNameSpaces += $Object
            }
            [Array]$DNSNameSpaces = $DNSNameSpaces|sort DNSNameSpace -Unique
            IF ($Host.Name -notlike "*ISE*") {spin}
                        

		    ### Cluster Nodes ###
            $ClusterNodeArray = @()
            $DAGs = Get-DatabaseAvailabilityGroup
            foreach ($Dag in $DAGs)
            {
                Write-Verbose "Executing Get-ClusterNode"
                If ($RunLocal -and $env:COMPUTERNAME -in ($Dag.Servers))
                {
                    $ClusterNodeArray += Get-ClusterNode
                }
                Else
                {
                    $ClusterNodeArray += Invoke-Command -ComputerName (($ExchangeServers|? {$_.name -in $Dag.Servers}| sort ResponseTime)[0]).name  -ScriptBlock {
                        Get-ClusterNode
                    }
                }   
                IF ($Host.Name -notlike "*ISE*") {spin}
            }


            ## Update Server ResponseTimes
            foreach ($Server in $ExchangeServers) {
                Write-Verbose "Executing Test-Connection $($Server.name)"
                $ResponseTime = (Test-Connection -ComputerName $Server.name -Count 1).ResponseTime
                If ($ResponseTime) {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value $ResponseTime -Force
                } Else {
                    $Server | Add-Member -MemberType NoteProperty -Name ResponseTime -Value 999 -Force
                }
                IF ($Host.Name -notlike "*ISE*") {spin}
            }


            ### Server Checks ###
            foreach ($Server in $ExchangeServers)
            {

                ### Sanitize Variables ###
                $OWAURLs = @()
                Remove-Variable Object,HubTransport,Queue,MaintMode,ResponseTime,DNSNamespace -ErrorAction SilentlyContinue

                ### Initialize Output object ###
                Write-Verbose "Initializing output object"
                $Object = New-Object Psobject -Property @{
                    Name = $Server.Name
                    "Time(ms)" = $Server.ResponseTime
                    HubTransport = $Null
                    Queue = $Null
                    "MaintMode" = $Null
                    Cluster = $Null
                    # $DNSNameSpaces are added later
                }
                IF ($Host.Name -notlike "*ISE*") {spin}


                If ($Server.ResponseTime -lt 60) {

		            ### Hub Transport ###
                    Write-Verbose "Executing Get-ServerComponentState"
                    $HubTransport = (Get-ServerComponentState -Identity $Server.Name -Component HubTransport).State
                    IF ($Host.Name -notlike "*ISE*") {spin}

        
		            ### Queue totals ###
                    Write-Verbose "Executing Get-Queue"
                    $Queue = (Get-Queue -Server $Server.Name -ErrorAction SilentlyContinue | Measure-Object -Property MessageCount -Sum).Sum
                    IF ($Host.Name -notlike "*ISE*") {spin}


		            ### Maintenance Mode ###
                    Write-Verbose "Executing Get-PExMaintenanceMode"
                    $MaintMode = Get-PExMaintenanceMode $Server.Name
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
                                ((Resolve-DnsName -Name $Using:DNSNameSpace.DNSNameSpace -DnsOnly)).IPAddress
                            }
                        }
                        $Object | Add-Member -MemberType NoteProperty -Name $DNSNameSpace.DNSNameSpace -Value $DNSNameSpace.IPAddress -Force
                        IF ($Host.Name -notlike "*ISE*") {spin}                
                    }
                }


                ### Update Output object ###
                Write-Verbose "Updating output object"
                #$Object.Name = $Server.Name
                #$Object."Time(ms)" = $Server.ResponseTime
                $Object.HubTransport = $HubTransport
                $Object.Queue = $Queue
                $Object."MaintMode" = "$($MaintMode.state)($($MaintMode.TotalActiveComponent))"
                $Object.Cluster = "Pending..."

                IF ($Host.Name -notlike "*ISE*") {spin}
                

                $Obj_Arr += $Object
            }

            
            ### Add ClusterNode info to output array ###
            Write-Verbose "Adding Cluster Node info to output array"
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
