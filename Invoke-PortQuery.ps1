<#
.SYNOPSIS
    Invoke-PortQuery
.DESCRIPTION
    Invoke-Puery query is used to test TCP and UDP port connectivity
.PARAMETER JSONPath
    Specifies the JSON path and file to process
.PARAMETER ServerList
    Array of servers to test for whether a port is listening or not
.PARAMETER Description
    Description of the server
    Example:  "Certificate Servers"
.PARAMETER ProtocolType
    Only TCP or UDP are valid options
.PARAMETER Ports
    Array of ports to check on the servers in the server list
.PARAMETER Log
    Whether or not to create .log file
    Default = True
.PARAMETER LogClixml
    Whether to create a clixml object, which is put in the logpath location
    Default = False
.PARAMETER LogPath
    Set to the path where you'd like the log files to reside
    Default = PSScriptRoot
.PARAMETER SuppressOutput
    Whether to supporess output of the data object (clixml)
    Default = True
.EXAMPLE
    Invoke-PortQuery -JSONPath <path>\<filename>.json
.EXAMPLE
    Invoke-PortQuery -ServerList MYSERVER02,MYSERVER01" -Description "Certificate Servers" -ProtocolType "TCP" -Ports "80,443,135"
.OUTPUTS
    Log file and clixml file
    Outputs jobobject hashtable
.NOTES
    Last Updated: 

    ========== HISTORY ==========
    Author: SweetestSufferance
    Created: 2021-04-02 11:01:42Z
    Package Version: 1.0.0.0 
        - Initial release
        - Opted out of the -q (quiet) option due to the results not being accurate for certain ports
#>
[CmdletBinding(DefaultParameterSetName = 'json')]
param (
    [Parameter(Mandatory = $true,
        ParameterSetName = 'json',
        HelpMessage = 'Enter path to json file',
        Position = 0)]
    [string[]]$JSONPath,
    [Parameter(Mandatory = $true,
        ParameterSetName = 'manual',
        HelpMessage = 'Enter a list of servers to check',
        Position = 0)]
    [Array[]]$ServerList,
    [Parameter(Mandatory = $true,
    ParameterSetName = 'manual',
    HelpMessage = 'Enter a description',
    Position = 1)]
    [string[]]$Description,
    [Parameter(Mandatory = $true,
        ParameterSetName = 'manual',
        HelpMessage = 'Set protocol to either TCP or UDP',
        Position = 2)]
    [string[]]$ProtocolType = "TCP",
    [Parameter(Mandatory = $true,
        ParameterSetName = 'manual',
        HelpMessage = 'Enter a list of ports to check',
        Position = 3)]
    [Array[]]$Ports,
    [Parameter(ParameterSetName = 'manual', HelpMessage = 'Set to "true" to enable logging')]
    [Parameter(ParameterSetName = 'json', HelpMessage = 'Set to "true" to enable logging')]
    $Log = $true,
    [Parameter(ParameterSetName = 'manual', HelpMessage = 'Set to "true" to export stream to clixml file')]
    [Parameter(ParameterSetName = 'json', HelpMessage = 'Set to "true" to export stream to clixml file')]
    $LogClixml = $false,
    [Parameter(ParameterSetName = 'manual', HelpMessage = 'Set path to override log file name and location')]
    [Parameter(ParameterSetName = 'json', HelpMessage = 'Set path to override log file name and location')]
    $LogPath = $PSScriptRoot,
    [Parameter(ParameterSetName = 'manual', HelpMessage = 'Set path to override log file name and location')]
    [Parameter(ParameterSetName = 'json', HelpMessage = 'Set path to override log file name and location')]
    $SuppressOutput = $true,
    [Parameter(ParameterSetName = 'manual', HelpMessage = 'Enter fully qualified domain name')]
    [Parameter(ParameterSetName = 'json', HelpMessage = 'Enter fully qualified domain name')]
    [string[]]$FQDN = "bsci.bossci.com"
)

try {
    $SiteName = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
}
catch{
    $SiteName= "NoDomain"
}
$DateTime = get-date -Format "yyyyMMddHHmmss"
$ComputerName = $env:COMPUTERNAME
$LogFileStub = "$($ComputerName)_$($SiteName)_$($DateTime)"
If (!(Test-Path -Path $LogPath)){
    Try{
        Write-Output "Attempting to create folder $($LogPath)"
        Set-Item -Path $LogPath -Force
        $LogFile = "$($LogPath)\$($LogFileStub).log"
    }
    Catch{
        Write-Warning -Message "Unable to folder $($LogPath)"
        Exit
    }
}
Else{
    $LogFile = "$($LogPath)\$($LogFileStub).log"
}

Switch ($PSCmdlet.ParameterSetName){
    "json" {
        If (Test-Path -Path $($JSONPath)){
            $ServerData = Get-Content "$($JSONPath)" -Raw | ConvertFrom-Json
        }
        Else{
            Write-Warning -Message "Please enter a valid path for the JSON file"
            Exit
        }
    }
    "manual" {
        If ($ProtocolType -notmatch "^[T,U][CP,DP].$"){
            Write-Warning -Message "$($ProtocolType) is not valid.  Please use either TCP or UDP"
            Exit
        }
        $ServerData = [PSCustomObject]@{"ServerList" = $ServerList; "Description" = $Description; "Protocols" = [PSCustomObject]@{"type" = $ProtocolType;"ports" = $Ports}}
    }
    "__AllParameterSets" {
        Write-Output "No parameter set name found, please select a valid parameter"
        Exit
    }
}

"Source: $ComputerName" | Out-File $LogFile -Append
$JobList = @()
$PortQueryPath = "$PSScriptRoot\Portqry.exe"

If(!(Test-Path $PortQueryPath)){
    Write-Warning "Portqry not found"
    Break
}
Else{
    $ServerData | ForEach-Object {
        $Description = $_.Description
        $Protocols = $_.Protocols
        $ServerList = $_.ServerList.Split(",")
        $Job = 
        {
            param(
                $ServerList,
                $Protocols,
                $PortQueryPath,
                $Description,
                $FQDN
            )
            $ServerList | ForEach-Object{
                $Server = "$($_).$FQDN"
                $Protocols | ForEach-Object {
                    $Ports = $_.ports.Split(",")
                    $Protocol = $_.type
                    $Ports | ForEach-Object {
                        $Port = $_
                        $ExecutionData = Invoke-Command -ScriptBlock{cmd.exe /c "$PortQueryPath -n $Server -e $Port -p $Protocol"}
                        $PortQuery = $ExecutionData | Where-Object {$_ -match "^[T,U][C,D][P]\ port.*"}
                        $Status = $PortQuery.Split(':')[1].Trim()
                        $Object = New-Object PSObject -Property ([ordered]@{
                            Description          = $Description.Trim('{}')
                            Server               = $Server.Trim('{}')
                            Protocol             = $Protocol.Trim('{}')
                            PortNumber           = $Port.Trim('{}')
                            Status               = $Status.Trim('{}')
                        })
                        Write-Output $Object
                    }
                }
            }
        }
        $JobID = Start-Job -Name "$Description" -ScriptBlock $Job -ArgumentList $ServerList,$Protocols,$PortQueryPath,$Description,$FQDN
        $JobList += $JobID.Id
    }
}

$JobsToProcess = Get-Job | Where-Object {$_.id -in $JobList}
Write-Output "$($JobsToProcess.Count)"
$JobObject = @()
While($JobsToProcess.Count){
    $JobsToProcess.Where{$_.State -eq 'Completed'} | ForEach-Object {
        $Data = $Null
        $JobId = $_.Id
        $JobName = $_.Name
        $JobState = $_.State
        $JobHasMoreData = $_.HasMoreData
        Write-Output "Processing of $JobName is $JobState"
        
        If ($JobState -eq "Completed" -and $JobHasMoreData -eq "True"){
            "Processing $($JobName)" | Out-File $LogFile -Append
            try{
                $Data = Receive-Job -Id $JobId | Select-Object -Property Description,Server,Protocol,PortNumber,Status
                If ($Log){
                    $Data | Select-Object -Property Server,Protocol,PortNumber,Status | Format-Table | Out-File -FilePath $LogFile -Append
                }
                $JobObject += $Data | ForEach-Object {$_}
                Remove-Job -Id $JobId
                $JobsToProcess = Get-Job | Where {$_.id -in $JobList}
            }
            catch{
                # To lazy to add anything here at the moment
            }
        }
        ElseIf($JobState -eq "Running"){
            # Probably not needed now that we're filtering off of completed state
        } 
        ElseIf($JobState -eq "Completed" -and $JobHasMoreData -eq "False"){
            # Not sure this section even works
            Write-Output "attempting to remove job id: $JobId"
            Remove-Job -Id $JobId
        }
        Else{
            Sleep 10
        }
    }
    $JobsToProcess = Get-Job | Where {$_.id -in $JobList}
}

If ($LogClixml){
    $JobObject | Export-Clixml -Path "$($LogPath)\$($LogFileStub).clixml"
}
If ($SuppressOutput){
    Return
}
Else{
    Return $JobObject
}


