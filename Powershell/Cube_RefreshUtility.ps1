param([string]$RefreshVersion = 'A', [string]$RefreshDim = 'Y', [string]$ConfigPath = '.\Config\config_pnl_refresh.json')

if (-not $ConfigPath.Contains(":")) {
    $ConfigPath = ( -Join ($PSScriptRoot, '\', $ConfigPath.TrimStart(".\")))
}

# refreshVersion -> A = 2 period of actuals; F = current forecast; B = current budget; Full = full refresh
# refreshDim -> O/Y/N; O = only, Y = yes (dims and fact(s) will be refreshed), N = no

#region Config
#### Changes possible below, but rarely ####
$VersionsList = @{
    'A'    = 'Actuals'
    'F'    = 'Forecast'
    'B'    = 'Budget'
    'Full' = 'Full'
}

$PeriodsList = @{
    'M' = 'P'
    'Q' = 'Q'
}
#### Changes possible above, but rarely ####

$folderDate = (Get-Date).ToString('MM-dd-yyyy')
$logFolder = ( -Join ($PSScriptRoot, '\Logs\', $folderDate, '\CubeRefresh\'))
New-Item -ItemType Directory -Force -Path $logFolder | Out-Null

$fileTime = (Get-Date).ToString('MM-dd-yyyy_HH-mm-ss')
$logPath = ( -Join ($logFolder, '\Cube_Refresh_Log_', $fileTime , '.txt'))
$log = ''

# Read config file (if available)
if ($ConfigPath.Length -gt 0) {
    if ( -not [IO.File]::Exists($ConfigPath) ) {
        $log = ( -Join ('The file "', $ConfigPath, '" was not found; using default values instead', "`n"))
    }
    else {
        try {
            $configParams = Get-Content -Path $ConfigPath | ConvertFrom-Json
            if ($configParams.Disabled) {
                $log = ( -Join (((Get-Date).ToString('MM/dd/yyyy HH:mm:ss')), ': Cube refresh script has been disabled in the file "', $ConfigPath , '"; quitting'))

                Write-Output $log
                Write-Output $log | Out-File -FilePath $logPath

                exit
            }
        }
        catch {
            $log = ( -Join ('Problems reading the config file: "', $ConfigPath, '"; quitting'))

            Write-Output $log
            Write-Output $log | Out-File -FilePath $logPath
            exit
        }
    }
}

# Cube related attributes

# One or more cube servers can be defined
$cubeServersL = @('madev02')


# Name of the cube
$cubeNamesL = @('MercuryTest')


# The main fact tables(s) to be refreshed.
# The first fact table should be the primary,
# and will be queried when determine current forecast
$cubeFactTableNameL = @('MercuryJim')


# A list of objects to ignore for refreshing
$cubeIgnoreObjectsL = @()


# Can be Q for quarters; M for months
$periods = $PeriodsList['M']


# Corporate DRS server, to get current forecast information
$drsServerL = 'mercuryproddrs'


# Verbose logging
$vloggingL = $false


# Current year actuals template
# Prior year actuals template
# Current forecast template
# Prior forecast template
# Current budget template
[System.Collections.ArrayList]$actualsTemplate = @()
$actualsTemplate.Add( -Join ('ACT_CY_', $periods , '{0}')) | Out-Null
$actualsPriorTemplate = ( -Join ('ACT_PY_', $periods , '{0}'))

[System.Collections.ArrayList]$forecastTemplate = @()
$forecastTemplate.Add( -Join ('FCST_', $periods , '{0}')) | Out-Null
$forecastPriorTemplate = ( -Join ('FCST_H_', $periods , '{0}'))

[System.Collections.ArrayList]$budgetTemplate = @()
$budgetTemplate.Add( -Join ('BUD_', $periods , '{0}')) | Out-Null
#endregion

########################  MAKE ALL CHANGES ABOVE  ########################

function Get-RefreshAllFC {
    Param (
        $serverName,
        $cubeName,
        $cubeFactTableName,
        $mercuryForecastId
    )

    $DAX_Query = ( -Join (
            'EVALUATE
        ROW (
            "RC", COUNTROWS (
                FILTER (
                    SUMMARIZECOLUMNS (
                        ', $cubeFactTableName , '[VersionID],
                        FILTER ( ', $cubeFactTableName , ', NOT LEFT ( ', $cubeFactTableName , '[VersionID], 2 ) IN { "13", "14" } )),
                    SEARCH ( "', $mercuryForecastId , '-", ', $cubeFactTableName , '[VersionID], 1, 0 )
                )
            )
        )'))

    # Load Assembly
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.AnalysisServices') | Out-Null

    $conn = New-Object System.Data.OleDb.OleDbConnection
    $conn.ConnectionString = ( -Join ('Provider=MSOLAP;Data Source=', $serverName, ';Catalog=', $cubeName, ' ;'))

    try { $conn.Open() } catch { }

    if (  $conn.State -eq 'Closed') {
        $log = ( -Join ('** FAILURE ** Not connected to cube database; server name -> ', ([string]$serverName).ToUpper() , '; cube name -> ', $cubeName, '; quitting', "`n"))

        Write-Output $log
        Write-Output $log | Out-File -FilePath $logPath

        [System.Environment]::Exit(1)

        exit
    }

    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $DAX_Query
    [string]$retRows = $cmd.ExecuteScalar()

    $conn.Close()

    return $retRows
}

$RefreshCubeBlock = {
    param (
        $ServerName,
        $RefreshParameters
    )
    # Load Assembly
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.AnalysisServices') | Out-Null

    [datetime]$startTime = Get-Date
    $logMsg = ( -Join ("`n", $ServerName, ' - Cube Refresh Start Time: ', ($startTime.ToString('MM/dd/yyyy HH:mm:ss'))))

    # Save pass parameters, to local variables
    $refreshDim = $RefreshParameters.RefreshDim
    $vlogging = if (-not $RefreshParameters.Process ) { $true } else { $RefreshParameters.VerboseLogging }

    if ($refreshDim -ne 'o') { $logMsg += ( -Join ("`n", $ServerName, ' - Refresh Version: ', $RefreshParameters.RefreshVersion)) }
    if ($refreshDim -eq 'o' ) { $logMsg += ( -Join ("`n", $ServerName, ' - ** Only Refreshing Dims **')) }
    $logMsg += "`n"


    # Connect to Tabular SSAS
    $srv = New-Object Microsoft.AnalysisServices.Server
    try {
        $srv.connect($ServerName)
    }
    catch {
        $logMsg = ( -Join ("`n", $serverName, ' - ** FAILURE ** Not connected to cube server; quitting'))
        $logMsg
        exit
    }

    $RefreshParameters.Cube.GetEnumerator() | ForEach-Object {
        # Connect to database (cube)
        $cubeName = $_
        $db = $srv.Databases.FindByName($cubeName);

        if ( $null -eq $db) {
            if ($srv.Connected) { $srv.Disconnect() }
            $logMsg = ( -Join ("`n", $ServerName, ' - ** FAILURE ** Not connected to cube database: ', $cubeName, '; quitting', "`n"))
            $logMsg
            exit
        }

        # The model for the cube
        $model = $db.model
        if ( $null -eq $model) {
            if ($srv.Connected) { $srv.Disconnect() }
            $logMsg = ( -Join ("`n", $ServerName, ' - ** FAILURE ** Not connected to cube model; quitting', "`n"))
            $logMsg
            exit
        }

        # If full refresh, iterate all tables and all partitions
        if ($RefreshParameters.FullRefresh) {
            $model.Tables.GetEnumerator() | ForEach-Object {
                $tbl = $_
                $loopCnt = 1
                $tbl.Partitions.GetEnumerator() | ForEach-Object {
                    if (-not $RefreshParameters.IgnoreObjects.Contains($_.Name) -and -not( $RefreshParameters.CubeFact.Contains($tbl.Name) -and $refreshDim -eq 'o' ) ) {
                        if ($RefreshParameters.Process) { $_.RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full) }
                        if ($vlogging -and $loopCnt -eq 1) { $logMsg += ( -Join ("`n", 'Refreshing:  table: ', $tbl.Name, '; cube: ', $cubeName, ' on ', $srv.Name)) }
                        $loopCnt++
                    }
                }
            }
        }
        else {
            $model.Tables.GetEnumerator() | ForEach-Object {
                $tbl = $_
                if ( -not $RefreshParameters.CubeFact.Contains($tbl.Name) -and -not $RefreshParameters.IgnoreObjects.Contains($tbl.Name) ) {
                    $tbl.Partitions.GetEnumerator() | ForEach-Object {
                        if ($refreshDim -eq 'o' -or $refreshDim -eq 'y' ) {
                            if ($vlogging) { $logMsg += ( -Join ("`n", 'Refreshing:  table: ', $tbl.Name, '; cube: ', $cubeName, ' on ', $srv.Name)) }
                            if ($RefreshParameters.Process) { $_.RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full) }
                        }
                    }
                }

                # We can prevent the refreshing of the fact (and only refresh other objects)
                if ( $RefreshParameters.CubeFact.Contains($tbl.Name) -and -not ($RefreshParameters.IgnoreObjects.Contains($_.Name)) -and $refreshDim -ne 'o' ) {
                    $tbl.Partitions.GetEnumerator() | ForEach-Object {
                        if ($RefreshParameters.CubeTemplate.Contains($_.Name) -and -not ($RefreshParameters.IgnoreObjects.Contains($_.Name)) ) {
                            if ($vlogging) { $logMsg += ( -Join ("`n", 'Refreshing:  table: ', $tbl.Name , '; partition: ', $_.Name, '; cube: ', $cubeName, ' on ', $srv.Name)) }
                            if ($RefreshParameters.Process) { $_.RequestRefresh([Microsoft.AnalysisServices.Tabular.RefreshType]::Full) }
                        }
                    }
                }
            }
        }

        try {
            if ($RefreshParameters.Process) { $db.Update([Microsoft.AnalysisServices.UpdateOptions]::ExpandFull) }
        }
        catch {
            $logMsg = ( -Join ("`n", $ServerName, ' Cube: ', $cubeName, ' - ** FAILURE ** while refreshing cube; ', $_.Exception.Message ))
            $logMsg
            exit
        }
    }

    if ($srv.Connected) { $srv.Disconnect() }

    [datetime]$endTime = Get-Date
    $logMsg += ( -Join ("`n`n", $ServerName, ' - Cube Refresh End Time: ', ($endTime.ToString('MM/dd/yyyy HH:mm:ss'))))
    $logMsg += ( -Join ("`n", $ServerName, ' - Elapsed Time: ', (New-TimeSpan -Start $startTime -End $endTime) , "`n"))
    $logMsg
}

#region Final settings
[System.Collections.ArrayList]$cubeServers = @()
if ($null -eq $configParams.cubeServers ) { $cubeServersL.GetEnumerator() | ForEach-Object { $cubeServers.Add($_) } | Out-Null } else { $configParams.cubeServers.GetEnumerator() | ForEach-Object { $cubeServers.Add($_) } | Out-Null }


[System.Collections.ArrayList]$cubeNames = @()
if ($null -eq $configParams.cubeNames ) { $cubeNamesL.GetEnumerator() | ForEach-Object { $cubeNames.Add($_) } | Out-Null } else { $configParams.cubeNames.GetEnumerator() | ForEach-Object { $cubeNames.Add($_) } | Out-Null }


[System.Collections.ArrayList]$cubeFactTableName = @()
if ($null -eq $configParams.cubeFactTableName ) { $cubeFactTableNameL.GetEnumerator() | ForEach-Object { $cubeFactTableName.Add($_) } | Out-Null } else { $configParams.cubeFactTableName.GetEnumerator() | ForEach-Object { $cubeFactTableName.Add($_) } | Out-Null }


[System.Collections.ArrayList]$cubeIgnoreObjects = @()
if ($null -eq $configParams.cubeIgnoreObjects ) { $cubeIgnoreObjectsL.GetEnumerator() | ForEach-Object { $cubeIgnoreObjects.Add($_) } | Out-Null } else { $configParams.cubeIgnoreObjects.GetEnumerator() | ForEach-Object { $cubeIgnoreObjects.Add($_) } | Out-Null }

$drsServer = if ($null -eq $configParams.drsServer ) { $drsServerL } else { $configParams.drsServer }
$vlogging = if ($null -eq $configParams.verboseLogging ) { $vloggingL } else { $configParams.verboseLogging }


$vR = if ( @('a', 'f', 'b', 'full').Contains( $RefreshVersion.ToLower() ) ) { $RefreshVersion.ToUpper() } else { 'A' }
$versionRefresh = $VersionsList[ $vR ]

# Current month is based on current calendar date
$currentMonth = (Get-Date).Month


# Mapping current month number to fiscal months/fiscal quarters for refreshing
switch ($currentMonth) {
    1 { $monthNums = @(6..7); $quarterNums = @(2, 3); break }
    2 { $monthNums = @(7..8); $quarterNums = @(3); break }
    3 { $monthNums = @(8..9); $quarterNums = @(3); break }
    4 { $monthNums = @(9..10); $quarterNums = @(3, 4); break }
    5 { $monthNums = @(10..11); $quarterNums = @(4); break }
    6 { $monthNums = @(11..12); $quarterNums = @(4); break }
    7 { $monthNums = @(12, 1); $quarterNums = @(1, 4); break }
    8 { $monthNums = @(1..2); $quarterNums = @(1); break }
    9 { $monthNums = @(2..3); $quarterNums = @(1); break }
    10 { $monthNums = @(3..4); $quarterNums = @(1, 2); break }
    11 { $monthNums = @(4..5); $quarterNums = @(2); break }
    Default { $monthNums = @(5..6); $quarterNums = @(2); break }
}

# If month is July, we need to refresh the partition, for the prior period
if ($currentMonth -eq 7) {
    $actualsTemplate.Add($actualsPriorTemplate) | Out-Null
}


# If current forecast (as defined by corp), is not contained in the current forecast partition, refresh the prior forecast partition
# DRS Query to get current forecast version
$DRS_Query = 'SELECT [ForecastVersionId] FROM [Forecast]..[ForecastVersion] WHERE [CurrentForecastVersionFlagId] = 1'

$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = ( -Join ('Data Source=', $drsServer, ';Database=Mars;Integrated Security=SSPI;'))

try { $conn.Open() } catch { }

if ( $conn.State -eq 'Closed') {
    $log = ( -Join ('** FAILURE ** Not connected to DRS database; quitting', "`n"))

    Write-Output $log
    Write-Output $log | Out-File -FilePath $logPath
    exit
}

$cmd = $conn.CreateCommand()

$cmd.CommandText = $DRS_Query
$currentFC = $cmd.ExecuteScalar()
$conn.Close()

# Assume first fact table, is the primary
if ($null -ne $cubeServers[0] -and $null -ne $cubeNames[0] -and $null -ne $cubeFactTableName[0]) {
    $retRows = Get-RefreshAllFC -serverName $cubeServers[0] -cubeName $cubeNames[0] -cubeFactTableName $cubeFactTableName[0] -mercuryForecastId $currentFC
}
else {
    # No historical if a missing attributes
    $retRows = 1
}

$priorFC = if ($retRows.Length -eq 0) { $true } else { $false }
if ($priorFC) {
    $forecastTemplate.Add($forecastPriorTemplate) | Out-Null
}


$fullRefresh = if ($versionRefresh -eq $VersionsList['Full']) { $true } else { $false }

# If refreshing actuals, we only need to refresh partitions for certain periods
if ($versionRefresh -eq $VersionsList['A']) {
    if ($periods -eq ($PeriodsList['M'])) {
        $periodNums = $monthNums
    }
    else {
        $periodNums = $quarterNums
    }
}
else {
    if ($periods -eq ($PeriodsList['M'])) {
        $periodNums = @(1..12)
    }
    else {
        $periodNums = @(1..4)
    }
}

if ( $versionRefresh -eq $VersionsList['A'] ) {
    $currentTemplate = $actualsTemplate
}
elseif ( $versionRefresh -eq $VersionsList['F'] ) {
    $currentTemplate = $forecastTemplate
}
else {
    $currentTemplate = $budgetTemplate
}

# If month is July, treat actuals a bit differently
[System.Collections.ArrayList]$cubeTemplate = @{ }
if ($versionRefresh -eq $versionslist['A'] -and $currentMonth -eq 7) {
    $cubeTemplate.Add($currentTemplate[0] -f $periodNums[1]) | Out-Null
    $cubeTemplate.Add($currentTemplate[1] -f $periodNums[0]) | Out-Null
}
else {
    $periodNums.GetEnumerator() | ForEach-Object {
        [string]$period = $_
        if ($period.Length -eq 1 -and $periods -ne $PeriodsList['Q']) {
            $period = -Join ('0', $period)
        }
        $currentTemplate.GetEnumerator() | ForEach-Object {
            $cubeTemplate.Add($_ -f $period) | Out-Null
        }
    }
}


# Refresh cube
# Only accept O, Y or N, otherwise default to Y
$dimRefresh = if (  @('o', 'y', 'n').Contains($RefreshDim.ToLower() )) { $RefreshDim.ToLower() } else { 'y' }
#endregion

$RefreshParameters = @{
    'Cube'           = $cubeNames                           # Name of the cube to refresh
    'CubeFact'       = $cubeFactTableName                   # An array of the main cube fact(s)
    'IgnoreObjects'  = $cubeIgnoreObjects                   # An array of objects to ignore, when refreshing
    'CubeTemplate'   = $cubeTemplate                        # List of partitions to refresh
    'RefreshDim'     = $dimRefresh                          # Refresh dims/fact partitions?
    'FullRefresh'    = $fullRefresh                         # Perform a full cube refresh?
    'RefreshVersion' = $versionslist[$refreshVersion]       # The version being refreshed
    'VerboseLogging' = $vlogging                            # Logging related to partitions and server?
    'ForecastId'     = $currentFC                           # The id of the current forecast
    'Process'        = $true                                # For testing purposes - normally 0
}

# Job section, to refresh cubes
$jobs = @()

$CubeServers.GetEnumerator() | ForEach-Object {
    $CubeServer = ([string]$_).ToUpper()

    $jobs += Start-Job -ScriptBlock $RefreshCubeBlock -ArgumentList $CubeServer, $RefreshParameters
}

Wait-Job -Job $jobs | Out-Null

# Error handling and restart if required
$errorOccurred = $true; $errorCntTotal = 1; $maxErrors = 3; $delayBetweenErrors = 300
while ($errorOccurred -and ($errorCntTotal -le $maxErrors)) {
    $errorCnt = 0; $items = @(); $i = 0; $sleep = $false

    $jobs.GetEnumerator() | ForEach-Object {
        Receive-Job -Job $_ -OutVariable logFinal | Out-Null
        if ( ([string]$logFinal) -like '*FAILURE*' ) {
            $log += ( -Join ("`n`n", 'Failure number ', $errorCntTotal))
            $sleep = $true
            $items += $i
            $errorCnt++
        }
        $log += $logFinal
        $i++
    }

    if ($jobs.Length -gt 0) {
        # Remove all jobs
        Remove-Job -Job $jobs -Force | Out-Null
    }

    if ($sleep) {
        Start-Sleep -Seconds $delayBetweenErrors
    }

    $jobs = @()
    [System.Collections.ArrayList]$CubeServers_New = @()
    $items.GetEnumerator() | ForEach-Object {
        $CubeServer = $CubeServers[$_].ToUpper()
        $jobs += Start-Job -ScriptBlock $RefreshCubeBlock -ArgumentList $CubeServer, $RefreshParameters

        $CubeServers_New.Add(@($CubeServer)) | Out-Null
    }

    $CubeServers = $CubeServers_New

    if ($errorCnt -eq 0) {
        $errorOccurred = $false
    }
    else {
        Wait-Job -Job $jobs | Out-Null
        $errorCntTotal++
    }
}

$fileTime = (Get-Date).ToString('MM-dd-yyyy_HH-mm-ss')
$logPath = ( -Join ($logFolder, '\Cube_Refresh_Log_', $fileTime , '.txt'))

Write-Output $log
Write-Output $log | Out-File -FilePath $logPath -Force

if ($errorOccurred) { [System.Environment]::Exit(1) }