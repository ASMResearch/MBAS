#region Cube related attributes

# Name of the cube server
$serverName = 'asazure://centralus.asazure.windows.net/mambas'


# Path to bim file
# If path is empty (''), then connect to server instead
$bimFilePath = ''
# $bimFilePath = 'C:\Users\jimb\OneDrive - Microsoft\MBAS\VS\PnL\PnL\Model.bim'

# Name of the cube
$cubeName = 'PnL'


# key = name of fact table; value = fact table attribute, that is passed to SQL '
$cubeFactTableName = @{'PnLData' = 'byDetail'; 'PnLData_ByProduct' = 'byProduct'; 'PnLData_ByProfitCenter' = 'byPC' }


# Name of the base parition - required is it provide structure and data source
$basePartitionName = 'Base'


# If set to $true, only MCB data
# $MCBOnly = $cubeName.Contains('MCB')
$MCBOnly = $true


# Query related attributes
# Key value, key is what is used for script logic, value[0] is the value passed to the SQL script, value[1] is used for partition naming
# A = actuals; B = budget; F = forecast - for forecast, value[1] is not used, instead, $fcVersion value[1] is used for partition name
$versions = @{'A' = @('Actuals', 'ACT'); 'B' = @('Budget', 'BUD'); 'F' = @('Forecast', 'FCST') }

# C = current year; P = prior year; P2 = prior year - 2
$yearVersions = @{'C' = @('CY', 'CY'); 'P' = @('PY', 'PY'); 'P2' = @('PY2', 'PY2' ) }
# $yearVersions = @{'C' = @('CY', 'CY') }

# C = current forecast; P = prior forecast
$fcVersions = @{'C' = @('Current', 'FCST'); 'P' = ('Historical', ( -Join ('FCST', $stringSeperator, 'H'))) }


# Quarters or months?
$quarters = $false
# $quarters = $true

if ($quarters) {
    $periodNums = @(1..4)
}
else {
    $periodNums = @(1..12)
}

# What character seperates strings in partition name
$stringSeperator = '_'

function Get-ASSql {
    param (
        $sqlParams
    )

    $sqlFile = ".\SQL\base_sql.sql"
    $SQLBase = (Get-Content -Path ( -Join ($PSScriptRoot, '\', $sqlFile.TrimStart(".\")))) -Join "`n"

    # Version, YearContext, MonthNum, ForecastType, FactTable, MCB Only?
    $SQLFinal = $SQLBase -f $sqlParams.SQLVariables[0], $sqlParams.SQLVariables[1], $sqlParams.SQLVariables[2], $sqlParams.SQLVariables[3], $sqlParams.SQLVariables[4], $sqlParams.SQLVariables[5]

    Write-Output $SQLFinal | Set-Content -Path ( -Join ($PSScriptRoot, '.\SQL\final_sql.sql'))
    Return $SQLFinal
}
#endregion


########################  MAKE ALL CHANGES ABOVE  ########################
Write-Output ( -Join ('Start Time: ', ((Get-Date).ToString('MM/dd/yyyy HH:mm:ss'))))

if ( $null -ne $bimFilePath -and $bimFilePath.Length -gt 0 -and -not [IO.File]::Exists($bimFilePath) ) {
    Write-Output 'BIM file not found; quitting'
    Exit
}

# If $bimFilePath is empty, then save to server, otherwise save to bimFile
$localFile = if ($null -eq $bimFilePath -or $bimFilePath.Length -eq 0) { $false } Else { $true }

$mcb = if ($MCBOnly) { 1 } else { 0 }

#region Generate partition attributes
[System.Collections.ArrayList]$partArray = @{ }

function Get-PartArray {
    param (
        $cubeFactTableAttribute
    )

    # For versions
    $versions.GetEnumerator() | ForEach-Object {
        $versionKey = $_.Key
        $versionValue_SQL = $_.Value[0]
        $versionValue_Partition = $_.Value[1]
        # For period numbers
        $loopCnt = 1
        ForEach ($periodNum in $periodNums) {
            If ($quarters) {
                $periodNum = (@((($periodNum * 3) - 2)..($periodNum * 3))) -Join ","
                $periodNumLbl = -Join ('Q', $loopCnt)
            }
            Else {
                If ($periodNum -lt 10) {
                    $periodNumLbl = -Join ('0', $periodNum)
                }
                Else {
                    $periodNumLbl = $periodNum
                }
                $periodNumLbl = -Join ('P', $periodNumLbl)
            }
            $periodNumLbl = -Join ($stringSeperator, $periodNumLbl)

            If ($versionKey -eq 'A') {
                # For year versions
                $yearVersions.GetEnumerator() | ForEach-Object {
                    # Version, Year Version, Period, ForecastVersion, Cube Fact Attribute, MCB Only?
                    $sqlArray = $versionValue_SQL, $_.Value[0], $periodNum, '', $cubeFactTableAttribute, $mcb
                    # Version, Year Version, Detailed, Period
                    $partName = ( -Join ($versionValue_Partition, $stringSeperator, $_.Value[1], $detailedValue, $periodNumLbl))
                    $partArray.Add([PSCustomObject]@{'SQLVariables' = $sqlArray; 'PartitionName' = $partName }) | Out-Null
                }
            }
            ElseIf ($versionKey -eq 'F') {
                # For FC versions
                $fcVersions.GetEnumerator() | ForEach-Object {
                    # Version, Year Version, Period, ForecastVersion, Cube Fact Attribute, MCB Only?
                    $sqlArray = $versionValue_SQL, $yearVersions['C'][0], $periodNum, $_.Value[0], $cubeFactTableAttribute, $mcb

                    # Version, Year Version, Detailed, Period
                    $partName = ( -Join ($_.Value[1], $detailedValue, $periodNumLbl))
                    $partArray.Add([PSCustomObject]@{'SQLVariables' = $sqlArray; 'PartitionName' = $partName }) | Out-Null
                }
            }
            Else {
                # Version, Year Version, Period, ForecastVersion, Cube Fact Attribute, MCB Only?
                $sqlArray = $versionValue_SQL, $yearVersions['C'][0], $periodNum, '', $cubeFactTableAttribute, $mcb
                # Version, Year Version, Detailed, Period
                $partName = ( -Join ($versionValue_Partition, $detailedValue, $periodNumLbl))
                $partArray.Add([PSCustomObject]@{'SQLVariables' = $sqlArray; 'PartitionName' = $partName }) | Out-Null
            }
            $loopCnt++
        }
    }
}
#endregion

# Load Assembly
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.AnalysisServices') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.AnalysisServices.Core') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.AnalysisServices.Tabular') | Out-Null

# Connect to Tabular SSAS
if ($localFile) {
    $modelBim = [IO.File]::ReadAllText($bimFilePath)
    $db = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::DeserializeDatabase($modelBim)
}
else {
    $srv = New-Object Microsoft.AnalysisServices.Server

    if ($serverName -like "asazure*") {
        $desktop = [Environment]::GetFolderPath("Desktop")

        Read-Host -Prompt 'Input your email address' | Set-Content -Path ( -Join ($desktop, "\email.txt") )
        $useremail = (Get-Content -Path ( -Join ($desktop, "\email.txt") ))

        # $epassword = (Get-Credential $useremail).password | ConvertFrom-SecureString | Set-Content ( -Join ($desktop, "\password.txt"))
        $spassword = (Get-Content -Path ( -Join ($desktop, "\password.txt") ))

        $password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((ConvertTo-SecureString -String $spassword)))

        $connString = "Provider=MSOLAP;Data Source=" + $serverName + ";User ID=" + $useremail + ";Password=" + $password + ";Persist Security Info=True; Impersonation Level=Impersonate;";

        $srv.connect($connString)
    }
    else {
        $srv.connect($serverName)
    }

    # Connect to database
    $db = $srv.Databases.FindByName($cubeName);
}

if ( $null -eq $db) {
    Write-Output 'Not connected to cube database; quitting'
    Exit
}

# The model for the cube
$model = $db.model
if ( $null -eq $model) {
    Write-Output 'Not connected to cube model; quitting'
    Exit
}

#region Main worker, for creating partitions
$cubeFactTableName.GetEnumerator() | ForEach-Object {
    $cubeFactTableAttribute = $_.Value
    $cubeFactTable = $_.Key

    $partArray.Clear()
    Get-PartArray -cubeFactTableAttribute $cubeFactTableAttribute

    # The table for the model
    $tbl = $model.Tables.Find($cubeFactTable)
    if ( $null -eq $tbl) {
        Write-Output ( -Join ($cubeFactTable, ' - Not connected to cube table; quitting'))
        Exit
    }

    # Get base data source and remove all other partitions
    $staticPartitions = @()
    $staticPartitions += $tbl.Partitions

    $baseExist = $false
    $staticPartitions.GetEnumerator() | ForEach-Object {
        If ($_.Name -eq $basePartitionName) {
            if (1 -eq 1) {
                $baseExist = $true
            }
        }
    }

    if ( -not ($baseExist) ) {
        Write-Output ( -Join ($cubeFactTable, ' - Base partition not found; quitting'))
        Exit
    }

    $staticPartitions.GetEnumerator() | ForEach-Object {
        If ($_.Name -eq $basePartitionName) {
            $ds = New-Object -TypeName Microsoft.AnalysisServices.Tabular.QueryPartitionSource
            $ds.DataSource = $_.Source.DataSource
        }
        Else {
            $tbl.Partitions.Remove($_.Name)
        }
    }
    if (-not $localFile) {
        $db.Update([Microsoft.AnalysisServices.UpdateOptions]::ExpandFull)
    }


    #region Create new partitions
    $partArray.GetEnumerator() | ForEach-Object {
        $indPart = $_
        $partition = New-Object -TypeName Microsoft.AnalysisServices.Tabular.Partition
        $partition.Source = New-Object -TypeName Microsoft.AnalysisServices.Tabular.QueryPartitionSource
        $partition.Source.DataSource = $ds.DataSource

        $partition.Source.Query = Get-ASSql -sqlParams $indPart
        $partition.Name = $_.PartitionName

        $partition.Mode = [Microsoft.AnalysisServices.Tabular.ModeType]::Import;
        $partition.DataView = [Microsoft.AnalysisServices.Tabular.DataViewType]::Default;

        $tbl.Partitions.Add($partition)
    }
    #endregion
}
#endregion

# Save changes to server
if (-not $localFile) {
    try {
        $db.Update([Microsoft.AnalysisServices.UpdateOptions]::ExpandFull)
        $srv.Disconnect()
        Write-Output('** Server Updated ** ')
    }
    catch {
        Write-Output( -Join ('Error refreshing the server: ', $serverName))
    }
}
else {
    $serializeOptions = New-Object Microsoft.AnalysisServices.Tabular.SerializeOptions
    $serializeOptions.IgnoreTimestamps = $true
    $serializeOptions.IgnoreInferredProperties = $true
    $serializeOptions.IgnoreInferredObjects = $true

    $cubeJson = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::SerializeDatabase($db, $serializeOptions)
    $cubeJson | Out-File $bimFilePath
    Write-Output('** BIM File Updated ** ')
}

Write-Output ( -Join ('End Time: ', ((Get-Date).ToString('MM/dd/yyyy HH:mm:ss'))))