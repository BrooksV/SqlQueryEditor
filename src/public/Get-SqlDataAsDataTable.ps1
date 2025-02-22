# Function to load data into a DataSet
Function Get-SqlDataAsDataTable {
    [CmdletBinding()]
    Param (
        [string]$ConnectionString = $syncHash.Params.ConnectionString,
        [string]$query = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME;", 
        [System.Data.SqlClient.SqlConnection]$SQLConnection = $syncHash.SqlConnection,
        [Boolean]$DisplayResults = $true
    )
    Write-Host "Get-SqlDataAsDataTable()"

    $localdb = ($SQLConnection.DataSource -split '\\')[-1]
    If (-not $localdb) {
        $localdb = ($ConnectionString -split ';').Where({$_.StartsWith('Data Source')}).ForEach({$_ -split '\\'})[-1]
    }
    $status = ((SqlLocalDB.exe info "$localdb") -split [environment]::NewLine).Where({$_})
    Switch ($status.Where({$_.StartsWith('State:')}).ForEach({($_ -split ' ')[-1]})) {
        'Stopped' {SqlLocalDB.exe Start "$localdb" ; break}
    }
    $Result = $null
    try {
        If ($SQLConnection.State -ne 'Open') {
            $SQLConnection = [System.Data.SqlClient.SqlConnection]::new()
            $SQLConnection.ConnectionString = $ConnectionString
            Try {
                $SQLConnection.Open()
            } Catch {
                If ($SQLConnection) {
                    $SQLConnection.Close()
                    $SQLConnection.Dispose()
                    $SQLConnection = $null
                }
                Write-host ($_ | Out-String) -ForegroundColor Red
            }
        }
        $SQLCommand = $SQLConnection.CreateCommand()
        $SQLCommand.CommandText = $query
        $SQLCommand.CommandTimeout = 600
        $SQLCommand.Connection = $SQLConnection
        # $SqlDataAdapter = [System.Data.SqlClient.SqlDataAdapter]::new($SQLCommand)
        # $dataset = [System.Data.DataSet]::new()
        # [void]$SqlDataAdapter.Fill($dataset)
        $SQLReader = $SQLCommand.ExecuteReader()
        If ($SQLReader) {
            $Result = [System.Data.DataTable]::new()
            $Result.Load($SQLReader)
            Write-Verbose ("`$Result.GetType().ToString() = ($($Result.GetType().ToString()))")
            Write-Verbose ("`$Result[0].GetType().ToString() = ($($Result[0].GetType().ToString()))")
            Write-Verbose ("`$Result -is [System.Array] = ($($Result -is [System.Array]))")
            Write-Verbose ("`$Result | GM = ($($Result | Get-Member | Out-String))")
            Write-Verbose ("`$Result[0] | GM = ($($Result[0] | Get-Member | Out-String))")
            If ($DisplayResults) {
                Write-Host ($Result[0] | Out-String)
            }
        } Else {
            Write-Warning "Get-SqlDataAsDataTable() Unable to load data from SQLReader"
            Return $null
        }
        Return ($Result -as [System.Data.DataTable])
        # Return ([System.Data.DataTable]$Result)
    } catch {
        Write-Error "Failed to load data: $_"
        return $null
    }
}
$RetResult = Get-SqlDataAsDataTable -ConnectionString 'Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=F:\DATA\BILLS\PSSCRIPTS\SCANMYBILLS\DATABASE1.MDF;Integrated Security=True' -Verbose

Write-Host ("`$RetResult.GetType().ToString() = ($($RetResult.GetType().ToString()))")
Write-Host ("`$RetResult[0].GetType().ToString() = ($($RetResult[0].GetType().ToString()))")

Write-Host ("`n`$RetResult")
Write-Host ($RetResult | Out-String)

Write-Host ("`$RetResult[0]")
Write-Host ($RetResult[3] | Out-String)

Write-Host ("`$RetResult -is [System.Array] = ($($RetResult -is [System.Array]))")
Write-Host ("`$RetResult -is [System.Collections.ArrayList] = ($($RetResult -is [System.Collections.ArrayList]))")
Write-Host ("`$RetResult -is [System.Data.DataTable] = ($($RetResult -is [System.Data.DataTable]))")
Write-Host ("`$RetResult -is [System.Data.DataRow] = ($($RetResult -is [System.Data.DataRow]))")
Write-Host ("`$RetResult -is [System.Data.DataSet] = ($($RetResult -is [System.Data.DataSet]))")
Write-Host ("`$RetResult -is [System.Data.DataTableReader] = ($($RetResult -is [System.Data.DataTableReader]))")
Write-Host ("`$RetResult -is [System.Data.DataRowCollection] = ($($RetResult -is [System.Data.DataRowCollection]))")
Write-Host ("`$RetResult -is [System.Data.SqlClient.SqlDataAdapter] = ($($RetResult -is [System.Data.SqlClient.SqlDataAdapter]))")
Write-Host ("`$RetResult.DefaultView -is [System.Data.DataView] = ($($RetResult.DefaultView -is [System.Data.DataView]))")

Write-Host ("`$RetResult[0] -is [System.Array] = ($($RetResult[0] -is [System.Array]))")
Write-Host ("`$RetResult[0] -is [System.Collections.ArrayList] = ($($RetResult[0] -is [System.Collections.ArrayList]))")
Write-Host ("`$RetResult[0] -is [System.Data.DataTable] = ($($RetResult[0] -is [System.Data.DataTable]))")
Write-Host ("`$RetResult[0] -is [System.Data.DataRow] = ($($RetResult[0] -is [System.Data.DataRow]))")
Write-Host ("`$RetResult[0] -is [System.Data.DataSet] = ($($RetResult[0] -is [System.Data.DataSet]))")
Write-Host ("`$RetResult[0] -is [System.Data.DataTableReader] = ($($RetResult[0] -is [System.Data.DataTableReader]))")
Write-Host ("`$RetResult[0] -is [System.Data.DataRowCollection] = ($($RetResult[0] -is [System.Data.DataRowCollection]))")
Write-Host ("`$RetResult[0] -is [System.Data.SqlClient.SqlDataAdapter] = ($($RetResult[0] -is [System.Data.SqlClient.SqlDataAdapter]))")
Write-Host ("`$RetResult[0].DefaultView -is [System.Data.DataView] = ($($RetResult[0].DefaultView -is [System.Data.DataView]))")
