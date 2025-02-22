# Function to load data into a DataSet
Function Get-SqlData {
    Param (
        [string]$query, 
        [System.Data.SqlClient.SqlConnection]$SQLConnection = $syncHash.SqlConnection,
        [string]$ConnectionString = $syncHash.Params.ConnectionString
    )
    try {
        Write-Host "Get-SqlData()"
        If ($SQLConnection.State -ne 'Open') {
            $SQLConnection = [System.Data.SqlClient.SqlConnection]::new()
            $SQLConnection.ConnectionString = $ConnectionString
            Try {
                $SQLConnection.Open()
                $syncHash.SqlConnection = $SQLConnection
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
        $SqlDataAdapter = [System.Data.SqlClient.SqlDataAdapter]::new($SQLCommand)
        $dataset = [System.Data.DataSet]::new()
        [void]$SqlDataAdapter.Fill($dataset)
        return $dataSet
    } catch {
        Write-Error "Failed to load data: $_"
        return $null
    }
}
