# .SYNOPSIS
# This function establishes a new database connection.

# .DESCRIPTION
# The New-DatabaseConnection function checks if there's an existing SQL connection. If not, it attempts to create a new connection 
# using the connection string provided in the syncHash.Params.ConnectionString. If the connection is successful, 
# it is stored in the syncHash.SqlConnection variable.

# .EXAMPLE
# $connection = New-DatabaseConnection

# .NOTES
# Author: Brooks Vaughn
# Date: 02/21/2025

Function New-DatabaseConnection {
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param()

    If ($PSCmdlet.ShouldProcess("Establishing new database connection")) {
        If ($syncHash.SqlConnection -and $syncHash.SqlConnection.State -eq 'Open') {
            Return $syncHash.SqlConnection
        }
    }
    Try {
        Write-Host "New-DatabaseConnection()"
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $syncHash.Params.ConnectionString
        $connection.Open()
        $syncHash.SqlConnection = $connection
    } Catch {
        Write-Error "Failed to establish database connection: $_"
        Return $null
    }
}
