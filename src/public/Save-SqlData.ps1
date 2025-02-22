# Function to save data from DataGrid to the database
Function Save-SqlData {
    Param (
        [string]$TableName,
        [Object]$DataGrid
    )

    If ([String]::IsNullOrEmpty($TableName)) {
        Write-Warning ('Save-SqlData(): Parameter -TableName is Missing or blank')
        Return
    }
    If ([String]::IsNullOrEmpty($DataGrid)) {
        Write-Warning ('Save-SqlData(): Parameter -DataGrid is Missing or null')
        Return
    }
    Write-Host ("Save-SqlData(): Saving Changes made to `$TableName=({0}) in `$DataGrid=({1}) of type ({2})" -f $TableName, $DataGrid.Name, $DataGrid.GetType().ToString())
    try {
        $SQLConnection = New-DatabaseConnection
        if ($null -eq $SQLConnection) {
             Write-Error "Database connection is not initialized."
             return
        }

        # $SQLCommand = $SQLConnection.CreateCommand()
        # $SQLCommand.CommandText = "SELECT * FROM [$TableName]"
        # $SQLCommand.CommandTimeout = 600
        # $SQLCommand.Connection = $SQLConnection

        # $SqlDataAdapter = [System.Data.SqlClient.SqlDataAdapter]::new($SQLCommand)
        # $SqlCommandBuilder = [System.Data.SqlClient.SqlCommandBuilder]::new($SqlDataAdapter)

        # # Update the database with the changes made to the DataTable
        # $SqlDataAdapter.Update($DataGrid)
        # Write-Host "Data saved successfully to the $TableName table."
    } catch {
        Write-Error "Failed to save data: $_"
    }

    #--------------------------------------
    # 
    #--------------------------------------
    Try {
        # $syncHash.UI.rowBeingEdited.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)

        <# Examine Changes #>
        # Create a DataView with the table.
        # $dataView = [System.Data.DataView]::new($syncHash.UI.currentDataGrid.ItemsSource.Table)
        # $dataView = [System.Data.DataView]::new($DataGrid.ItemsSource.Table)
        # $dataView = [System.Data.DataView]::new($DataGrid.ItemsSource)
        $dataView = [System.Data.DataView]::new($DataGrid.ItemsSource)

        # Set the RowStateFilter to display only added and modified rows.
        $dataView.RowStateFilter = ([System.Data.DataViewRowState]::Deleted -bor [System.Data.DataViewRowState]::Added -bor [System.Data.DataViewRowState]::ModifiedCurrent)

        # ForEach ($row in $dataView) {Write-Host ($row | FT -AutoSize | Out-String).Trim()}
        Write-Host ($dataView | Format-Table -AutoSize | Out-String).Trim()
        #>
        
        #----------------------------------------------------------------------------------------
        # Need to replace the current SqlDataAdapter Results with the modified DataGrid Results
        #   Do we need to replace everything or can we just updated the changed rows?
        #----------------------------------------------------------------------------------------
        # $syncHash.UI.SqlResults.Result = $DataGrid.ItemsSource

        # $syncHash.UI.DataSet.TableIndex = [int][TableIndexes]$TableName

        # $syncHash.UI.DataSet.TableIndex = $syncHash.UI.currentDataTable.TableIndex
        # $syncHash.UI.currentDataTable.Parent.SaveChanges()
        # $syncHash.UI.DataSet.SaveChanges()

        # $WPF_dgTableEdit.DataContext.Tables[0].AcceptChanges()
        # $syncHash.UI.IsCurrentDataGridDirty = $syncHash.UI.currentDataTable.isDirty = $false

        If ($This.ConnectionString) {
            $SaveChangesConnectionString = $This.ConnectionString
        } Else {
            $SaveChangesConnectionString = $This.BuildConnectionString()
        }

        $SaveChangesConnection = [System.Data.SqlClient.SqlConnection]::new()
        $SaveChangesConnection.ConnectionString = $SaveChangesConnectionString

        Try {
            $This.KeepAlive = $false
            $This.OpenConnection()

            #--------------------------------------------------
            # Create a DataView to examine the changes
            #--------------------------------------------------
            # $dataView = [System.Data.DataView]::new($This.Result.Tables[0])
            # # Set the RowStateFilter to display only added and modified rows.
            # $dataView.RowStateFilter = ([System.Data.DataViewRowState]::Deleted -bor [System.Data.DataViewRowState]::Added -bor [System.Data.DataViewRowState]::ModifiedCurrent)
            # ForEach ($row in $dataView) {Write-Host ($row | FT -AutoSize | Out-String).Trim()}
            # Write-Host ($dataView | FT -AutoSize | Out-String).Trim()

            If (-not [String]::IsNullOrEmpty($table.SqlDataAdapter.DeleteCommand)) {
                $table.SqlDataAdapter.DeleteCommand.Connection = $This.SQLConnection
            }
            If (-not [String]::IsNullOrEmpty($table.SqlDataAdapter.UpdateCommand)) {
                $table.SqlDataAdapter.UpdateCommand.Connection = $This.SQLConnection
            }
            If (-not [String]::IsNullOrEmpty($table.SqlDataAdapter.InsertCommand)) {
                $table.SqlDataAdapter.InsertCommand.Connection = $This.SQLConnection
            }
            Try { # First process deletes.
                $table.SqlDataAdapter.Update($table.Result.Tables[0].Select($null, $null, [System.Data.DataViewRowState]::Deleted))
            } Catch {}
            Try { # Next process updates.
                $table.SqlDataAdapter.Update($table.Result.Tables[0].Select($null, $null, [System.Data.DataViewRowState]::ModifiedCurrent))
            } Catch {}
            Try { # Finally, process inserts.
                $table.SqlDataAdapter.Update($table.Result.Tables[0].Select($null, $null, [System.Data.DataViewRowState]::Added))
                $table.Result.Tables[0].AcceptChanges()
            } Catch {}
        } Catch {
            return $(Write-host ($_ | Out-String) -ForegroundColor Red) 
        } Finally {
            $table.Result.AcceptChanges()
            $This.CloseConnection()
        }

        If ($This.DisplayResults) {
            Return $table.Result.Tables[0]
        } Else {
            Return $null
        }

        #?# Reload-TableData

    } Catch {
        Write-Log -LogMethod $LogMethod -Message ("Save-SqlData() Error: ({0})" -f $error[0]) # .Exception.Message)
    }

    # Apply-ScriptFilters -QueryType $SyncHash.Params.WorkbookType
    # $syncHash.UI.currentDataGrid = $DataGrid
    # $syncHash.UI.currentDataGridView.Source = $syncHash.UI.currentDataGrid.ItemsSource

    $WPF_mnuAct.Items.Where({$_.Name -in @('mnuDelete','mnuEdit','mnuCancel','mnuSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})    
}
