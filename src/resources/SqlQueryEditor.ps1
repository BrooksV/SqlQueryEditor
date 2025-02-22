<#

. "C:\Git\SqlQueryEditor\dist\SqlQueryEditor\SqlQueryEditor.ps1"

#>

[System.Diagnostics.DebuggerHidden()]
param()

Write-Warning ($error | Out-String)

#region Initialize Variables and Check STA Mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $CommandLine = $MyInvocation.Line.Replace($MyInvocation.InvocationName, $MyInvocation.MyCommand.Definition)
    Write-Warning 'Script is not running in STA Apartment State.'
    Write-Warning 'Attempting to restart this script with the -Sta flag...'
    Write-Verbose "Script: $CommandLine"
    Start-Process -FilePath PowerShell.exe -ArgumentList "$CommandLine -Sta"
    exit
}

Function Initialize-Module {
    Param (
        [string]$moduleName
    )
    # Check if the module is already installed
    $module = Get-Module -ListAvailable -Name $moduleName
    if ($null -eq $module) {
        # Module is not installed, install it
        Write-Output "Module '$moduleName' is not installed. Installing..."
        Install-Module -Name $moduleName -Repository PSGallery -Scope CurrentUser
    
        # Import the newly installed module
        Write-Output "Importing module '$moduleName'..."
        Import-Module -Name $moduleName
    } else {
        # Module is already installed, import it
        Write-Output "Module '$moduleName' is already installed. Importing..."
        Import-Module -Name $moduleName
    }
    # Verify the module is imported
    if (Get-Module -Name $moduleName) {
        Write-Output "Module '$moduleName' has been successfully imported."
    } else {
        Write-Output "Failed to import module '$moduleName'."
    }
}    
# Verify required modules are installed and imported
Initialize-Module -moduleName "SqlQueryClass"
Initialize-Module -moduleName "GuiMyPS"
# Import-Module C:\Git\GuiMyPS\dist\GuiMyPS\GuiMyPS.psd1 -Verbose -Force

# Initialize-Module -moduleName "ModuleTools"
Import-Module C:\Git\ModuleTools\dist\ModuleTools\ModuleTools.psd1 # -Verbose -Force

<# How to use local version of SqlQueryClass module # >
# Import the SqlQueryClass module from the PowerShell Gallery with the required version
$moduleVersion = (Find-Module -Name SqlQueryClass).Version
Import-Module -Name SqlQueryClass -RequiredVersion $moduleVersion -Scope Local -Force -Verbose

# Remove any existing instances of the SqlQueryClass module
if (Get-Module -Name SqlQueryClass) {
    Remove-Module -Name SqlQueryClass -Force -Verbose
}

# Import the SqlQueryEditor module 
Import-Module ".\dist\SqlQueryEditor\SqlQueryEditor.psd1" -Force -Verbose
#>

<# Not ready to test exports to excel # >
# Install and Import ImportExcel module
If (-Not (Get-Module -Name ImportExcel -ListAvailable)) {
    Install-Module -Name ImportExcel -Force
}
Import-Module ImportExcel
#>

# Import the SqlQueryEditor module as local functions were moved to a module
# Import-Module ".\dist\SqlQueryEditor\SqlQueryEditor.psd1" -Force -Verbose

$projectData = Get-MTProjectInfo

# Load dependencies and WPF assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Data

# Initialize Constants
# $eol = [Environment]::NewLine
# $StartTime = Get-Date
# $TimeStamp = $StartTime.ToString('yyyy-MMdd-HHmm')     # Time Stamp appended to Saved output files
$Error.Clear()

# Initialize Script Variables
# $syncHash = [System.Collections.Hashtable]::Synchronized((New-Object System.Collections.Hashtable))
$hashTable = @{}
$syncHash = [System.Collections.Hashtable]::Synchronized($hashTable)
$syncHash.Add('Form', [object])
$syncHash.Add('Errors', [System.Collections.ArrayList]@())
$syncHash.Add('Params', @{
    Title = 'SqlQueryEditor'
    DocumentFolder = 'F:\Data\Bills'
    LogFolder = 'C:\Temp\SqlQuery'
    LogFileNamePattern = '{0}\Transcript_Log_{1}.txt'
    # ScriptsFolder = 'C:\Git\SqlQueryEditor\scripts'
    ScriptsFolder = 'C:\Git\SqlQueryEditor\dist\SqlQueryEditor'
    Cred = [System.Management.Automation.PSCredential]::Empty
    SqlServer = '(localdb)\MSSQLLocalDB'
    DatabaseName = 'F:\DATA\BILLS\PSSCRIPTS\SCANMYBILLS\DATABASE1.MDF'
    # ConnectionString = "Server=(localdb)\MSSQLLocalDB;AttachDbFilename=F:\DATA\BILLS\PSSCRIPTS\SCANMYBILLS\DATABASE1.MDF;Integrated Security=True"
    ConnectionString = "Data Source=(localdb)\MSSQLLocalDB;AttachDbFilename=F:\DATA\BILLS\PSSCRIPTS\SCANMYBILLS\DATABASE1.MDF;Integrated Security=True"
    IsRunning = $false
})
$syncHash.Add('UI', [PSCustomObject]@{
    # Set the initial content of the toggle button to reflect Record Edit Mode
    EditMode = [EditMode]::Record # [EditMode]::Table # 
    SqlResults = $null
    LeftColumnPreviousWidth = $null
    timer = $null # [System.Windows.Threading.DispatcherTimer]
    IsCurrentDataGridDirty = $false
})

# Define EditMode enum
enum EditMode { Record; Table }
$syncHash.UI.EditMode = [EditMode]::Record
enum EditButtonMode { Edit; Cancel; }
enum MenuItemMode { Disabled; Enabled; }

# Set up Transcript Logging
Start-Transcript -Path "$($syncHash.Params.LogFolder)\SqlQuery_Transcript.log" -Append

#endregion

#region Load and Process XAML
Write-Host ("Current Location: ($(Get-Location))") -ForegroundColor Magenta
Write-Host ("PSScriptRoot: ($($PSScriptRoot))") -ForegroundColor Magenta
$syncHash.Form = New-XamlWindow -xaml "$PSScriptRoot\SqlQueryEditor.xaml"

#region Event handler for Click Events
$handler_MenuItem_Click = {
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host ("`$handler_menu_Click() Menu Item clicked: {0}" -f $theSender.Name)

    Switch -Regex ($theSender.Name) {
        # Main Menu
        'mnuExit|mnuExit2' { $syncHash.Form.Close()
            Break }
        'mnuQuit' {
            $syncHash.Form.Close()
            Break }
        'mnuMExportCopyToClipboard' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                Save-DataGrid -InputObject $WPF_dataGridSqlQuery -SaveAs Clipboard -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsClipboard Aborted')
            }
            Break }
        'mnuDExportCopyToClipboard' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                Save-DataGrid -InputObject $WPF_dataGridSqlQueryParms -SaveAs Clipboard -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsClipboard Aborted')
            }
            Break }
        'mnuMSaveAsCSV' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                Save-DataGrid -InputObject $WPF_dataGridSqlQuery -SaveAs CVS -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsCSV Aborted')
            }
            Break }
        'mnuDSaveAsCSV' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                Save-DataGrid -InputObject $WPF_dataGridSqlQueryParms -SaveAs CVS -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsCSV Aborted')
            }
            Break }
        'mnuMSaveAsExcel' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                $WPF_dataGridSqlQuery | Save-DataGrid -SaveAs Excel -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsExcel Aborted')
            }
            Break }
        'mnuDSaveAsExcel' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                $WPF_dataGridSqlQueryParms | Save-DataGrid -SaveAs Excel -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsExcel Aborted')
            }
            Break }
        'mnuMExportAsExcel' { 
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                $WPF_dataGridSqlQuery.ItemsSource | Save-DatasetToExcel -Path ("{0}\{1}" -f $syncHash.Params.LogFolder, 'SQLQuery_Export.xlsx')
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAllToExcel Aborted')
            }
            Break }
        'mnuDExportAsExcel' { 
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                $WPF_dataGridSqlQueryParms.ItemsSource | Save-DatasetToExcel -Path ("{0}\{1}" -f $syncHash.Params.LogFolder, 'SQLQueryParms_Export.xlsx')
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAllToExcel Aborted')
            }
            Break }
   
        'mnuMRefresh' {
            $SelectedItem = $WPF_dataGridSqlQuery.SelectedItem

            If (-not [String]::IsNullOrEmpty($SelectedItem)) {
                $WPF_dataGridSqlQuery.SelectedItem = $SelectedItem
                $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            }
            $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMEdit','mnuMCancel','mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_dataGridSqlQuery.ItemsSource = $null
            $WPF_dataGridSqlQuery.Items.Clear()
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQuery']]
            $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
            #?# Load-TableData -TableName 'SqlQuery' -DataGrid $WPF_dataGridSqlQuery
            $WPF_dataGridSqlQuery.ScrollIntoView(-1)
            Break }
        'mnuDRefresh' {
            $SelectedItem = $WPF_dataGridSqlQueryParms.SelectedItem
            If (-not [String]::IsNullOrEmpty($SelectedItem)) {
                $WPF_dataGridSqlQueryParms.SelectedItem = $SelectedItem
                $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            }
            $WPF_menuDetailDataGrid.Items.Where({$_.Name -in @('mnuMEdit','mnuMCancel','mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_dataGridSqlQueryParms.ItemsSource = $null
            $WPF_dataGridSqlQueryParms.Items.Clear()
            $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
            $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView
            # Load-TableData -TableName 'SqlQueryParms' -DataGrid $WPF_dataGridSqlQueryParms
            $WPF_dataGridSqlQueryParms.ScrollIntoView(-1)
            Break }
        'mnuMFirst' {
            $WPF_dataGridSqlQuery.SelectedIndex = 0
            $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            Break }
        'mnuDFirst' {
            $WPF_dataGridSqlQueryParms.SelectedIndex = 0
            $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            Break }
        'mnuMLast' {
            If ($WPF_dataGridSqlQuery.Items.Count -ge 2) {
                $WPF_dataGridSqlQuery.SelectedIndex = $WPF_dataGridSqlQuery.Items.Count - 2
            } Else {
                $WPF_dataGridSqlQuery.SelectedIndex = 0
            }
            #?# $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            Break }
        'mnuDLast' {
            If ($WPF_dataGridSqlQueryParms.Items.Count -ge 2) {
                $WPF_dataGridSqlQueryParms.SelectedIndex = $WPF_dataGridSqlQueryParms.Items.Count - 2
            } Else {
                $WPF_dataGridSqlQueryParms.SelectedIndex = 0
            }
            $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            Break }
        'mnuMPrevious' {
            If ($WPF_dataGridSqlQuery.SelectedIndex -gt 0) {
                $WPF_dataGridSqlQuery.SelectedIndex--
            } Else {
                # Wrap to Last Record
                $WPF_dataGridSqlQuery.SelectedIndex = $WPF_dataGridSqlQuery.Items.Count - 2
            }
            $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            Break }
        'mnuDPrevious' {
            If ($WPF_dataGridSqlQueryParms.SelectedIndex -gt 0) {
                $WPF_dataGridSqlQueryParms.SelectedIndex--
            } Else {
                # Wrap to Last Record
                $WPF_dataGridSqlQueryParms.SelectedIndex = $WPF_dataGridSqlQueryParms.Items.Count - 2
            }
            $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            Break }
        'mnuMNext' {
            If ($WPF_dataGridSqlQuery.SelectedIndex + 2 -lt $WPF_dataGridSqlQuery.Items.Count) {
                $WPF_dataGridSqlQuery.SelectedIndex++
            } Else {
                # Wrap to First Record (Skip over New Record)
                $WPF_dataGridSqlQuery.SelectedIndex = 0
            }
            $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            Break }
        'mnuDNext' {
            If ($WPF_dataGridSqlQueryParms.SelectedIndex + 2 -lt $WPF_dataGridSqlQueryParms.Items.Count) {
                $WPF_dataGridSqlQueryParms.SelectedIndex++
            } Else {
                # Wrap to First Record (Skip over New Record)
                $WPF_dataGridSqlQueryParms.SelectedIndex = 0
            }
            $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            Break }
        'mnuMAdd' {
            $WPF_dataGridSqlQuery.SelectedItem = $WPF_dataGridSqlQuery.ItemsSource.AddNew()
            $WPF_dataGridSqlQuery.ScrollIntoView($WPF_dataGridSqlQuery.SelectedItem)
            Break }
        'mnuDAdd' {
            $WPF_dataGridSqlQueryParms.SelectedItem = $WPF_dataGridSqlQueryParms.ItemsSource.AddNew()
            $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            Break }
        'mnuMEdit' {
            $WPF_menuDetailDataGrid.Items.Where({$_.Name -in @('mnuMCancel')}).ForEach({$_.Tag = ([MenuItemMode]::Enabled -as [Int]).ToString()})
            Write-Host ('`$WPF_dataGridSqlQuery.SelectedIndex: {0}' -f $WPF_dataGridSqlQuery.SelectedIndex)
            # Write-Host ('`$syncHash.UI.rowBeingEdited: {0}' -f $WPF_dataGridSqlQuery.rowBeingEdited)

            # $selectedRow = Get-Row -index $WPF_dataGridSqlQuery.SelectedIndex
            # Write-Host ('`$selectedRow: {0}' -f $selectedRow.Item)

            # $Cell = Get-Cell -row $WPF_dataGridSqlQuery.SelectedIndex -column 1
            # Write-Host ('`$Cell: Column: {0}, Content: ({1})' -f $Cell.Column.Header, $Cell.Content.Text)

            # If ($cell) {
            #     # $cell.BringIntoView()
	        #     $cell.Focus()
            #     $cell.IsEditing = $true
            #     $WPF_dataGridSqlQuery.SelectedIndex = -1
            # }
            Break }
        'mnuMCancel' {
            # Reset MenuItems when Data is Refreshed
            # 'mnuFirst','mnuLast','mnuPrevious','mnuNext','mnuEdit','mnuCancel.Tag','mnuSave','mnuDelete'
            $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMEdit','mnuMCancel','mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_dataGridSqlQuery.ItemsSource = $null
            $WPF_dataGridSqlQuery.Items.Clear()
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQuery']]
            $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView            
            #?# Load-TableData -TableName 'SqlQuery' -DataGrid $WPF_dataGridSqlQuery
            $WPF_dataGridSqlQuery.ScrollIntoView(-1)
            Break }
        'mnuDCancel' {
            # Reset MenuItems when Data is Refreshed
            # 'mnuFirst','mnuLast','mnuPrevious','mnuNext','mnuEdit','mnuCancel.Tag','mnuSave','mnuDelete'
            $WPF_menuDetailDataGrid.Items.Where({$_.Name -in @('mnuDEdit','mnuDCancel','mnuDSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_dataGridSqlQueryParms.ItemsSource = $null
            $WPF_dataGridSqlQueryParms.Items.Clear()
            $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
            $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView
            #?# Load-TableData -TableName 'SqlQueryParms' -DataGrid $WPF_dataGridSqlQueryParms
            $WPF_dataGridSqlQueryParms.ScrollIntoView(-1)
            Break }
        'mnuMSave' {
            # Write-Host ('`$syncHash.UI.currentDataGrid.SelectedIndex: {0}' -f $syncHash.UI.currentDataGrid.SelectedIndex)
            # Write-Host ('`$syncHash.UI.rowBeingEdited: {0}' -f ($syncHash.UI.rowBeingEdited | Select-Object -Property * | Out-String))
            # $selectedRow = Get-Row -index $syncHash.UI.currentDataGrid.SelectedIndex
            # Write-Host ('`$selectedRow: {0}' -f $selectedRow.Item)
            # Set-Message -Message ('DataSet Has Changes: {0}' -f ($syncHash.UI.currentDataTable.Result.HasChanges()))
            # Set-Message -Message ('Committing Changes: {0}' -f ($syncHash.UI.currentDataGrid.CommitEdit()))
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQuery']]
            $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
            #?# Save-SqlData -TableName 'SqlQuery' -DataGrid $syncHash.UI.SqlResults
            Break }
        'mnuDSave' {
            # Write-Host ('`$syncHash.UI.currentDataGrid.SelectedIndex: {0}' -f $syncHash.UI.currentDataGrid.SelectedIndex)
            # Write-Host ('`$syncHash.UI.rowBeingEdited: {0}' -f ($syncHash.UI.rowBeingEdited | Select-Object -Property * | Out-String))
            # $selectedRow = Get-Row -index $syncHash.UI.currentDataGrid.SelectedIndex
            # Write-Host ('`$selectedRow: {0}' -f $selectedRow.Item)
            # Set-Message -Message ('DataSet Has Changes: {0}' -f ($syncHash.UI.currentDataTable.Result.HasChanges()))
            # Set-Message -Message ('Committing Changes: {0}' -f ($syncHash.UI.currentDataGrid.CommitEdit()))

            $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
            $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView

            #?# Save-SqlData -TableName 'SqlQueryParms' -DataGrid $syncHash.DetailData
            
            Break }
        'mnuMDelete' {
            #?# $WPF_menuMasterDataGrid.SelectedItem.Delete()
            $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMDelete','mnuMEdit')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Enabled -as [Int]).ToString()})
            $syncHash.UI.IsCurrentDataGridDirty = $true
            Break }
        'mnuDDelete' {
            #?# $WPF_menuMasterDataGridParms.SelectedItem.Delete()
            $WPF_menuDetailDataGridParms.Items.Where({$_.Name -in @('mnuDDelete','mnuDEdit')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            $WPF_menuDetailDataGridParms.Items.Where({$_.Name -in @('mnuDSave')}).ForEach({$_.Tag = ([MenuItemMode]::Enabled -as [Int]).ToString()})
            $syncHash.UI.IsCurrentDataGridDirty = $true
            Break }
        default {
            Write-Host ("{0}: {1}({2})" -f $theSender.Name, $e.OriginalSource.Name, $e.OriginalSource.ToString())
        }
    }
}
#---------------------------------------------------------------------------
# Add MenuItem Click Events to all MenuItems by locating Main Menu Object 
#    and adding a Generic Menu Click Event Handler
#---------------------------------------------------------------------------
# Add-ClickToMenuItems -MenuObj $WPF_menuMasterDataGrid -Handler $handler_menu_Click -Verbose:$false
# Add-ClickToMenuItems -MenuObj $WPF_menuDetailDataGrid -Handler $handler_menu_Click -Verbose:$false
$elements = @()
# $elements += Find-EveryControl -Element $syncHash.Form -ControlType 'System.Windows.Controls.Button'
# $elements.ForEach({$_.Element.Add_Click($handler_Button_Click)})
# $elements += Find-EveryControl -Element $form -ControlType 'System.Windows.Controls.Primitives.ToggleButton'
$elements += Find-EveryControl -Element $syncHash.Form -ControlType 'System.Windows.Controls.MenuItem'
$elements.ForEach({$_.Element.Add_Click($handler_MenuItem_Click)})

$handler_toggleEditMode = {
    param([object]$theSender, [System.Windows.RoutedEventArgs]$e)
    Write-Host ("`$handler_toggleEditMode() {0}: ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim() -ForegroundColor Yellow
    If ($syncHash.UI.EditMode -eq [EditMode]::Table) {
        $syncHash.UI.EditMode = [EditMode]::Record
        $WPF_toggleEditMode.Content = "Record Edit Mode - Click for Table Mode"
    } Else {
        $syncHash.UI.EditMode = [EditMode]::Table
        $WPF_toggleEditMode.Content = "Table Edit Mode - Click for Record Mode"
    }
}
$WPF_toggleEditMode.add_Checked($handler_toggleEditMode)
$WPF_toggleEditMode.add_Unchecked($handler_toggleEditMode)

# $WPF_dataGrid.ItemsSource = $syncHash.UI.SqlResults.Tables[0].Result.DefaultView

    # Event handler for LeftSide Expander Collapsed
    $WPF_expLeftSide.add_Collapsed({
        # Save the current width
        $syncHash.UI.LeftColumnPreviousWidth = $WPF_LeftColumn.Width
        $WPF_LeftColumn.Width = [System.Windows.GridLength]::Auto
    })

    # Event handler for LeftSide Expander Expanded
    $WPF_expLeftSide.add_Expanded({
        $WPF_LeftColumn.Width = $syncHash.UI.LeftColumnPreviousWidth
    })

    # Event handler for LeftSide Expander Collapsed
    $WPF_expSelectQueryGrid.add_Collapsed({
        # Save the current width
        $syncHash.UI.LeftColumnPreviousWidth = $WPF_LeftColumnM.Width
        $WPF_LeftColumnM.Width = [System.Windows.GridLength]::Auto
    })

    # Event handler for LeftSide Expander Expanded
    $WPF_expSelectQueryGrid.add_Expanded({
        $WPF_LeftColumnM.Width = $syncHash.UI.LeftColumnPreviousWidth
    })
    
#endregion

#region Helper Functions
Write-Host "Helper Functions were Moved to SqlQueryEditor Module"
Write-Host "Data Management Functions were Moved to SqlQueryEditor Module"
Write-Host (((Get-Command -Module "$($projectData.ProjectName)" -Syntax) -split [System.Environment]::NewLine).Where({-not [String]::IsNullOrWhiteSpace($_)}).ForEach({'- {0}' -f $_}) | Out-String).TrimEnd()
#endregion

# Display all $WPF_* named variables which were automatically created from XAML element names that have a preceding "_" in their name
Get-FormVariable

#region Event handlers for UI components 

# Event handler for TreeView selection changed
$WPF_treeViewSqlQueries.add_SelectedItemChanged({
    Param ($theSender, $e)
    Write-Verbose ("`$WPF_treeViewSqlQueries.add_SelectedItemChanged(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
    $selectedItem = $WPF_treeViewSqlQueries.SelectedItem
    if ($selectedItem) {
        if ($syncHash.UI.EditMode -eq [EditMode]::Record) {
            $WPF_dataGridSqlQuery.ItemsSource = @($selectedItem)
        } Else {
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQuery']]
            $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
            # $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
            # $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView
            $WPF_dataGridSqlQuery.SelectedItem = $selectedItem
            $WPF_dataGridSqlQuery.ScrollIntoView($selectedItem)
        }

        $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
        $dataView = [System.Data.DataView]::new($tableD.Result[0].Tables[0])

        If (-not [String]::IsNullOrEmpty($selectedItem.Parm_Id) -and $selectedItem.Parm_Id -gt 0) {
            # Record Edit Mode logic
            # Set the RowStateFilter to display only added and modified rows.
            # $dataView.RowStateFilter = ([System.Data.DataViewRowState]::Deleted -bor [System.Data.DataViewRowState]::Added -bor [System.Data.DataViewRowState]::ModifiedCurrent)
            $dataView.RowFilter = "Id = $($selectedItem.Parm_Id)"
    
            # ForEach ($row in $dataView) {Write-Host ($row | FT -AutoSize | Out-String).Trim()}
            Write-Host ($dataView | Format-Table -AutoSize | Out-String).Trim()
            $WPF_dataGridSqlQueryParms.ItemsSource = $dataView
        } else {
            # Table Edit Mode logic
            $dataView.RowFilter = "Id = -1"
            $WPF_dataGridSqlQueryParms.ItemsSource = $dataView
        }
    } Else {
        # $WPF_dataGridSqlQuery.ItemsSource = $syncHash.UI.SqlResults.Tables[0].DefaultView
        $WPF_dataGridSqlQuery.SelectedItem = $null
        $WPF_dataGridSqlQueryParms.ItemsSource = $null
    }
})

# Event handler for TabControl selection changed
$WPF_tabControl.add_SelectionChanged({
    Param ([object]$theSender, [System.Windows.Controls.SelectionChangedEventArgs]$e)
    Write-Verbose ("`$WPF_tabControl.add_SelectionChanged(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
    $selectedTab = $WPF_tabControl.SelectedItem
    Write-Host "`$WPF_TabControl selection changed. Selected tab: ($($selectedTab.Header))"
    
    # Log details of the selected tab
    if ($selectedTab.Header -eq "SQLQuery Editor") {
        Write-Host "Selected SQLQuery Editor tab. Checking elements..."
        
        if (-not $WPF_dataGridSqlQuery) {
            Write-Error "`$WPF_dataGridSqlQuery is null."
        }
        if (-not $WPF_dataGridSqlQueryParms) {
            Write-Error "`$WPF_dataGridSqlQueryParms is null."
        }
        if (-not $WPF_toggleEditMode) {
            Write-Error "`$WPF_toggleEditMode is null."
        }
    }
})
#endregion

#region Form Initialization 

# Event handler for ContentRendered
$syncHash.Form.Add_ContentRendered({
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host "`$syncHash.Form.Add_ContentRendered()"
    Write-Verbose ("`$syncHash.Form.Add_ContentRendered(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()

    # XAML Element Names beginning with underscore _ are auto-converted to Variables with $WPF_ prefix by the Load and Process XAML code region
    # $WPF_treeViewSqlQueries = $syncHash.Form.FindName("treeViewSqlQueries")
    # $WPF_dataGridSqlQuery = $syncHash.Form.FindName("dataGridSqlQuery")
    # $WPF_dataGridSqlQueryParms = $syncHash.Form.FindName("dataGridSqlQueryParms")

    # Ensure elements are correctly instantiated
    if (-not $WPF_treeViewSqlQueries -or -not $WPF_dataGridSqlQuery -or -not $WPF_dataGridSqlQueryParms) {
        Write-Error "One or more elements could not be found in the XAML."
        exit
    }

    # Use the `New-SqlQueryDataSet` function to create and initialize the `SqlQueryDataSet` instance.
    $syncHash.UI.SqlResults = New-SqlQueryDataSet -DisplayResults $false -ConnectionString $syncHash.Params.ConnectionString

    # Load Master Data
    $syncHash.UI.SqlResults.AddQuery('SqlQuery','SELECT * FROM [dbo].[SqlQuery]')
    $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQuery']]
    $tableM.ResultType = 'DataAdapter'
    $syncHash.UI.SqlResults.Execute($tableM)
    Write-Host ($tableM | Out-String) -ForegroundColor Blue
    $WPF_treeViewSqlQueries.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
    $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView

    $WPF_dataGrid.ItemsSource = $tableM.Result[0].Tables[0].DefaultView

    # Load Detail Data
    $syncHash.UI.SqlResults.AddQuery('SqlQueryParms','SELECT * FROM [dbo].[SqlQueryParms]')
    # $tableD = $syncHash.UI.SqlResults[$syncHash.UI.SqlResults.TableIndex]
    $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableNames['SqlQueryParms']]
    $tableD.ResultType = 'DataAdapter'
    $syncHash.UI.SqlResults.Execute($tableD)
    Write-Host ($tableD | Out-String) -ForegroundColor Blue
    $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView
})
#endregion

# Event handler for SourceInitialized
$syncHash.Form.Add_SourceInitialized({
    Write-Host "Form Source Initialized"
    $syncHash.UI.timer = [System.Windows.Threading.DispatcherTimer]::new()
    $syncHash.UI.timer.Interval = [TimeSpan]"0:0:0.150"
    $syncHash.UI.timer.Add_Tick({
        # Update message handler logic here
    })
    $syncHash.UI.timer.Start()
    if($syncHash.UI.timer.IsEnabled) {
        Write-Host "Message Handler is running"
    } else {
        Write-Host "Message Handler Timer didn't start"
    }
})
#endregion

#region Form Closing and Cleanup
# Event handler for Closing
$syncHash.Form.Add_Closing({
    Param ([object]$theSender, [System.ComponentModel.CancelEventArgs]$e)
    Write-Verbose ("`$syncHash.Form.Add_Closing(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
    Write-Host "Form Closing"
    # Check DataGrids for unsaved edits and close connections
    $syncHash.Running = $false
    $syncHash.UI.timer.Stop()
    while ($syncHash.UI.timer.IsEnabled) {
        Write-Host "Shutting down Messages Handler..."
        Start-Sleep -Seconds 1
    }
    If ($syncHash.SqlConnection) {
        $syncHash.SqlConnection.Close()
    }
})

# Event handler for Closed
$syncHash.Form.Add_Closed({
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Verbose ("`$syncHash.Form.Add_Closed(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
    Write-Host "Form Closed"
    $syncHash.Running = $false
    $syncHash.Params.IsRunning = $false
    Stop-Transcript
})
#endregion

# Set the TabControl to display tbSelection
Write-Host "Setting initial UI Component Defaults before Showing Form"
$WPF_tabControl.SelectedIndex = 2

#region Display the Form
# Due to some bizarre bug with ShowDialog and xaml we need to invoke this asynchronously to prevent a segfault
$async = $syncHash.Form.Dispatcher.InvokeAsync({
    $syncHash.Form.ShowDialog() | Out-Null
    $syncHash.Error = $Error 
})
$async.Wait() | Out-Null
#endregion



#=====================================================
<# # >
Add-Type -AssemblyName PresentationFramework

$inputXML = Get-Content -Path "$($syncHash.Params.ScriptsFolder)\SqlQueryEditor.xaml" -Raw

# Cleanup XAML by removing unwanted namespaces and expanding Win to Window
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:Na",'Na' -replace '^<Win.*', '<Window'

# Convert XAML input string to XML Object
[xml]$xaml = $inputXML
$xaml.Window.RemoveAttribute('x:Class')
$xaml.Window.RemoveAttribute('mc:Ignorable')

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$syncHash.Form = [Windows.Markup.XamlReader]::Load( $reader )


# Check for a text changed value (which we cannot parse)
If ($xaml.SelectNodes("//*[@Name]") | Where-Object {$_.TextChanged}) {
    Write-Host ("Error: This Snippet can't convert any lines which contain a 'textChanged' property. `n please manually remove these entries")
    $xaml.SelectNodes("//*[@Name]") | Where-Object {$_.TextChanged} | ForEach-Object { 
        write-warning "Please remove the TextChanged property from this entry $($_.Name)" 
    }
    Return
}

=================================================================
$xaml.Window.xmlns
$xaml.Window.NamespaceURI
$xaml.Window.GetAttribute("xmlns")
$xaml.Window.Attributes
$xaml.Window.GetAttribute("xmlns:x")

# List all attributes of the Window element
foreach ($attribute in ($xaml.Window.Attributes)) {
    Write-Host ('{0}: {1}: {2}' -f $attribute.Name, $attribute.LocalName, $attribute.Value)
}

# Filter and output the 'xmlns' attributes
foreach ($attribute in $xaml.Window.Attributes) {
    if ($attribute.Prefix -eq "xmlns" -or $attribute.Name -eq "xmlns") {
        # Write-Host ('{0}: {1}: {2}' -f $attribute.Name, $attribute.LocalName, $attribute.Value)
        Write-Host ('{0}: {2}' -f $attribute.Name, $attribute.LocalName, $attribute.Value)
    }
}
# xmlns: http://schemas.microsoft.com/winfx/2006/xaml/presentation
# xmlns:x: http://schemas.microsoft.com/winfx/2006/xaml
# xmlns:d: http://schemas.microsoft.com/expression/blend/2008
# xmlns:mc: http://schemas.openxmlformats.org/markup-compatibility/2006


$attributes = $xaml.Window.Attributes
$attributes

$attribute.Name
$attribute.NamespaceURI
$attributes.Name
$attributes.Value
$attributes[1].Name
$attributes[1].LocalName
$attributes[1].NamespaceURI
$attributes[1].Value

# Filter and output the 'xmlns' attributes
$attributes = $xaml.Window.Attributes
foreach ($attribute in $attributes) {
    if ($attribute.Prefix -eq "xmlns" -or $attribute.Name -eq "xmlns") {
        Write-Host ('{0}: {1}: {2}' -f $attribute.Name, $attribute.LocalName, $attribute.Value)
    }
}

# Retrieve the default namespace
$nsDefault = $xaml.Window.NamespaceURI
# Retrieve the 'x' namespace
$nsX = $xaml.Window.GetAttribute("xmlns:x")
# Output the namespaces
$nsDefault
$nsX
$nsManager = New-Object System.Xml.XmlNamespaceManager($xaml.NameTable)
$nsManager.AddNamespace("default", $nsDefault)
$nsManager.AddNamespace("x", $nsX)
$nsManager

$xaml.SelectNodes("//*[@Name]", $nsManager)
$xaml.SelectNodes("//*[@x:Name]", $nsManager) | Select-Object -Property Name
$xaml.SelectNodes("//default:saveButton", $nsManager)
$xaml.SelectNodes("//x:saveButton", $nsManager)

#>