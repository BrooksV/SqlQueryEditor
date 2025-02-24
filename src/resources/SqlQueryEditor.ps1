<#

. "C:\Git\SqlQueryEditor\dist\SqlQueryEditor\SqlQueryEditor.ps1"

#>

# Write-Warning ($error | Out-String)

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
# Initialize-Module -moduleName "SqlQueryClass"
Import-Module C:\Git\SqlQueryClass\dist\SqlQueryClass\SqlQueryClass.psd1 -Force # -Verbose -Force

# Initialize-Module -moduleName "GuiMyPS"
Import-Module C:\Git\GuiMyPS\dist\GuiMyPS\GuiMyPS.psd1 -Force # -Verbose -Force

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
    width = 999
})
$syncHash.Add('UI', [PSCustomObject]@{
    # Set the initial content of the toggle button to reflect Record Edit Mode
    EditMode = [EditMode]::Record # [EditMode]::Table # 
    SqlResults = $null
    LeftColumnPreviousWidth = $null
    timer = $null # [System.Windows.Threading.DispatcherTimer]
    IsCurrentDataGridDirty = $false
    LastSelectedChangeName = $null
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
    Set-Message -Message ("`$handler_menu_Click() Menu Item clicked: {0}" -f $theSender.Name)
    Switch -Regex ($theSender.Name) {
        # Main Menu
        'mnuExit|mnuExit2' { $syncHash.Form.Close()
            Break }
        'mnuQuit' {
            $syncHash.Form.Close()
            Break }
        'mnuMExportCopyToClipboard' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                Save-DataGridContent -InputObject $WPF_dataGridSqlQuery -SaveAs Clipboard -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsClipboard Aborted')
            }
            Break }
        'mnuDExportCopyToClipboard' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                Save-DataGridContent -InputObject $WPF_dataGridSqlQueryParms -SaveAs Clipboard -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsClipboard Aborted')
            }
            Break }
        'mnuMExportAsCSV' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                Save-DataGridContent -InputObject $WPF_dataGridSqlQuery -SaveAs CSV -Path "C:\Temp\dt$($WPF_dataGridSqlQuery.Name).csv"
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsCSV Aborted')
            }
            Break }
        'mnuDExportAsCSV' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                Save-DataGridContent -InputObject $WPF_dataGridSqlQueryParms -SaveAs CSV -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsCSV Aborted')
            }
            Break }
        'mnuMExportAsExcel' {
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                $WPF_dataGridSqlQuery | Save-DataGridContent -SaveAs Excel -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsExcel Aborted')
            }
            Break }
        'mnuDExportAsExcel' {
            If ($WPF_dataGridSqlQueryParms.ItemsSource) {
                $WPF_dataGridSqlQueryParms | Save-DataGridContent -SaveAs Excel -Verbose
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAsExcel Aborted')
            }
            Break }
        'mnuMImportFromExcel' { 
            If ($WPF_dataGridSqlQuery.ItemsSource) {
                $WPF_dataGridSqlQuery.ItemsSource | Save-DatasetToExcel -Path ("{0}\{1}" -f $syncHash.Params.LogFolder, 'SQLQuery_Export.xlsx')
            } Else {
                Set-Message -Message ('DataGrid cannot be empty. SaveAllToExcel Aborted')
            }
            Break }
        'mnuDImportFromExcel' { 
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
            Sync-EveryDataBinding
            $WPF_dataGridSqlQuery.ScrollIntoView(-1)
            Break }
        'mnuDRefresh' {
            $SelectedItem = $WPF_dataGridSqlQueryParms.SelectedItem
            If (-not [String]::IsNullOrEmpty($SelectedItem)) {
                $WPF_dataGridSqlQueryParms.SelectedItem = $SelectedItem
                $WPF_dataGridSqlQueryParms.ScrollIntoView($WPF_dataGridSqlQueryParms.SelectedItem)
            }
            $WPF_menuDetailDataGrid.Items.Where({$_.Name -in @('mnuDEdit','mnuDCancel','mnuDSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            Sync-EveryDataBinding
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
        'mnuNew' {
            $syncHash.UI.SqlResults.TableIndex = $syncHash.UI.SqlResults.TableNames['SqlQuery']
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableIndex]
            $tableM.IsDirty = $true
            $WPF_dataGridSqlQuery.SelectedItem = $WPF_dataGridSqlQuery.ItemsSource.AddNew()
            $WPF_dataGridSqlQuery.SelectedItem.DataSource = 'SqlQuery'
            $WPF_dataGridSqlQuery.SelectedItem.Name = 'All Sorted'
            $WPF_dataGridSqlQuery.SelectedItem.Description = 'Selects All Records Sorted by Date'
            $WPF_dataGridSqlQuery.SelectedItem.SqlFormat = 'SELECT * FROM [dbo].SqlQuery ORDER BY {0} DESC;'
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
            Sync-EveryDataBinding
            $WPF_dataGridSqlQuery.ScrollIntoView(-1)
            Break }
        'mnuDCancel' {
            # Reset MenuItems when Data is Refreshed
            # 'mnuFirst','mnuLast','mnuPrevious','mnuNext','mnuEdit','mnuCancel.Tag','mnuSave','mnuDelete'
            $WPF_menuDetailDataGrid.Items.Where({$_.Name -in @('mnuDEdit','mnuDCancel','mnuDSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
            Sync-EveryDataBinding
            $WPF_dataGridSqlQueryParms.ScrollIntoView(-1)
            Break }
        'mnuMSave' {
            # Commit edits at the row level
            $WPF_dataGridSqlQuery.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
            # Retrieve DataSet / DataTable
            $syncHash.UI.SqlResults.TableIndex = $syncHash.UI.SqlResults.TableNames['SqlQuery']
            $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableIndex]

            Write-Host ('`$WPF_dataGridSqlQuery.SelectedIndex: {0}' -f $WPF_dataGridSqlQuery.SelectedIndex)
            Write-Host ('$tableM.Result[0] DataSet Has Changes: {0}' -f ($tableM.Result[0].HasChanges()))
            Write-Host ('$WPF_dataGridSqlQuery Committing Changes: {0}' -f ($WPF_dataGridSqlQuery.CommitEdit()))
            If (-not $tableM.IsDirty) {
                Write-Host ("Will Save changes even though No changes were made to DataGrid: {0}" -f $WPF_dataGridSqlQuery.Name)
            } Else {
                Write-Host ("Saving Changes made to DataGrid: {0}" -f $WPF_dataGridSqlQuery.Name)
            }

            Try {
                $syncHash.UI.SqlResults.SaveChanges()
            } Catch {
                Write-Host ("Save() Error: ({0})" -f ($error[0] | Out-String).TrimEnd()) -ForegroundColor Red
            } Finally {
                Sync-EveryDataBinding
            }
            # Reset MenuItems when Data is Refreshed
            $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMDelete','mnuMEdit','mnuMCancel','mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})            
            Break }
        'mnuDSave' {
            # Commit edits at the row level
            $WPF_dataGridSqlQueryParms.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
            # Retrieve DataSet / DataTable
            $syncHash.UI.SqlResults.TableIndex = $syncHash.UI.SqlResults.TableNames['SqlQueryParms']
            $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableIndex]

            Write-Host ('`$WPF_dataGridSqlQueryParms.SelectedIndex: {0}' -f $WPF_dataGridSqlQueryParms.SelectedIndex)
            Write-Host ('$tableM.Result[0] DataSet Has Changes: {0}' -f ($tableD.Result[0].HasChanges()))
            Write-Host ('$WPF_dataGridSqlQueryParms Committing Changes: {0}' -f ($WPF_dataGridSqlQueryParms.CommitEdit()))
            If (-not $tableD.IsDirty) {
                Write-Host ("Will Save changes even though No changes were made to DataGrid: {0}" -f $WPF_dataGridSqlQueryParms.Name)
            } Else {
                Write-Host ("Saving Changes made to DataGrid: {0}" -f $WPF_dataGridSqlQueryParms.Name)
            }

            Try {
                $syncHash.UI.SqlResults.SaveChanges()
            } Catch {
                Write-Host ("Save() Error: ({0})" -f ($error[0] | Out-String).TrimEnd()) -ForegroundColor Red
            } Finally {
                Sync-EveryDataBinding
            }
            # Reset MenuItems when Data is Refreshed
            $WPF_menuDetailDataGridParms.Items.Where({$_.Name -in @('mnuMDelete','mnuMEdit','mnuMCancel','mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})            
            Break }
        'mnuMDelete' {
            If ((New-Popup -Buttons YesNo -Icon Question -message "Is it okay to delete this record?" -title 'Okay to Delete') -eq 6) {
                $syncHash.UI.SqlResults.TableIndex = $syncHash.UI.SqlResults.TableNames['SqlQuery']
                $tableM = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableIndex]
                $WPF_dataGridSqlQuery.SelectedItem.Delete()
                $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMDelete','mnuMEdit')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
                $WPF_menuMasterDataGrid.Items.Where({$_.Name -in @('mnuMSave')}).ForEach({$_.Tag = ([MenuItemMode]::Enabled -as [Int]).ToString()})
                $tableM.IsDirty = $true
            }
            Break }
        'mnuDDelete' {
            If ((New-Popup -Buttons YesNo -Icon Question -message "Is it okay to delete this record?" -title 'Okay to Delete') -eq 6) {
                $syncHash.UI.SqlResults.TableIndex = $syncHash.UI.SqlResults.TableNames['SqlQueryParms']
                $tableD = $syncHash.UI.SqlResults.Tables[$syncHash.UI.SqlResults.TableIndex]
                $WPF_dataGridSqlQueryParms.SelectedItem.Delete()
                $WPF_menuDetailDataGridParms.Items.Where({$_.Name -in @('mnuDDelete','mnuDEdit')}).ForEach({$_.Tag = ([MenuItemMode]::Disabled -as [Int]).ToString()})
                $WPF_menuDetailDataGridParms.Items.Where({$_.Name -in @('mnuDSave')}).ForEach({$_.Tag = ([MenuItemMode]::Enabled -as [Int]).ToString()})
                $tableD.IsDirty = $true
            }
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
    Write-Verbose ("`$handler_toggleEditMode() {0}: ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim() 
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
# Write-Host "Helper Functions were Moved to SqlQueryEditor Module"
# Write-Host "Data Management Functions were Moved to SqlQueryEditor Module"
# Write-Host (((Get-Command -Module "$($projectData.ProjectName)" -Syntax) -split [System.Environment]::NewLine).Where({-not [String]::IsNullOrWhiteSpace($_)}).ForEach({'- {0}' -f $_}) | Out-String).TrimEnd()
#endregion

# Display all $WPF_* named variables which were automatically created from XAML element names that have a preceding "_" in their name
Set-Message -Message (Get-FormVariable | Out-String -Width $syncHash.Params.width).TrimEnd() -NewlineBefore
Set-Message -Message ("`$projectData: $($projectData | Out-String -Width $syncHash.Params.width)").TrimEnd() -NewlineBefore
Set-Message -Message ("`$syncHash: $($syncHash | Out-String -Width $syncHash.Params.width)").TrimEnd() -NewlineBefore
Set-Message -Message ("`$syncHash.Params: $($syncHash.Params | Out-String -Width $syncHash.Params.width)").TrimEnd() -NewlineBefore
Set-Message -Message ("`$syncHash.UI: $($syncHash.UI | Out-String -Width $syncHash.Params.width)").TrimEnd() -NewlineBefore

#region Event handlers for UI components 

# Event handler for TreeView selection changed
$WPF_treeViewSqlQueries.add_SelectedItemChanged({
    Param ($theSender, $e)

    If ($syncHash.UI.LastSelectedChangeName -eq '_dataGridSqlQuery') {
        $syncHash.UI.LastSelectedChangeName = $e.Source.Name
        Return
    }
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
            # Write-Host ($dataView | Format-Table -AutoSize | Out-String).Trim()
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

    If ($syncHash.UI.LastSelectedChangeName -eq '_treeViewSqlQueries') {
        $syncHash.UI.LastSelectedChangeName = $e.Source.Name
        Return
    }

    $selectedTab = $WPF_tabControl.SelectedItem

    Switch ($e.OriginalSource.GetType().ToString()) {
        'System.Windows.Controls.TabControl' {
            Set-Message -Message "`$WPF_tabControl.add_SelectionChanged() [$($selectedTab.GetType().ToString())]. Selected tab: ($($selectedTab.Header))"
            Break
        }
        'System.Windows.Controls.DataGrid' {
            If ($e.Source.Name -eq '_dataGridSqlQuery') {
                $selectedItem = $e.Source.SelectedItem
                ForEach ($treeItem in $WPF_treeViewSqlQueries.Items) {
                    If ($treeItem.ToString() -eq '{NewItemPlaceholder}') { # -or [String]::IsNullOrEmpty($treeItem.Id)
                        Continue
                    }
                    If ($treeItem.Id -and $treeItem.Id -eq $selectedItem.Id) {
                        $treeViewItem = [System.Windows.Controls.TreeViewItem]$WPF_treeViewSqlQueries.ItemContainerGenerator.ContainerFromItem($treeItem)
                        $treeViewItem.IsSelected = $true
                        $treeViewItem.Focus()
                        $treeViewItem.BringIntoView()
                        Break
                    }
                }
                #>
            } ElseIf ($e.Source.Name -eq '_dataGrid') {
                Break
            } Else {
                Write-Host ("`$WPF_tabControl.add_SelectionChanged({0}): SourceName ({1}) ({2})" -f $theSender.Name, $e.Source.Name, ($e | Format-List | Out-String)).TrimEnd() -ForegroundColor Magenta
            }
            Break
        }
        Default {
            # Write-Host ("`$WPF_tabControl.add_SelectionChanged({0}): SourceName ({1}) ({2})" -f $theSender.Name, $e.Source.Name, ($e | Format-List | Out-String)).TrimEnd() -ForegroundColor Blue
            Set-Message -Message ("`$WPF_tabControl.add_SelectionChanged([{0}]): Sender ({1}), SourceName ({2}) ({3})" -f $selectedTab.GetType().ToString(), $theSender.Name, $e.Source.Name, ($e | Format-List | Out-String)).TrimEnd()
            Break
        }
    }
})
#endregion

#region Form Initialization 

Function Sync-EveryDataBinding {
    Param ()
    Set-Message -Message ("`Sync-EveryDataBinding()")
    If ([String]::IsNullOrEmpty($syncHash.UI.SqlResults)) {
        $syncHash.UI.SqlResults = New-SqlQueryDataSet -DisplayResults $false -ConnectionString $syncHash.Params.ConnectionString
    }
    # Load Master Data
    $tableIndex = $syncHash.UI.SqlResults.TableNames['SqlQuery']
    If ($null -eq $tableIndex) {
        $tableIndex = $syncHash.UI.SqlResults.AddQuery('SqlQuery','SELECT * FROM [dbo].[SqlQuery]')
    }
    $tableM = $syncHash.UI.SqlResults.Tables[$tableIndex]
    $tableM.ResultType = 'DataAdapter'
    $syncHash.UI.SqlResults.Execute($tableM)
    $tableM.IsDirty = $false
    Write-Verbose ($tableM | Out-String).TrimEnd()
    $WPF_treeViewSqlQueries.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
    $WPF_dataGridSqlQuery.ItemsSource = $tableM.Result[0].Tables[0].DefaultView
    $WPF_dataGrid.ItemsSource = $tableM.Result[0].Tables[0].DefaultView

    # Load Detail Data
    $tableIndex = $syncHash.UI.SqlResults.TableNames['SqlQueryParam']
    If ($null -eq $tableIndex) {
        $tableIndex = $syncHash.UI.SqlResults.AddQuery('SqlQueryParms','SELECT * FROM [dbo].[SqlQueryParms]')
    }
    $tableD = $syncHash.UI.SqlResults.Tables[$tableIndex]
    $tableD.ResultType = 'DataAdapter'
    $syncHash.UI.SqlResults.Execute($tableD)
    $tableD.IsDirty = $false
    Write-Verbose ($tableD | Out-String).TrimEnd()
    $WPF_dataGridSqlQueryParms.ItemsSource = $tableD.Result[0].Tables[0].DefaultView
}

# Event handler for ContentRendered
$syncHash.Form.Add_ContentRendered({
    Param ([object]$theSender, [System.EventArgs]$e)
    # Write-Host "`$syncHash.Form.Add_ContentRendered()"
    # Write-Verbose ("`$syncHash.Form.Add_ContentRendered(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
    Set-Message -Message ("`$syncHash.Form.Add_ContentRendered(): {0} ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim() -NewlineBefore

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
    Sync-EveryDataBinding
})
#endregion

# Event handler for SourceInitialized
$syncHash.Form.Add_SourceInitialized({
    Set-Message -Message ("`$syncHash.Form.Add_SourceInitialized() Form Source Initialized: {0}" -f $theSender.Name) -NewlineBefore
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
# Use Register-ObjectEvent to handle the window's Closing event
<# # >
Register-ObjectEvent -InputObject $syncHash.Form -EventName 'Closing' -Action {
    param($theSender, $e)
    # Ensure any cleanup actions are performed before the window is closed
    Write-Host "Window is closing. Cleaning up resources."
    # Additional cleanup code can be added here if needed
}
#>

$syncHash.Form.Add_Closing({
    Param ([object]$theSender, [System.ComponentModel.CancelEventArgs]$e)
    Write-Host ("`$syncHash.Form.Add_Closing({0}): ({1})" -f $theSender.Name, ($e | Format-List | Out-String)).Trim()
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
    Try {
        Write-Host ("`$syncHash.Form.Add_Closed($($theSender.Name)) Cleaning up resources`n$($e | Format-List | Out-String)").Trim()
        $syncHash.Running = $false
        $syncHash.Params.IsRunning = $false
        Stop-Transcript
    } Catch {
        Write-Warning "An error occurred while handling the Closing event: $_"
    } Finally {
        Write-Host ("`$syncHash.Form.Add_Closed()")
    }
})
#endregion

# Set the TabControl to display tbSelection
Set-Message -Message "Setting initial UI Component Defaults before Showing Form" -NewlineBefore
$WPF_tabControl.SelectedIndex = 1

#region Display the Form
# Due to some bizarre bug with ShowDialog and xaml we need to invoke this asynchronously to prevent a segfault
$async = $syncHash.Form.Dispatcher.InvokeAsync({
    Write-Host "`$syncHash.Form.ShowDialog()"
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