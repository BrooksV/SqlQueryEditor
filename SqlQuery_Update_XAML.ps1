# Define the XAML file path
$xamlFilePath = "C:\Git\SqlQueryEditor\scripts\SqlQueryEditor.xaml"

# Define the new XAML content
$newXamlContent = @"
<Window x:Class="SqlQueryEditor.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    mc:Ignorable="d"
    Title="SqlQueryEditor" Height="800" Width="1200" Topmost="False"
    ResizeMode="CanResizeWithGrip" ShowInTaskbar="True"
    WindowStartupLocation="CenterScreen"
    x:Name="MainForm"
    Background="AliceBlue" UseLayoutRounding="True">

    <Grid>
        <TabControl x:Name="_tabControl" Grid.Row="0" Margin="5" Padding="0"
                    HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                    Background="LightGray" BorderThickness="2" BorderBrush="Gray"
                    SelectedIndex="0" IsSynchronizedWithCurrentItem="True">

            <TabItem x:Name="tbDocumentTree" Tag="0" Header="Document Management" ToolTip="Browsing, Scanning, and Managing Documents">
                <!-- Add Document Management UI components here -->
            </TabItem>
            <TabItem x:Name="tbLkUpTableEdit" Tag="1" Header="Data Grid" ToolTip="Edit Data and Lookup Tables">
                <!-- Add Data Grid UI components here -->
            </TabItem>
            <TabItem x:Name="tbSelection" Tag="2" Header="SQLQuery Editor" ToolTip="Select SQL Queries to Execute, Edit, Delete, or Add to SQL Query Database">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="20*" />
                        <ColumnDefinition Width="5" />
                        <ColumnDefinition Width="75*" />
                    </Grid.ColumnDefinitions>

                    <!-- Toggle Button -->
                    <ToggleButton x:Name="toggleEditMode" Grid.Row="0" Grid.Column="0" Margin="5" Content="Toggle Edit Mode" />

                    <!-- TreeView -->
                    <TreeView x:Name="treeViewSqlQueries" Grid.Row="1" Grid.Column="0" Margin="5">
                        <!-- TreeView items will be dynamically populated from the database -->
                    </TreeView>

                    <!-- GridSplitter -->
                    <GridSplitter Grid.Row="1" Grid.Column="1" Width="5" HorizontalAlignment="Left" VerticalAlignment="Stretch" Background="Gray"/>

                    <!-- Master-Detail Grids -->
                    <Grid Grid.Row="1" Grid.Column="2">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="Auto" />
                            <RowDefinition Height="*" />
                        </Grid.RowDefinitions>

                        <!-- Master DataGrid Menu -->
                        <Menu Grid.Row="0">
                            <MenuItem Header="File">
                                <MenuItem Header="Export">
                                    <MenuItem Header="Copy to Clipboard" />
                                    <MenuItem Header="Export as CSV" />
                                    <MenuItem Header="Export as Excel" />
                                    <MenuItem Header="Import from Excel" />
                                </MenuItem>
                                <MenuItem Header="Add" />
                                <MenuItem Header="Cancel" />
                                <MenuItem Header="Save" />
                                <MenuItem Header="Delete" />
                            </MenuItem>
                            <MenuItem Header="Edit">
                                <MenuItem Header="First" />
                                <MenuItem Header="Previous" />
                                <MenuItem Header="Next" />
                                <MenuItem Header="Last" />
                            </MenuItem>
                        </Menu>

                        <!-- Master DataGrid -->
                        <DataGrid x:Name="dataGridSqlQuery" Grid.Row="1" AutoGenerateColumns="True" Margin="5">
                            <!-- DataGrid for Master SqlQuery records -->
                        </DataGrid>

                        <!-- Detail DataGrid Menu -->
                        <Menu Grid.Row="2">
                            <MenuItem Header="Actions">
                                <MenuItem Header="Add" />
                                <MenuItem Header="Cancel" />
                                <MenuItem Header="Save" />
                                <MenuItem Header="Delete" />
                            </MenuItem>
                        </Menu>

                        <!-- Detail DataGrid -->
                        <DataGrid x:Name="dataGridSqlQueryParms" Grid.Row="3" AutoGenerateColumns="True" Margin="5">
                            <!-- DataGrid for Detail SqlQueryParms records -->
                        </DataGrid>
                    </Grid>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

# Write the new content to the XAML file
Set-Content -Path $xamlFilePath -Value $newXamlContent -Force

Write-Host "SqlQueryEditor.xaml has been successfully updated."
