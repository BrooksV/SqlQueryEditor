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

# Add required .Net Assemblies required for WPF
Add-Type -AssemblyName PresentationFramework

# Define variable `$syncHash` for global use to store synchronized data and provide access to object data.
$syncHash = [System.Collections.Hashtable]::Synchronized((New-Object System.Collections.Hashtable))
$syncHash.Add('UI', [PSCustomObject]@{
    SqlResults = $null
    LeftColumnPreviousWidth = $null
})

# Database Configuration
# Using sample database configuration data from tests `.\tests\TestDatabase1.parameters.psd1` and the SQL Express test database `.\tests\TestDatabase1.mdf`.

$SqlServer = '(localdb)\MSSQLLocalDB'
$Database = 'TestDatabase1'
$ConnectionString = "Data Source=$SqlServer;AttachDbFilename=C:\Git\SqlQueryClass\tests\TestDatabase1.mdf;Integrated Security=True"

# Use the `New-SqlQueryDataSet` function to create and initialize the `SqlQueryDataSet` instance.
$syncHash.UI.SqlResults = New-SqlQueryDataSet -SQLServer $SqlServer -Database $Database -ConnectionString $ConnectionString

$xamlString = @"
<Window x:Class="SqlQueryClass.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:system="clr-namespace:System;assembly=mscorlib"
    xmlns:local="clr-namespace:SqlQueryClass"
    mc:Ignorable="d"
    Title="SQL Query DataGrid" 
    Height="800" Width="1200" Topmost="False" 
    ResizeMode="CanResizeWithGrip" ShowInTaskbar = "True"
    WindowStartupLocation = "CenterScreen"
    x:Name="MainForm"
    FocusManager.FocusedElement="{Binding ElementName=_scriptView}"
    Background="AliceBlue" UseLayoutRounding="True"
    >
    <!-- UI Style Resources -->
    <Window.Resources>
        <SolidColorBrush x:Key="WindowBackgroundBrush" Color="AliceBlue" />
        <SolidColorBrush x:Key="SolidBorderBrush" Color="DarkBlue" />
        <SolidColorBrush x:Key="SolidExpanderHeaderBrush" Color="DarkGreen" />

        <LinearGradientBrush x:Key="NormalBrush" StartPoint="0,0" EndPoint="0,1">
            <GradientBrush.GradientStops>
                <GradientStopCollection>
                    <GradientStop Color="#FFFFD190" Offset="0.2"/>
                    <GradientStop Color="Orange" Offset="0.85"/>
                    <GradientStop Color="#FFFFD190" Offset="1"/>
                </GradientStopCollection>
            </GradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Resources>

    <!-- UI Components -->
    <Grid x:Name="TableEditGrid">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>

        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" x:Name="_LeftColumn"/>
            <ColumnDefinition Width="6"/>
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>

        <DockPanel Name="dpTableEdit" Grid.Row="0" Grid.RowSpan="1" Grid.Column="0" Grid.ColumnSpan="3" VerticalAlignment="Stretch" Margin="0" Background="LightBlue"/>
        <GridSplitter Name="gridSplitterTableEdit" ShowsPreview="True" Grid.Row="0" Grid.RowSpan="1" Grid.Column="1" Grid.ColumnSpan="1" Margin="0" Padding="0" 
            ResizeDirection="Auto" Height="Auto" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="{StaticResource NormalBrush}" 
            Width="6" />
        <TextBlock Grid.Row="0" Grid.Column="1" Grid.ColumnSpan="1" FontSize="9" Height="Auto" Text=" « ∙ ∙ ∙ ∙ ∙ ∙ ∙ ∙ » " FontWeight="Black" Foreground="DarkBlue" Background="#BBC5D7" Tag="Vertical" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False">
            <TextBlock.LayoutTransform>
                <RotateTransform Angle="90" />
            </TextBlock.LayoutTransform> 
        </TextBlock>

        <Grid Name="RightSideGrid" Grid.Row="0" Grid.Column="2" Grid.ColumnSpan="1" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="0" Background="PaleGoldenrod" >
            <!-- Right of the Splitter -->
            <DockPanel x:Name="dpRightSide" Margin="-2,-2,0,0">
                <ScrollViewer>
                    <StackPanel>
                        <Expander Name="expRightSide" Tag="Right Side Expander" 
                            ExpandDirection="Down" IsExpanded="True" HorizontalAlignment="Stretch" Height="Auto" Width="Auto" 
                            BorderThickness="1" BorderBrush="{StaticResource SolidBorderBrush}" Background="{StaticResource WindowBackgroundBrush}"
                            Margin="5,5,5,5" Padding="2,2,2,2" 
                            HorizontalContentAlignment="Stretch" VerticalContentAlignment="Top" 
                            >
                            <Expander.Header>
                            <DockPanel VerticalAlignment="Stretch">
                                <TextBlock Text="{Binding Tag, RelativeSource={RelativeSource AncestorType={x:Type Expander}}}" VerticalAlignment="Top" Height="14" FontSize="11" FontWeight="Bold" Foreground="{StaticResource SolidExpanderHeaderBrush}" DockPanel.Dock="Left" />
                            </DockPanel>
                            </Expander.Header>
                            <StackPanel x:Name="spRightSide" Orientation="Vertical" MinHeight="25" Height="Auto" Margin="0,0,0,0" HorizontalAlignment="Stretch" >
                                <!-- RightSide Components Go Here -->
                                <DataGrid x:Name='_dataGrid' AutoGenerateColumns='True' />
                            </StackPanel>
                        </Expander>
                    </StackPanel>
                </ScrollViewer>
            </DockPanel>
        </Grid>
        <!-- **********************************************************************
            Left of the Splitter - Controls and Layout
        ********************************************************************** -->
        <Expander Name="_expLeftSide" Grid.Column="0" Grid.ColumnSpan="1" Grid.Row="0" Grid.RowSpan="1"
            ExpandDirection="Left" IsExpanded="True" HorizontalAlignment="Stretch" Height="Auto" Width="Auto" 
            BorderThickness="1" BorderBrush="{StaticResource SolidBorderBrush}" Background="{StaticResource WindowBackgroundBrush}" 
            Margin="5,5,5,5" Padding="0" 
            HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch" 
            >
            <Expander.Header>
            <DockPanel VerticalAlignment="Stretch">
                <TextBlock Text="{Binding Tag, RelativeSource={RelativeSource AncestorType={x:Type Expander}}}" VerticalAlignment="Top" Height="14" FontSize="11" FontWeight="Bold" Foreground="{StaticResource SolidExpanderHeaderBrush}" DockPanel.Dock="Left" />
            </DockPanel>
            </Expander.Header>

            <Grid x:Name="_LeftSideGrid"  Margin="0,0,0,0"  
                    HorizontalAlignment="Stretch" MaxWidth="{Binding ElementName=MainWindow, Path=ActualWidth}" 
                    VerticalAlignment="Stretch" MaxHeight="{Binding ElementName=MainWindow, Path=ActualHeight}"
                    >
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" MinWidth="300"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Menu / Options -->
                <StackPanel x:Name="spLeftSideOptions" Margin="0,0,0,0" Grid.Row="0" Grid.RowSpan="1" Grid.Column="0" Grid.ColumnSpan="2" 
                    HorizontalAlignment="Stretch" VerticalAlignment="Stretch"
                    >
                    <GroupBox x:Name="gbxLeftSideOptions" Header="SQL Query Editor Options" 
                            HorizontalAlignment="Stretch"
                            VerticalAlignment="Top" Height="Auto" 
                            Margin="0,0,0,0" Padding="0,0,0,0">
                        <StackPanel x:Name="spLeftSideMenu" Margin="2" Orientation="Vertical" VerticalAlignment="Top" HorizontalAlignment="Stretch">
                            <Menu x:Name="_LeftSideMenu" Padding="3,3,3,3" Background="#FFF3EDE0" BorderBrush="LightBlue" BorderThickness="1">
                                <MenuItem x:Name="mnuExit" Header="Exit"/>
                            </Menu>
                        </StackPanel>
                    </GroupBox>
                </StackPanel>
                <!-- **********************************************************************
                Left Side of Splitter
                ********************************************************************** -->
                <Border x:Name="bdrLeftSide" BorderThickness="2" BorderBrush="{StaticResource NormalBrush}" Margin="0,0,0,0" Grid.Row="1" Grid.RowSpan="2" Grid.Column="0" Grid.ColumnSpan="3" 
                        HorizontalAlignment="Stretch" VerticalAlignment="Top"
                        >
                    <!-- Place Components Here -->
                </Border>
            </Grid>
        </Expander>
        <!-- End of Directory Tree Grid - Left Side of Splitter -->
    </Grid>
</Window>
"@

$handler_Button_Click = {
    Param ([object]$theSender, [System.EventArgs]$e)
    Write-Host ("`$handler_Button_Click() Item clicked: {0}" -f $theSender.Name)
    Switch -Regex ($theSender.Name) {
            '^mnuExit$' {
                $rootElement = Find-RootElement -Element $theSender
                If ($rootElement) {
                    $rootElement.Close()
                }
                Break
            }
        default {
            Write-Host ("{0}: {1}({2})" -f $theSender.Name, $e.OriginalSource.Name, $e.OriginalSource.ToString())
        }
    }
}

try {
    $syncHash.Form = New-XamlWindow -xaml $xamlString
    $elements = @()
    $elements += Find-EveryControl -Element $syncHash.Form -ControlType 'System.Windows.Controls.MenuItem'
    $elements.ForEach({$_.Element.Add_Click($handler_Button_Click)})

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

    $syncHash.UI.SqlResults.DisplayResults = $false
    $syncHash.UI.SqlResults.ExecuteQuery("SELECT * FROM [dbo].[SqlQuery]")
    $WPF_dataGrid.ItemsSource = $syncHash.UI.SqlResults.Tables[0].Result.DefaultView

    $syncHash.Form.ShowDialog()
} catch {
    Write-Warning ($_ | Format-List | Out-String)
}
