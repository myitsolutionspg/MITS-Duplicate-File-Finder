<# 
  MITS-DuplicateFileFinder-GUI_v1.1.0.ps1
  PS 5.1 WPF GUI - displays results from memory; writes audit CSVs for reporting.
  Fixes: reads ScanStats from scanner, detects "0 files scanned" early, robust filtering.
#>

#requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName Microsoft.VisualBasic

Set-StrictMode -Version 2

# StrictMode-safe async scan state (used by DispatcherTimer callbacks)
$script:psScan = $null
$script:asyncScan = $null
$script:scanTimer = $null
$ErrorActionPreference = "Stop"

# -------------------- Paths / Branding --------------------
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir     = Join-Path $ScriptRoot "out"
$AssetsDir  = Join-Path $ScriptRoot "assets"
$LogoPath   = Join-Path $AssetsDir "mits_logo.png"
$ScannerScript = Join-Path $ScriptRoot "Find-DuplicateFiles.ps1"
$VersionLabel = "v1.1.0"

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MITS Duplicate File Finder"
        Height="860" Width="1200" MinHeight="780" MinWidth="1100" WindowStartupLocation="CenterScreen">
  <Grid Background="#F5F7FA" Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="3*"/>
      <RowDefinition Height="2*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" CornerRadius="14" Background="#7CC7E9" Padding="14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <Border Width="62" Height="62" CornerRadius="12" Background="#FFFFFF" Padding="8" VerticalAlignment="Center">
          <Image x:Name="imgLogo" Stretch="Uniform"/>
        </Border>

        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <TextBlock Text="Duplicate File Finder" FontSize="28" FontWeight="SemiBold" HorizontalAlignment="Center"/>
          <TextBlock x:Name="txtVersion" Text="vX" FontSize="12" Opacity="0.8" HorizontalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Column="2" VerticalAlignment="Center" HorizontalAlignment="Right">
          <TextBlock Text="By Melky Warinak" FontSize="12" HorizontalAlignment="Right"/>
          <TextBlock>
            <Hyperlink NavigateUri="https://myitsolutionspg.com" >
              myitsolutionspg.com
            </Hyperlink>
          </TextBlock>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Scan settings -->
    <Border Grid.Row="1" Margin="0,12,0,12" Background="#FFFFFF" CornerRadius="14" Padding="18">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Row="0" Grid.Column="0">
          <TextBlock Text="Scan target (folder or drive)" FontWeight="SemiBold" Margin="0,0,0,6"/>
          <TextBox x:Name="txtPath" Height="30" FontSize="13" />
        </StackPanel>

        <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="12,0,0,0">
          <Button x:Name="btnBrowse" Content="Browse..." Width="110" Height="30" Margin="0,0,10,0"/>
          <Button x:Name="btnOpenOut" Content="Open Output Folder" Width="150" Height="30"/>
        </StackPanel>

        <WrapPanel Grid.Row="1" Grid.ColumnSpan="2" Margin="0,12,0,0" VerticalAlignment="Center">
          <CheckBox x:Name="chkRecurse" Content="Recurse" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkConfirm" Content="ConfirmContent (SHA256)" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkConfirmedOnly" Content="ConfirmedOnly" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkHidden" Content="IncludeHidden/System" IsChecked="False" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <TextBlock Text="Show:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <RadioButton x:Name="rbShowAll" Content="ALL" GroupName="ShowMode" IsChecked="False" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <RadioButton x:Name="rbShowRemove" Content="REMOVE" GroupName="ShowMode" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <RadioButton x:Name="rbShowKeep" Content="KEEP" GroupName="ShowMode" IsChecked="False" Margin="0,0,18,0" VerticalAlignment="Center"/>

          <TextBlock Text="KeepRule:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <ComboBox x:Name="cmbKeepRule" Width="160" SelectedIndex="0" Margin="0,0,18,0" VerticalAlignment="Center">
            <ComboBoxItem Content="ShortestPath"/>
            <ComboBoxItem Content="LongestPath"/>
            <ComboBoxItem Content="NewestWriteTime"/>
            <ComboBoxItem Content="OldestWriteTime"/>
            <ComboBoxItem Content="NewestCreationTime"/>
            <ComboBoxItem Content="OldestCreationTime"/>
          </ComboBox>

          <TextBlock Text="MinSizeMB:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <TextBox x:Name="txtMinSize" Width="70" Text="0" Margin="0,0,18,0" VerticalAlignment="Center"/>

          <TextBlock Text="SampleRemovePaths:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <TextBox x:Name="txtSampleRemove" Width="50" Text="5"/>
        </WrapPanel>

        <WrapPanel Grid.Row="2" Grid.ColumnSpan="2" Margin="0,12,0,0" VerticalAlignment="Center">
          <Button x:Name="btnScan" Content="Scan" Width="110" Height="34" Margin="0,0,12,0"/>
          <ProgressBar x:Name="pb" Width="240" Height="16" IsIndeterminate="False" Margin="0,0,12,0"/>
          <TextBlock x:Name="txtStatus" Text="Idle" VerticalAlignment="Center" Margin="0,0,24,0"/>

          <Button x:Name="btnSelectAllRemove" Content="Select All REMOVE" Width="160" Height="34" Margin="0,0,10,0"/>
          <Button x:Name="btnClearSel" Content="Clear Selection" Width="130" Height="34" Margin="0,0,20,0"/>

          <CheckBox x:Name="chkWhatIf" Content="WhatIf (Dry run)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>
          <CheckBox x:Name="chkRecycle" Content="Delete to Recycle Bin (safer)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>

          <Button x:Name="btnQuarantine" Content="Move Selected to Quarantine" Width="220" Height="34" Margin="0,8,10,0"/>
          <Button x:Name="btnDelete" Content="Delete Selected" Width="140" Height="34" Margin="0,8,0,0" Background="#FFD9D9"/>
        </WrapPanel>
      </Grid>
    </Border>

    <!-- Results -->
    <Border Grid.Row="2" Background="#FFFFFF" CornerRadius="14" Padding="14" >
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

          <TextBlock x:Name="txtCounts" Text="Results: (no data loaded)" VerticalAlignment="Center" FontWeight="SemiBold"/>

          <TextBlock Grid.Column="1" Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="txtFilter" Grid.Column="2" Width="320" Height="28"/>
          <Button x:Name="btnClearFilter" Grid.Column="3" Width="32" Height="28" Content="X" Padding="0" FontSize="12" FontWeight="SemiBold" Margin="10,0,0,0" ToolTip="Clear filter"/>
        </Grid>

        <DataGrid x:Name="dg"
                  MinHeight="150"
                  ScrollViewer.HorizontalScrollBarVisibility="Visible"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
Grid.Row="1"
                  AutoGenerateColumns="False"
                  CanUserAddRows="False"
                  IsReadOnly="False"
                  SelectionMode="Extended"
                  SelectionUnit="FullRow"
                  HeadersVisibility="Column"
                  Margin="0,10,0,0">
                <DataGrid.ContextMenu>
                  <ContextMenu>
                    <MenuItem x:Name="miCopyFullPath" Header="Copy FullPath"/>
                  </ContextMenu>
                </DataGrid.ContextMenu>
          <DataGrid.Columns>
            <DataGridCheckBoxColumn Header="Select" Binding="{Binding Selected, Mode=TwoWay}" Width="70"/>
            <DataGridTextColumn Header="Confidence" Binding="{Binding Confidence}" Width="110"/>
            <DataGridTextColumn Header="Action" Binding="{Binding Action}" Width="90"/>
            <DataGridTextColumn Header="SizeMB" Binding="{Binding SizeMB}" Width="80"/>
            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="220"/>
            <DataGridTextColumn Header="FullPath" Binding="{Binding FullPath}" Width="2*">
  <DataGridTextColumn.ElementStyle>
    <Style TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="ToolTip" Value="{Binding FullPath}"/>
    </Style>
  </DataGridTextColumn.ElementStyle>
</DataGridTextColumn>
<DataGridTextColumn Header="KeepPath" Binding="{Binding KeepPath}" Width="2*">
  <DataGridTextColumn.ElementStyle>
    <Style TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="ToolTip" Value="{Binding KeepPath}"/>
    </Style>
  </DataGridTextColumn.ElementStyle>
</DataGridTextColumn>
<DataGridTextColumn Header="GroupId" Binding="{Binding GroupId}" Width="140" />
<DataGridTextColumn Header="FullHash" Binding="{Binding FullHash}" Width="2*">
  <DataGridTextColumn.ElementStyle>
    <Style TargetType="TextBlock">
      <Setter Property="TextWrapping" Value="NoWrap"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
      <Setter Property="ToolTip" Value="{Binding FullHash}"/>
    </Style>
  </DataGridTextColumn.ElementStyle>
</DataGridTextColumn>
          </DataGrid.Columns>
        </DataGrid>
      </Grid>
    </Border>

    <!-- Logs -->
    <Border Grid.Row="3" Margin="0,12,0,12" Background="#FFFFFF" CornerRadius="14" Padding="14">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Logs" FontWeight="SemiBold"/>
        <TextBox x:Name="txtLog" Grid.Row="1" FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
      </Grid>
    </Border>

    <!-- Disclaimer -->
    <Border Grid.Row="4" Background="#7CC7E9" CornerRadius="14" Padding="10">
      <TextBlock Text="Disclaimer: This tool can move/delete files. Always review the Summary CSV first and use Quarantine (Move) before deletion. My IT Solutions (PNG) is not liable for data loss due to misuse."
                 TextWrapping="Wrap"/>
    </Border>

  </Grid>
</Window>
"@
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$GuiLog = Join-Path $OutDir ("gui_log_{0}.txt" -f (Get-Date).ToString("yyyyMMdd_HHmmss"))

function Write-GuiLog {
  param([string]$Message)
  $line = "[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Message
  try { Add-Content -LiteralPath $GuiLog -Value $line -Encoding UTF8 } catch {}
  try { if ($script:txtLog) { $script:txtLog.AppendText($line + [Environment]::NewLine); $script:txtLog.ScrollToEnd() } } catch {}
}

# -------------------- XAML --------------------
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="MITS Duplicate File Finder"
        Height="860" Width="1200" WindowStartupLocation="CenterScreen">
  <Grid Background="#F5F7FA" Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="190"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" CornerRadius="14" Background="#7CC7E9" Padding="14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <Border Width="62" Height="62" CornerRadius="12" Background="#FFFFFF" Padding="8" VerticalAlignment="Center">
          <Image x:Name="imgLogo" Stretch="Uniform"/>
        </Border>

        <StackPanel Grid.Column="1" VerticalAlignment="Center">
          <TextBlock Text="Duplicate File Finder" FontSize="28" FontWeight="SemiBold" HorizontalAlignment="Center"/>
          <TextBlock x:Name="txtVersion" Text="vX" FontSize="12" Opacity="0.8" HorizontalAlignment="Center"/>
        </StackPanel>

        <StackPanel Grid.Column="2" VerticalAlignment="Center" HorizontalAlignment="Right">
          <TextBlock Text="By Melky Warinak" FontSize="12" HorizontalAlignment="Right"/>
          <TextBlock>
            <Hyperlink NavigateUri="https://myitsolutionspg.com" >
              myitsolutionspg.com
            </Hyperlink>
          </TextBlock>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Scan settings -->
    <Border Grid.Row="1" Margin="0,12,0,12" Background="#FFFFFF" CornerRadius="14" Padding="14">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Row="0" Grid.Column="0">
          <TextBlock Text="Scan target (folder or drive)" FontWeight="SemiBold" Margin="0,0,0,6"/>
          <TextBox x:Name="txtPath" Height="30" FontSize="13" />
        </StackPanel>

        <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="12,0,0,0">
          <Button x:Name="btnBrowse" Content="Browse..." Width="110" Height="30" Margin="0,0,10,0"/>
          <Button x:Name="btnOpenOut" Content="Open Output Folder" Width="150" Height="30"/>
        </StackPanel>

        <WrapPanel Grid.Row="1" Grid.ColumnSpan="2" Margin="0,12,0,0" VerticalAlignment="Center">
          <CheckBox x:Name="chkRecurse" Content="Recurse" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkConfirm" Content="ConfirmContent (SHA256)" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkConfirmedOnly" Content="ConfirmedOnly" IsChecked="True" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <CheckBox x:Name="chkHidden" Content="IncludeHidden/System" IsChecked="False" Margin="0,0,18,0" VerticalAlignment="Center"/>
          <TextBlock Text="Show:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <RadioButton x:Name="rbShowAll" Content="ALL" GroupName="ShowMode" IsChecked="False" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <RadioButton x:Name="rbShowRemove" Content="REMOVE" GroupName="ShowMode" IsChecked="True" Margin="0,0,10,0" VerticalAlignment="Center"/>
          <RadioButton x:Name="rbShowKeep" Content="KEEP" GroupName="ShowMode" IsChecked="False" Margin="0,0,18,0" VerticalAlignment="Center"/>

          <TextBlock Text="KeepRule:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <ComboBox x:Name="cmbKeepRule" Width="160" SelectedIndex="0" Margin="0,0,18,0" VerticalAlignment="Center">
            <ComboBoxItem Content="ShortestPath"/>
            <ComboBoxItem Content="LongestPath"/>
            <ComboBoxItem Content="NewestWriteTime"/>
            <ComboBoxItem Content="OldestWriteTime"/>
            <ComboBoxItem Content="NewestCreationTime"/>
            <ComboBoxItem Content="OldestCreationTime"/>
          </ComboBox>

          <TextBlock Text="MinSizeMB:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <TextBox x:Name="txtMinSize" Width="70" Text="0" Margin="0,0,18,0" VerticalAlignment="Center"/>

          <TextBlock Text="SampleRemovePaths:" VerticalAlignment="Center" Margin="10,0,6,0"/>
          <TextBox x:Name="txtSampleRemove" Width="50" Text="5"/>
        </WrapPanel>

        <WrapPanel Grid.Row="2" Grid.ColumnSpan="2" Margin="0,12,0,0" VerticalAlignment="Center">
          <Button x:Name="btnScan" Content="Scan" Width="110" Height="34" Margin="0,0,12,0"/>
          <ProgressBar x:Name="pb" Width="240" Height="16" IsIndeterminate="False" Margin="0,0,12,0"/>
          <TextBlock x:Name="txtStatus" Text="Idle" VerticalAlignment="Center" Margin="0,0,24,0"/>

          <Button x:Name="btnSelectAllRemove" Content="Select All REMOVE" Width="160" Height="34" Margin="0,0,10,0"/>
          <Button x:Name="btnClearSel" Content="Clear Selection" Width="130" Height="34" Margin="0,0,20,0"/>

          <CheckBox x:Name="chkWhatIf" Content="WhatIf (Dry run)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>
          <CheckBox x:Name="chkRecycle" Content="Delete to Recycle Bin (safer)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,18,0"/>

          <Button x:Name="btnQuarantine" Content="Move Selected to Quarantine" Width="220" Height="34" Margin="0,8,10,0"/>
          <Button x:Name="btnDelete" Content="Delete Selected" Width="140" Height="34" Margin="0,8,0,0" Background="#FFD9D9"/>
        </WrapPanel>
      </Grid>
    </Border>

    <!-- Results -->
    <Border Grid.Row="2" Background="#FFFFFF" CornerRadius="14" Padding="14" >
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>

          <TextBlock x:Name="txtCounts" Text="Results: (no data loaded)" VerticalAlignment="Center" FontWeight="SemiBold"/>

          <TextBlock Grid.Column="1" Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="txtFilter" Grid.Column="2" Width="320" Height="28"/>
          <Button x:Name="btnClearFilter" Grid.Column="3" Width="32" Height="28" Content="X" Padding="0" FontSize="12" FontWeight="SemiBold" Margin="10,0,0,0" ToolTip="Clear filter"/>
        </Grid>

        <DataGrid x:Name="dg"
                  MinHeight="150"
                  ScrollViewer.HorizontalScrollBarVisibility="Auto"
                  ScrollViewer.VerticalScrollBarVisibility="Auto"
Grid.Row="1"
                  AutoGenerateColumns="False"
                  CanUserAddRows="False"
                  IsReadOnly="False"
                  SelectionMode="Extended"
                  SelectionUnit="FullRow"
                  HeadersVisibility="Column"
                  Margin="0,10,0,0">
                <DataGrid.ContextMenu>
                  <ContextMenu>
                    <MenuItem x:Name="miCopyFullPath" Header="Copy FullPath"/>
                  </ContextMenu>
                </DataGrid.ContextMenu>
          <DataGrid.Columns>
            <DataGridCheckBoxColumn Header="Select" Binding="{Binding Selected, Mode=TwoWay}" Width="70"/>
            <DataGridTextColumn Header="Confidence" Binding="{Binding Confidence}" Width="110"/>
            <DataGridTextColumn Header="Action" Binding="{Binding Action}" Width="90"/>
            <DataGridTextColumn Header="SizeMB" Binding="{Binding SizeMB}" Width="80"/>
            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="220"/>
            <DataGridTextColumn Header="FullPath" Binding="{Binding FullPath}" Width="*"/>
            <DataGridTextColumn Header="KeepPath" Binding="{Binding KeepPath}" Width="*"/>
            <DataGridTextColumn Header="GroupId" Binding="{Binding GroupId}" Width="120"/>
            <DataGridTextColumn Header="FullHash" Binding="{Binding FullHash}" Width="250"/>
          </DataGrid.Columns>
        </DataGrid>
      </Grid>
    </Border>

    <!-- Logs -->
    <Border Grid.Row="3" Margin="0,12,0,12" Background="#FFFFFF" CornerRadius="14" Padding="14">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Logs" FontWeight="SemiBold"/>
        <TextBox x:Name="txtLog" Grid.Row="1" FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
      </Grid>
    </Border>

    <!-- Disclaimer -->
    <Border Grid.Row="4" Background="#7CC7E9" CornerRadius="14" Padding="10">
      <TextBlock Text="Disclaimer: This tool can move/delete files. Always review the Summary CSV first and use Quarantine (Move) before deletion. My IT Solutions (PNG) is not liable for data loss due to misuse."
                 TextWrapping="Wrap"/>
    </Border>

  </Grid>
</Window>
"@

# -------------------- Load Window --------------------
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Link handler for Hyperlink
$null = $window.AddHandler([System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
  [System.Windows.Navigation.RequestNavigateEventHandler]{
    param($s,$e)
    Start-Process $e.Uri.AbsoluteUri | Out-Null
    $e.Handled = $true
  }
)

# Find controls
$imgLogo = $window.FindName("imgLogo")
$txtVersion = $window.FindName("txtVersion")
$txtVersion.Text = $VersionLabel

$txtPath = $window.FindName("txtPath")
$btnBrowse = $window.FindName("btnBrowse")
$btnOpenOut = $window.FindName("btnOpenOut")

$chkRecurse = $window.FindName("chkRecurse")
$chkConfirm = $window.FindName("chkConfirm")
$chkConfirmedOnly = $window.FindName("chkConfirmedOnly")
$chkHidden = $window.FindName("chkHidden")
$rbShowAll   = $window.FindName("rbShowAll")
$rbShowRemove= $window.FindName("rbShowRemove")
$rbShowKeep  = $window.FindName("rbShowKeep")
$cmbKeepRule = $window.FindName("cmbKeepRule")
$txtMinSize = $window.FindName("txtMinSize")
$txtSampleRemove = $window.FindName("txtSampleRemove")

$btnScan = $window.FindName("btnScan")
$pb = $window.FindName("pb")
$txtStatus = $window.FindName("txtStatus")

$btnSelectAllRemove = $window.FindName("btnSelectAllRemove")
$btnClearSel = $window.FindName("btnClearSel")
$chkWhatIf = $window.FindName("chkWhatIf")
$chkRecycle = $window.FindName("chkRecycle")
$btnQuarantine = $window.FindName("btnQuarantine")
$btnDelete = $window.FindName("btnDelete")

$dg = $window.FindName("dg")
$txtCounts = $window.FindName("txtCounts")
$txtFilter = $window.FindName("txtFilter")
$btnClearFilter = $window.FindName("btnClearFilter")
$script:txtLog = $window.FindName("txtLog")

# Load logo
try {
  if (Test-Path -LiteralPath $LogoPath) {
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.UriSource = New-Object System.Uri($LogoPath)
    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.EndInit()
    $imgLogo.Source = $bmp
    Write-GuiLog "Loaded logo: $LogoPath"
  } else {
    Write-GuiLog "Logo not found (optional): $LogoPath"
  }
} catch {
  Write-GuiLog "Logo load failed: $($_.Exception.Message)"
}

Write-GuiLog "Ready."
Write-GuiLog "GUI log file: $GuiLog"
Write-GuiLog "Scanner: $ScannerScript"
Write-GuiLog "Outputs: $OutDir"

# -------------------- Data Binding --------------------
$script:Results = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$script:ResultsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Results)
$dg.ItemsSource = $script:ResultsView


function Get-RowProp {
  param(
    [Parameter(Mandatory=$true)]$Row,
    [Parameter(Mandatory=$true)][string]$Name
  )
  try {
    if ($null -eq $Row) { return $null }
    $p = $Row.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
  } catch {
    return $null
  }
}

function Apply-GridFilter {
  try {
    # Store filter state in script-scope so WPF's ICollectionView.Filter (invoked later) can read it safely under StrictMode
    $script:FilterText = ([string]$txtFilter.Text)
    if ($script:FilterText -eq $null) { $script:FilterText = "" }
    $script:FilterText = $script:FilterText.Trim().ToLowerInvariant()

    $script:FilterMode = $(if ($rbShowKeep.IsChecked -eq $true) { "KEEP" } elseif ($rbShowAll.IsChecked -eq $true) { "ALL" } else { "REMOVE" })
    if ($null -eq $script:ResultsView) { return }
if (-not $script:ResultsView.Filter) {
      # Assign the filter ONCE; it reads script-scope state that we update above.
      $script:ResultsView.Filter = {
        param($row)
        try {
          if ($null -eq $row) { return $false }

          if ($script:FilterMode -eq "REMOVE") {
            if (([string](Get-RowProp -Row $row -Name "Action")) -ne "REMOVE") { return $false }
          } elseif ($script:FilterMode -eq "KEEP") {
            if (([string](Get-RowProp -Row $row -Name "Action")) -ne "KEEP") { return $false }
          }

          if ([string]::IsNullOrWhiteSpace($script:FilterText)) { return $true }

          $hay = (([string](Get-RowProp -Row $row -Name "Name")) + " " + ([string](Get-RowProp -Row $row -Name "FullPath")) + " " + ([string](Get-RowProp -Row $row -Name "KeepPath")) + " " + ([string](Get-RowProp -Row $row -Name "GroupId")) + " " + ([string](Get-RowProp -Row $row -Name "FullHash"))).ToLowerInvariant()
          return ($hay -like ("*" + $script:FilterText + "*"))
        } catch {
          if (-not $script:FilterErrorLogged) {
            $script:FilterErrorLogged = $true
            try { Write-GuiLog ("Grid filter error: " + $_.Exception.Message) } catch {}
          }
          return $false
        }
      }
    }

    $script:ResultsView.Refresh()
    Update-Counts
  } catch {
    Write-GuiLog ("Apply-GridFilter error: " + $_.Exception.Message)
  }
}


function Update-Counts {
  try {
    $total = 0
    if ($script:Results) { $total = @($script:Results).Count }

    $showing = 0
    if ($script:ResultsView) {
      foreach ($x in $script:ResultsView) { $showing++ }
    }

    $txtCounts.Text = ("Results: Showing {0} / Total {1}" -f $showing, $total)
  } catch {
    Write-GuiLog ("Update-Counts error: " + $_.Exception.Message)
  }
}


# -------------------- Helpers --------------------
function Set-Busy {
  param([bool]$Busy,[string]$Status)
  $btnScan.IsEnabled = -not $Busy
  $btnBrowse.IsEnabled = -not $Busy
  $pb.IsIndeterminate = $Busy
  $txtStatus.Text = $Status
}

function Get-KeepRuleValue {
  $item = $cmbKeepRule.SelectedItem
  if ($null -eq $item) { return "ShortestPath" }
  return [string]$item.Content
}

function Select-AllRemove {
  try {
    foreach ($r in $script:Results) {
      if (([string]$r.Action) -eq "REMOVE") { $r.Selected = $true }
    }
  } catch {
    Write-GuiLog "Select-AllRemove failed: $($_.Exception.Message)"
  }
}

function Clear-Selection {
  try {
    foreach ($r in $script:Results) { $r.Selected = $false }
  } catch {
    Write-GuiLog "Clear-Selection failed: $($_.Exception.Message)"
  }
}

function Refresh-Grid {
  try {
    if ($null -ne $script:ResultsView) { $script:ResultsView.Refresh() }
    if ($null -ne $dg) { $dg.Items.Refresh() }
  } catch {
    Write-GuiLog "Refresh-Grid failed: $($_.Exception.Message)"
  }
}

function Get-SelectedRows {
  return @($script:Results | Where-Object { $_.Selected -eq $true })
}

# -------------------- Events --------------------
$btnClearFilter.Add_Click({ $txtFilter.Text = ""; Apply-GridFilter })
$txtFilter.Add_TextChanged({ Apply-GridFilter })

$rbShowAll.Add_Checked({ Apply-GridFilter })
$rbShowRemove.Add_Checked({ Apply-GridFilter })
$rbShowKeep.Add_Checked({ Apply-GridFilter })
# If ConfirmedOnly turned OFF, re-apply grid filter
$chkConfirmedOnly.Add_Unchecked({ Apply-GridFilter })
$chkConfirmedOnly.Add_Checked({ Apply-GridFilter })

$btnSelectAllRemove.Add_Click({ try { Select-AllRemove; Refresh-Grid; Update-Counts } catch { Write-GuiLog "SelectAll button crashed: $($_.Exception.Message)" } })
$btnClearSel.Add_Click({ try { Clear-Selection; Refresh-Grid; Update-Counts } catch { Write-GuiLog "ClearSelection button crashed: $($_.Exception.Message)" } })

$btnOpenOut.Add_Click({
  try { Start-Process $OutDir | Out-Null } catch {}
})

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = "Select a folder to scan"
  $dlg.SelectedPath = if ([string]::IsNullOrWhiteSpace($txtPath.Text)) { $env:USERPROFILE } else { $txtPath.Text }
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $txtPath.Text = $dlg.SelectedPath
  }
})

# -------------------- Scan (background runspace) --------------------
function Invoke-Scan {
  param([string]$TargetPath)

  if (-not (Test-Path -LiteralPath $ScannerScript)) {
    [System.Windows.MessageBox]::Show("Scanner script not found:`n$ScannerScript","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) | Out-Null
    return
  }

  $minSize = 0
  [void][int]::TryParse(([string]$txtMinSize.Text), [ref]$minSize)

  $sampleN = 5
  [void][int]::TryParse(([string]$txtSampleRemove.Text), [ref]$sampleN)

  $keepRule = Get-KeepRuleValue

  $ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $detailCsv = Join-Path $OutDir ("duplicates_confirmed_keepremove_{0}.csv" -f $ts)
  $summaryCsv = Join-Path $OutDir ("duplicates_summary_{0}.csv" -f $ts)

  Write-GuiLog "Scanning: $TargetPath"
  Write-GuiLog "Audit CSV (detail): $detailCsv"
  Write-GuiLog "Audit CSV (summary): $summaryCsv"

  Set-Busy -Busy $true -Status "Scanning..."

  # Build runspace
  $script:psScan = [PowerShell]::Create()
  $rs = [runspacefactory]::CreateRunspace()
  $rs.ApartmentState = "STA"
  $rs.ThreadOptions = "ReuseThread"
  $rs.Open()
  $script:psScan.Runspace = $rs

  $null = $script:psScan.AddScript({
    param($ScannerScript,$TargetPath,$Recurse,$IncludeHidden,$MinSizeMB,$ConfirmContent,$ConfirmedOnly,$KeepRule,$DetailCsv,$SummaryCsv,$SampleRemove)
    . $ScannerScript
    & $ScannerScript -Path $TargetPath -Recurse:([bool]$Recurse) -IncludeHidden:([bool]$IncludeHidden) -MinSizeMB $MinSizeMB -ConfirmContent:([bool]$ConfirmContent) -ConfirmedOnly:([bool]$ConfirmedOnly) -KeepRule $KeepRule -ReportPath $DetailCsv -SummaryReportPath $SummaryCsv -SampleRemovePaths $SampleRemove
  }).AddArgument($ScannerScript).AddArgument($TargetPath).AddArgument($chkRecurse.IsChecked -eq $true).AddArgument($chkHidden.IsChecked -eq $true).AddArgument($minSize).AddArgument($chkConfirm.IsChecked -eq $true).AddArgument($chkConfirmedOnly.IsChecked -eq $true).AddArgument($keepRule).AddArgument($detailCsv).AddArgument($summaryCsv).AddArgument($sampleN)

  $script:asyncScan = $script:psScan.BeginInvoke()

  # Poll completion (UI thread safe via Dispatcher)
  $script:scanTimer = New-Object System.Windows.Threading.DispatcherTimer
  $script:scanTimer.Interval = [TimeSpan]::FromMilliseconds(250)
  $script:scanTimer.Add_Tick({
    if ($script:asyncScan -and $script:asyncScan.IsCompleted) {
      $script:scanTimer.Stop()
      $results = @()
      try {
        $results = $script:psScan.EndInvoke($script:asyncScan)
      } catch {
        Write-GuiLog "Scan failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Scan failed:`n$($_.Exception.Message)","Scan Failed",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) | Out-Null
      } finally {
        try { $script:psScan.Dispose() } catch {}
        try { $rs.Close(); $rs.Dispose() } catch {}
      }

      # Separate stats and rows
      $stats = $null
      $rows = @()
      foreach ($o in $results) {
        if ($null -eq $o) { continue }
        if (($o.PSObject.Properties.Name -contains "_Type") -and ([string]$o._Type -eq "ScanStats")) {
          $stats = $o
        } else {
          # ensure Selected property exists for UI
          if (-not ($o.PSObject.Properties.Name -contains "Selected")) {
            Add-Member -InputObject $o -MemberType NoteProperty -Name "Selected" -Value $false
          }
          $rows += $o
        }
      }

      if ($stats) {
        Write-GuiLog ("ScanStats: FilesEnumerated={0}, FilesScanned={1}, SizeGroupsGE2={2}, ConfirmedGroups={3}, ResultRows={4}, ErrorsSuppressed={5}, DurationSec={6}" -f `
          $stats.FilesEnumerated,$stats.FilesScanned,$stats.SizeGroupsGE2,$stats.ConfirmedGroups,$stats.ResultRows,$stats.ErrorsSuppressed,$stats.DurationSec)
      } else {
        Write-GuiLog "ScanStats: (missing) - scanner did not emit stats."
      }

      # Update UI collection
      $script:Results.Clear()
      foreach ($r in $rows) { $script:Results.Add($r) }
      Apply-GridFilter

      Set-Busy -Busy $false -Status "Idle"

      # Decide message
      if ($stats -and [int]$stats.FilesScanned -eq 0) {
        [System.Windows.MessageBox]::Show(
          "No files were enumerated/scanned from the target path.`n`nThis usually means Access Denied, the drive/folder isn't reachable, or traversal was blocked. Try running PowerShell as Administrator and/or enable IncludeHidden/System.",
          "No files scanned",
          [System.Windows.MessageBoxButton]::OK,
          [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
      }

  if (@($script:Results).Count -eq 0) {
        [System.Windows.MessageBox]::Show(
          "Scan completed successfully, but no duplicates matched the current filters.`n`nTip: If you expected results, try unchecking 'Confirmed only' and also untick 'Show REMOVE only', then scan again.",
          "No results",
          [System.Windows.MessageBoxButton]::OK,
          [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
      }
    }
  })
  $script:scanTimer.Start()
}

$btnScan.Add_Click({
  $p = [string]$txtPath.Text
  if ([string]::IsNullOrWhiteSpace($p)) {
    [System.Windows.MessageBox]::Show("Please select a folder or drive path to scan.","Missing path",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
    return
  }
  Invoke-Scan -TargetPath $p
})

# -------------------- Actions: Quarantine / Delete --------------------
$btnQuarantine.Add_Click({
  $selected = Get-SelectedRows
  $selected = @($selected)
  if ($selected.Count -eq 0) {
    [System.Windows.MessageBox]::Show("No rows selected.","Quarantine",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
    return
  }
  $remove = @($selected | Where-Object { ([string]$_.Action) -eq "REMOVE" })
  if ($remove.Count -eq 0) {
    [System.Windows.MessageBox]::Show("No selected rows are marked as REMOVE.","Quarantine",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
    return
  }

  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = "Select Quarantine folder"
  $dlg.SelectedPath = $OutDir
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
  $qdir = $dlg.SelectedPath
  if (-not (Test-Path -LiteralPath $qdir)) { New-Item -ItemType Directory -Path $qdir -Force | Out-Null }

  $whatIf = ($chkWhatIf.IsChecked -eq $true)
    $deletedRows = New-Object System.Collections.Generic.List[object]

  foreach ($r in $remove) {
    $src = [string]$r.FullPath
    if (-not (Test-Path -LiteralPath $src)) { Write-GuiLog "Quarantine skip (missing): $src"; continue }
    $dest = Join-Path $qdir ([System.IO.Path]::GetFileName($src))
    $i = 1
    while (Test-Path -LiteralPath $dest) {
      $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
      $ext = [System.IO.Path]::GetExtension($src)
      $dest = Join-Path $qdir ("{0}__dup{1}{2}" -f $base,$i,$ext)
      $i++
    }
    if ($whatIf) {
      Write-GuiLog "WhatIf: Move '$src' -> '$dest'"
    } else {
      try { Move-Item -LiteralPath $src -Destination $dest -Force; Write-GuiLog "Moved to quarantine: $src" } catch { Write-GuiLog "Quarantine move failed: $src :: $($_.Exception.Message)" }
    }
  }

  [System.Windows.MessageBox]::Show("Quarantine operation completed. (Check logs for details)","Quarantine",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
})


# --- Right-click: Copy FullPath ---
$miCopyFullPath = $window.FindName("miCopyFullPath")
if ($miCopyFullPath) {
  $miCopyFullPath.Add_Click({
    try {
      $sel = @($dg.SelectedItems)
      if (-not $sel -or $sel.Count -eq 0) { $sel = @($dg.SelectedItem) }
      $paths = @()
      foreach ($r in $sel) {
        $p = [string](Get-RowProp -Row $r -Name "FullPath")
        if (-not [string]::IsNullOrWhiteSpace($p)) { $paths += $p }
      }
      if ($paths.Count -gt 0) {
        [System.Windows.Clipboard]::SetText(($paths -join [Environment]::NewLine))
        Write-GuiLog ("Copied FullPath (" + $paths.Count + " line(s)) to clipboard.")
      }
    } catch {
      try { Write-GuiLog ("Copy FullPath failed: " + $_.Exception.Message) } catch {}
    }
  })
}

# --- Group drilldown (double-click): filter to GroupId + show ALL + highlight members ---
$dg.Add_MouseDoubleClick({
  try {
    $row = $dg.SelectedItem
    if ($null -eq $row) { return }
    $gid = [string](Get-RowProp -Row $row -Name "GroupId")
    if ([string]::IsNullOrWhiteSpace($gid)) { return }

    if ($rbShowAll) { $rbShowAll.IsChecked = $true }
    if ($txtFilter) { $txtFilter.Text = $gid }

    Apply-GridFilter

    $dg.SelectedItems.Clear()
    $first = $null
    foreach ($r in $script:ResultsView) {
      if (([string](Get-RowProp -Row $r -Name "GroupId")) -eq $gid) {
        if ($null -eq $first) { $first = $r }
        $null = $dg.SelectedItems.Add($r)
      }
    }
    if ($first) { $dg.ScrollIntoView($first) }
    Write-GuiLog ("Drilldown GroupId=" + $gid + " SelectedRows=" + $dg.SelectedItems.Count)
  } catch {
    try { Write-GuiLog ("Drilldown failed: " + $_.Exception.Message) } catch {}
  }
})

$btnDelete.Add_Click({
  try {
  $selected = Get-SelectedRows
    $selected = @($selected)
    if ($selected.Count -eq 0) {
      [System.Windows.MessageBox]::Show("No rows selected.","Delete",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
      return
    }
    $remove = @($selected | Where-Object { ([string]$_.Action) -eq "REMOVE" })
    if ($remove.Count -eq 0) {
      [System.Windows.MessageBox]::Show("No selected rows are marked as REMOVE.","Delete",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
      return
    }
  
    # NOTE: wrap the -f formatted string in parentheses so MessageBox::Show gets clean, comma-separated args
    $confirmMsg = ("Delete {0} file(s) marked as REMOVE?{1}{1}Tip: Use Quarantine first." -f $remove.Count, [Environment]::NewLine)
    $confirm = [System.Windows.MessageBox]::Show(
      $confirmMsg,
      "Confirm Delete",
      [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Warning
    )
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }
  
    $whatIf = ($chkWhatIf.IsChecked -eq $true)
    $useRecycle = ($chkRecycle.IsChecked -eq $true)

    # Track successfully deleted rows (required under StrictMode)
    $deletedRows = New-Object 'System.Collections.Generic.List[object]'

  
    foreach ($r in $remove) {
      $src = [string]$r.FullPath
      if (-not (Test-Path -LiteralPath $src)) { Write-GuiLog "Delete skip (missing): $src"; continue }
  
      if ($whatIf) {
        Write-GuiLog "WhatIf: Delete '$src' (RecycleBin=$useRecycle)"
        continue
      }
  
      try {
        if ($useRecycle) {
          [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
            $src,
            [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
            [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
          )
        } else {
          Remove-Item -LiteralPath $src -Force
        }
        Write-GuiLog "Deleted: $src"
        $null = $deletedRows.Add($r)
      } catch {
        Write-GuiLog "Delete failed: $src :: $($_.Exception.Message)"
      }
    }
  
    

    # Remove deleted rows from in-memory results so the grid updates immediately
    if (-not $whatIf -and $deletedRows.Count -gt 0) {
      $window.Dispatcher.Invoke([Action]{
        foreach ($dr in $deletedRows) { $null = $script:Results.Remove($dr) }
        Refresh-Grid
        Update-Counts
      })
    } else {
      Refresh-Grid
      Update-Counts
    }

[System.Windows.MessageBox]::Show("Delete operation completed. (Check logs for details)","Delete",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
  } catch {
    Write-GuiLog "Delete handler crashed: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show("Delete failed: $($_.Exception.Message)","Delete Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) | Out-Null
  }
})

# Default path
$txtPath.Text = (Get-Location).Path

# Show
$null = $window.ShowDialog()