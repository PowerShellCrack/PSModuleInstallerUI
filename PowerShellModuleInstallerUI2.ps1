<#
.SYNOPSIS
    This script is used to install PowerShell modules from the PowerShell Gallery.

.DESCRIPTION
    This script is used to install PowerShell modules from the PowerShell Gallery using a GUI interface.

.NOTES
    File Name      : PowerShellModuleInstallerUI.ps1
    Author         : Dick Tracy II
    Prerequisite   : PowerShell 5.1 or later

#>

[CmdletBinding()]
param(
    [string]$StoredDataPath,
    [string]$TagDetectionPath,
    [string]$LogFilePath, 
    [switch]$ForceNewModuleData,
    [switch]$ForceNewSolutionData,
    [switch]$SkipSolutionData,
    [switch]$SimulateInstall
)


##*=============================================
##*  FUNCTION
##*=============================================

##=============================================
## Sequence Window UI
##=============================================
Function Show-SequenceWindow {
    <#
    .SYNOPSIS
        Show the Sequence Window

    .EXAMPLE
        Show-SequenceWindow
    #>
    [CmdletBinding()]
    Param(
        $Config,
        $Message,
        [ValidateSet('Light','Dark')]
        [string]$Theme,
        $RunningApps,
        $TopPosition,
        $LeftPosition
    )

    [string]${CmdletName} = $MyInvocation.MyCommand
    Write-Verbose ("{0}: Sequencer started" -f ${CmdletName})

    # build a hash table with locale data to pass to runspace
    $syncHash = [hashtable]::Synchronized(@{})
    $Runspace =[runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $Runspace
    $syncHash.Config = $Config
    $syncHash.TopPosition = $TopPosition
    $syncHash.LeftPosition = $LeftPositio
    $syncHash.Message = $Message
    $syncHash.Theme = $Theme
    #build runspace
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open() | Out-Null
    $Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $Script:Pwshell = [PowerShell]::Create().AddScript({
        [string]$xaml = @"
<Window x:Class="PowerShellModuleSequencer.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellModuleSequencer"
        mc:Ignorable="d"
        WindowStartupLocation="CenterScreen"
        Title="MainWindow" Height="768" Width="1366"
        ResizeMode="NoResize" WindowStyle="None"
        BorderBrush="#0071B7"
        BorderThickness="0.5">
    <Window.Resources>

    </Window.Resources>
    <Grid>
        <TextBlock x:Name="txtMainTitle" HorizontalAlignment="Left" Text="PowerShell Module Installer" VerticalAlignment="Top" FontSize="32" Margin="10,10,0,0" TextAlignment="Left" FontFamily="Segoe UI Light" Foreground="Black" Width="635"/>
        <TextBox x:Name="txtVersion" HorizontalAlignment="Right" Height="25" VerticalAlignment="Top" Width="67" IsEnabled="False" Margin="0,10,10,0" BorderThickness="0" HorizontalContentAlignment="Right" />

        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Width="1366" Background="LightBlue" Opacity="80" Height="145">
            <TextBox x:Name="txtMessage" HorizontalAlignment="Center" HorizontalContentAlignment="Center" IsEnabled="False" BorderThickness="0" FontSize="26" FontWeight="Bold" Background="Transparent" Text="Please Wait..." TextWrapping="NoWrap" />
            <ProgressBar x:Name="ProgressBarMain" Width="630" Height="15" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="5" IsIndeterminate="True" />
            <ProgressBar x:Name="ProgressBarSub" Width="630" Height="15" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="5" />
            <TextBox x:Name="txtStatus" HorizontalAlignment="Center" Height="55" Width="630" IsEnabled="False" BorderThickness="0" Background="Transparent" FontSize="18" Text="Loading..." TextWrapping="Wrap"/>

        </StackPanel>
        <TextBox x:Name="txtPercentage" HorizontalAlignment="Right" Height="25" VerticalAlignment="Bottom" Width="67" IsEnabled="False" BorderThickness="0" TextWrapping="NoWrap" Margin="10"/>

        <Button x:Name="btnExit" Content="Exit" HorizontalAlignment="Left" VerticalAlignment="Top"  Width="80" Height="46" FontSize="16" Margin="10,700,0,0" IsEnabled="False"/>
    </Grid>
</Window>
"@

        #Load assembies to display UI
        [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')

        [xml]$xaml = $xaml -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
        $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
        $syncHash.window = [Windows.Markup.XamlReader]::Load($reader)

        #===========================================================================
        # Store Form Objects In PowerShell
        #===========================================================================
        $xaml.SelectNodes("//*[@Name]") | %{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

        # INNER  FUNCTIONS
        #Closes UI objects and exits (within runspace)
        Function Close-UISequenceWindow
        {
            if ($syncHash.hadCritError) { Write-Host -Message "Background thread had a critical error" -ForegroundColor red }
            #if runspace has not errored Dispose the UI
            if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
        }

        $syncHash.txtMainTitle.Text = $syncHash.Config.Title
        $syncHash.txtMessage.Text = $syncHash.Message
        $syncHash.txtVersion.Text = $syncHash.Config.Version
        $syncHash.btnExit.IsEnabled = $false
        $syncHash.btnExit.Visibility = 'Hidden'

        #Add smooth closing for Window
        $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
    	$syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-UISequenceWindow })
    	$syncHash.Window.Add_Closed({ $syncHash.isClosed = $True })

        #always force windows on bottom
        $syncHash.Window.Topmost = $True
        If($syncHash.TopPosition){
            $syncHash.Window.WindowStartupLocation = 'CenterOwner'
            $syncHash.Window.Top = $syncHash.TopPosition
        }
        If($syncHash.LeftPosition){
            $syncHash.Window.WindowStartupLocation = 'CenterOwner'
            $syncHash.Window.Left = $syncHash.LeftPosition
        }

        #Allow UI to be dragged around screen
        $syncHash.Window.Add_MouseLeftButtonDown( {
            $syncHash.Window.DragMove()
        })

        #action for exit button
        $syncHash.btnExit.Add_Click({
            Close-UISequenceWindow
        })

        $syncHash.Window.ShowDialog()
        #$Runspace.Close()
        #$Runspace.Dispose()
        $syncHash.Error = $Error
    }) # end scriptblock

    #collect data from runspace
    $Data = $syncHash

    #invoke scriptblock in runspace
    $Script:Pwshell.Runspace = $Runspace
    $AsyncHandle = $Script:Pwshell.BeginInvoke()

    #cleanup registered object
    Register-ObjectEvent -InputObject $syncHash.Runspace `
            -EventName 'AvailabilityChanged' `
            -Action {

                    if($Sender.RunspaceAvailability -eq "Available")
                    {
                        $Sender.Closeasync()
                        $Sender.Dispose()
                        # Speed up resource release by calling the garbage collector explicitly.
                        # Note that this will pause *all* threads briefly.
                        [GC]::Collect()
                    }

                } | Out-Null

    If($Data.Error){Write-Verbose ("{0}: Sequencer errored: {1}" -f ${CmdletName}, $Data.Error) }
    Else{Write-Verbose ("{0}: Sequencer closed" -f ${CmdletName})}
    Return $Data
}

#region FUNCTION: close window from sequencer
function Close-SequenceWindow
{
    Param (
        [Parameter(Mandatory=$true, Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        $Runspace
    )
    [string]${CmdletName} = $MyInvocation.MyCommand

    Write-Verbose ("{0}: Closing Sequencer window..." -f ${CmdletName})
    $Runspace.Window.Dispatcher.Invoke([action]{
        $Runspace.Window.close()
    },'Normal')
}
#endregion


#region Update progress bar for Sequencer
function Update-SequenceProgressBar
{
    [CmdletBinding(DefaultParameterSetName='percent')]
    Param (
        [Parameter(Mandatory=$true)]
        $Runspace,
        [parameter(Mandatory=$true)]
        $ProgressBar,
        [parameter(Mandatory=$true, ParameterSetName="percent")]
        [int]$PercentComplete,
        [parameter(Mandatory=$true, ParameterSetName="steps")]
        [int]$Step,
        [parameter(Mandatory=$true, ParameterSetName="steps")]
        [int]$MaxSteps,
        [parameter(Mandatory=$false, ParameterSetName="steps")]
        [int]$Timespan = 1,
        [parameter(Mandatory=$true, ParameterSetName="indeterminate")]
        [switch]$Indeterminate,
        [String]$Message = $Null,
        [ValidateSet("LightGreen","Yellow","Red","Blue","Green","Black")]
        [string]$Color = 'Green'
    )

    [string]${CmdletName} = $MyInvocation.MyCommand

    Try{
        #build field object from name
        If ($PSCmdlet.ParameterSetName -eq "steps")
        {
            #calculate percentage
            $PercentComplete = (($Step / $MaxSteps) * 100)
            #determine where increment will start
            If($Step -eq 1){
                $IncrementFrom = 1
            }Else{
                $IncrementFrom = ((($Step-1) / $MaxSteps) * 100)
            }
            $IncrementBy = ($PercentComplete-$IncrementFrom)/$Timespan
        }

        if($PSCmdlet.ParameterSetName -eq "indeterminate"){
           Write-Verbose ("{0}: Setting [{2}] to indeterminate with status: {1}" -f ${CmdletName},$Message,$ProgressBar)

            $Runspace.$ProgressBar.Dispatcher.Invoke([action]{
                $Runspace.$ProgressBar.Visibility = 'Visible'
                $Runspace.$ProgressBar.IsIndeterminate = $True
                $Runspace.$ProgressBar.Foreground = $Color
                $Runspace.txtPercentage.Visibility = 'Hidden'
                $Runspace.txtPercentage.Text = ' '

                $Runspace.txtStatus.Text = $Message
            }.GetNewClosure())

        }
        else
        {
            if(($PercentComplete -gt 0) -and ($PercentComplete -lt 100))
            {
                If($Timespan -gt 1){
                    $Runspace.Window.Dispatcher.Invoke([action]{
                        $t=1
                        #Determine the incement to go by based on timespan and difference
                        Do{
                            $IncrementTo = $IncrementFrom + ($IncrementBy * $t)

                            $Runspace.$ProgressBar.Visibility = 'Visible'
                            $Runspace.$ProgressBar.IsIndeterminate = $False
                            $Runspace.$ProgressBar.Value = $IncrementTo
                            $Runspace.$ProgressBar.Foreground = $Color

                            $Runspace.txtPercentage.Visibility = 'Visible'
                            $Runspace.txtPercentage.Text = ('' + $IncrementTo + '%')

                            $Runspace.txtStatus.Text = $Message

                            $t++
                            Start-Sleep 1

                        } Until ($IncrementTo -ge $PercentComplete -or $t -gt $Timespan)
                    }.GetNewClosure())
                }
                Else{
                   Write-Verbose ("{0}: Setting [{3}]  to {1}% with status: {2}" -f ${CmdletName},$PercentComplete,$Message,$ProgressBar)
                    $Runspace.Window.Dispatcher.Invoke([action]{

                        $Runspace.$ProgressBar.Visibility = 'Visible'
                        $Runspace.$ProgressBar.IsIndeterminate = $False
                        $Runspace.$ProgressBar.Value = $PercentComplete
                        $Runspace.$ProgressBar.Foreground = $Color

                        $Runspace.txtPercentage.Visibility = 'Visible'
                        $Runspace.txtPercentage.Text = ('' + $PercentComplete + '%')

                        $Runspace.txtStatus.Text = $Message
                    },'Normal')
                }
            }
            elseif($PercentComplete -eq 100)
            {
               Write-Verbose ("{0}: Setting [{2}] to complete with status: {1}" -f ${CmdletName},$Message,$ProgressBar)
                $Runspace.Window.Dispatcher.Invoke([action]{

                        $Runspace.$ProgressBar.Visibility = 'Visible'
                        $Runspace.$ProgressBar.IsIndeterminate = $False
                        $Runspace.$ProgressBar.Value = $PercentComplete
                        $Runspace.$ProgressBar.Foreground = $Color

                        $Runspace.txtPercentage.Visibility = 'Visible'
                        $Runspace.txtPercentage.Text = ('' + $PercentComplete + '%')

                        $Runspace.txtStatus.Text = $Message
                }.GetNewClosure())
            }
            else{
                Write-Verbose ("{0}: [{1}] is out of range" -f ${CmdletName},$ProgressBar)
            }
        }
    }Catch{}
}


##=============================================
## MAIN UI
##=============================================
Function Show-UIMainWindow
{
    <#
    .SYNOPSIS
        Shows the Powershell Module Install UI

    .PARAMETER XamlFile
    #>
    Param(
        [String]$TabContents,
        $TabElements,
        $ModuleData,
        $SolutionData,
        $InstalledModules,
        $Config,
        [String]$LogPath,
        [ValidateSet('Light','Dark')]
        [string]$Theme,
        $TopPosition,
        $LeftPosition,
        [switch]$DisableProcessCheck,
        [switch]$Wait
    )
    <#
    $XamlFile=$XAMLFilePath
    $StylePath=$StylePath
    $FunctionPath=$FunctionPath
    $Config=$ParamProps
    $Wait=$true
    #>
    #build runspace
    $syncHash = [hashtable]::Synchronized(@{})
    $PSRunSpace =[runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $PSRunSpace
    $syncHash.TabControlItems = $TabContents
    $syncHash.TabElements = $TabElements
    $syncHash.Config = $Config
    $syncHash.LogPath = $LogPath
    $syncHash.Theme = $Theme
    $syncHash.TopPosition = $TopPosition
    $syncHash.LeftPosition = $LeftPosition
    $syncHash.ModuleData = $ModuleData
    $syncHash.InstalledModules = $InstalledModules
    $syncHash.DisableProcessCheck = $DisableProcessCheck
    $syncHash.SolutionData = $SolutionData
    $syncHash.ModuleList = @()
    $syncHash.OutputData = @{}
    $syncHash.AdditionalDownloads = @()
    $PSRunSpace.ApartmentState = "STA"
    $PSRunSpace.ThreadOptions = "ReuseThread"
    $PSRunSpace.Open() | Out-Null
    $PSRunSpace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $PowerShellCommand = [PowerShell]::Create().AddScript({
    $MainXaml = @"
<Window x:Class="PowerShellModuleSelectorUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellModuleSelectorUI"
        mc:Ignorable="d"
        WindowStartupLocation="CenterScreen"
        Title="MainWindow" Height="768" Width="1366"
        ResizeMode="NoResize" WindowStyle="None"
        BorderBrush="#0071B7"
        BorderThickness="0.5">
    <Window.Resources>
        <ResourceDictionary>
            <Color x:Key="ControlLightColor">Gray</Color>
            <Color x:Key="ControlMediumColor">#FF7381F9</Color>
            <Color x:Key="ControlDarkColor">#FF211AA9</Color>
            <Color x:Key="BorderLightColor">#FFCCCCCC</Color>
            <Color x:Key="BorderMediumColor">#FF888888</Color>
            <Color x:Key="BorderDarkColor">#FF444444</Color>

            <Style x:Key="TabControlLeftSide" TargetType="{x:Type TabControl}">
                <Setter Property="OverridesDefaultStyle" Value="True" />
                <Setter Property="SnapsToDevicePixels" Value="True" />
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type TabControl}">
                            <Grid KeyboardNavigation.TabNavigation="Local">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" />
                                    <ColumnDefinition Width="*" />
                                </Grid.ColumnDefinitions>
                                <VisualStateManager.VisualStateGroups>
                                    <VisualStateGroup x:Name="CommonStates">
                                        <VisualState x:Name="Disabled">
                                            <Storyboard>
                                                <ColorAnimationUsingKeyFrames Storyboard.TargetName="Border" Storyboard.TargetProperty="(Border.BorderBrush).(SolidColorBrush.Color)">
                                                    <EasingColorKeyFrame KeyTime="0" Value="#FFAAAAAA" />
                                                </ColorAnimationUsingKeyFrames>
                                            </Storyboard>
                                        </VisualState>
                                    </VisualStateGroup>
                                </VisualStateManager.VisualStateGroups>
                                <ContentPresenter x:Name="PART_SelectedContentHost"
									  Grid.Column="1"
									  Margin="0"
									  ContentSource="SelectedContent" />
                                <StackPanel x:Name="HeaderPanel"
								Grid.Row="0"
								Margin="0,0,4,-1"
								Panel.ZIndex="1"
								Background="Transparent"
								IsItemsHost="True"
								KeyboardNavigation.TabIndex="1" />
                                <Border x:Name="Border"
							Grid.Row="1"
							BorderThickness="1"
							CornerRadius="2"
							KeyboardNavigation.DirectionalNavigation="Contained"
							KeyboardNavigation.TabIndex="2"
							KeyboardNavigation.TabNavigation="Local" />
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="WhiteTabItems" TargetType="{x:Type TabItem}">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type TabItem}">

                            <Grid x:Name="Root">
                                <VisualStateManager.VisualStateGroups>
                                    <VisualStateGroup x:Name="SelectionStates">
                                        <VisualState x:Name="Unselected" />
                                        <VisualState x:Name="Selected">
                                            <Storyboard>
                                                <ColorAnimationUsingKeyFrames Storyboard.TargetName="Border" Storyboard.TargetProperty="(Border.BorderBrush).(SolidColorBrush.Color)">
                                                    <EasingColorKeyFrame KeyTime="0" Value="#FFF" />
                                                </ColorAnimationUsingKeyFrames>
                                            </Storyboard>
                                        </VisualState>
                                    </VisualStateGroup>
                                    <VisualStateGroup x:Name="CommonStates">
                                        <VisualState x:Name="Normal" />
                                        <VisualState x:Name="MouseOver" />
                                        <VisualState x:Name="Disabled" />
                                    </VisualStateGroup>
                                </VisualStateManager.VisualStateGroups>
                                <Border x:Name="Border"
							Margin="0,0,0,0"
							BorderBrush="#FF1D3245"
							BorderThickness="0,0,0,0"/>
                                <TextBlock Margin="12,10,12,10" Text="{TemplateBinding Header}">
                                    <TextBlock.LayoutTransform>
                                        <TransformGroup>
                                            <ScaleTransform />
                                            <SkewTransform />
                                            <RotateTransform Angle="0" />
                                            <TranslateTransform />
                                        </TransformGroup>
                                    </TextBlock.LayoutTransform>
                                </TextBlock>
                            </Grid>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="Border" Property="BorderBrush" Value="Black" />
                                    <Setter TargetName="Border" Property="Background" Value="lightGray" />
                                    <Setter Property="Background" Value="Black" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Panel.ZIndex" Value="100" />
                                    <Setter Property="Foreground" Value="Black"/>
                                    <Setter TargetName="Border" Property="Background" Value="lightGray" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="2" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="False">
                                    <Setter Property="Panel.ZIndex" Value="100" />
                                    <Setter Property="Foreground" Value="#212121"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <!-- TabItem Style -->
            <Style x:Key="LineTabStyle" TargetType="{x:Type TabItem}" >
                <!--<Setter Property="Foreground" Value="#FFE6E6E6"/>-->
                <Setter Property="Template">
                    <Setter.Value>

                        <ControlTemplate TargetType="{x:Type TabItem}">
                            <Grid>
                                <Border
                                    Name="Border"
                                    Margin="15"
                                    Padding="5"
                                    CornerRadius="0">
                                    <ContentPresenter Name="ContentSite" VerticalAlignment="Center"
                                        HorizontalAlignment="Center" ContentSource="Header"
                                        RecognizesAccessKey="True" />
                                </Border>
                            </Grid>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="#313131" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#4c4c4c" />
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="False">
                                    <Setter Property="Foreground" Value="Gray" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#4c4c4c" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">

                                    <Setter Property="Foreground" Value="Black" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="#4B90CB" />
                                </Trigger>
                                <Trigger Property="IsSelected" Value="False">

                                    <Setter Property="Foreground" Value="Gray" />
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,3" />
                                    <Setter TargetName="Border" Property="BorderBrush" Value="White" />
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>

            </Style>

            <Style x:Key="DataGridContentCellCentering" TargetType="{x:Type DataGridCell}">
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type DataGridCell}">
                            <Grid Background="{TemplateBinding Background}">
                                <ContentPresenter VerticalAlignment="Center" />
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="CleanLineStyle" TargetType="{x:Type Button}">
                <Setter Property="Foreground" Value="Black" />
                <Setter Property="FontSize" Value="12" />
                <Setter Property="SnapsToDevicePixels" Value="True" />

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button" >

                            <Border Name="border"
                                BorderThickness="1"
                                Padding="4,2"
                                BorderBrush="#A19F9D"
                                CornerRadius="1"
                                Background="#F3F2F1">
                                <ContentPresenter HorizontalAlignment="Center"
                                                VerticalAlignment="Center"
                                                TextBlock.TextAlignment="Center"
                                                />
                            </Border>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#0078D4" />
                                    <Setter Property="Button.Foreground" Value="#FFE8EDF9" />
                                </Trigger>

                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#106EBE" />
                                    <Setter Property="Button.Foreground" Value="#FFE8EDF9" />
                                    <Setter Property="Background" Value="#106EBE" />
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#F3F2F1" />
                                    <Setter Property="Button.Foreground" Value="#A19F9D" />
                                </Trigger>
                                <Trigger Property="IsFocused" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#A19F9D" />
                                    <Setter Property="Button.Background" Value="#336891" />
                                </Trigger>

                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="CheckRadioFocusVisual">
                <Setter Property="Control.Template">
                    <Setter.Value>
                        <ControlTemplate>
                            <Rectangle Margin="14,0,0,0" SnapsToDevicePixels="true" Stroke="{DynamicResource {x:Static SystemColors.ControlTextBrushKey}}" StrokeThickness="1" StrokeDashArray="1 2"/>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
            <Style x:Key="SliderCheckBox" TargetType="{x:Type CheckBox}">
                <Setter Property="Foreground" Value="{DynamicResource {x:Static SystemColors.ControlTextBrushKey}}"/>
                <Setter Property="BorderThickness" Value="1"/>
                <Setter Property="Cursor" Value="Hand" />
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="{x:Type CheckBox}">
                            <ControlTemplate.Resources>
                                <Storyboard x:Key="StoryboardIsChecked">
                                    <DoubleAnimationUsingKeyFrames Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[3].(TranslateTransform.X)" Storyboard.TargetName="CheckFlag">
                                        <EasingDoubleKeyFrame KeyTime="0" Value="0"/>
                                        <EasingDoubleKeyFrame KeyTime="0:0:0.2" Value="14"/>
                                    </DoubleAnimationUsingKeyFrames>
                                </Storyboard>
                                <Storyboard x:Key="StoryboardIsCheckedOff">
                                    <DoubleAnimationUsingKeyFrames Storyboard.TargetProperty="(UIElement.RenderTransform).(TransformGroup.Children)[3].(TranslateTransform.X)" Storyboard.TargetName="CheckFlag">
                                        <EasingDoubleKeyFrame KeyTime="0" Value="14"/>
                                        <EasingDoubleKeyFrame KeyTime="0:0:0.2" Value="0"/>
                                    </DoubleAnimationUsingKeyFrames>
                                </Storyboard>
                            </ControlTemplate.Resources>
                            <BulletDecorator Background="Transparent" SnapsToDevicePixels="true">
                                <BulletDecorator.Bullet>
                                    <Border x:Name="ForegroundPanel" BorderThickness="1" Width="35" Height="20" CornerRadius="10">
                                        <Canvas>
                                            <Border Background="White" x:Name="CheckFlag" CornerRadius="10" VerticalAlignment="Center" BorderThickness="1" Width="19" Height="18" RenderTransformOrigin="0.5,0.5">
                                                <Border.RenderTransform>
                                                    <TransformGroup>
                                                        <ScaleTransform/>
                                                        <SkewTransform/>
                                                        <RotateTransform/>
                                                        <TranslateTransform/>
                                                    </TransformGroup>
                                                </Border.RenderTransform>
                                                <Border.Effect>
                                                    <DropShadowEffect ShadowDepth="1" Direction="180" />
                                                </Border.Effect>
                                            </Border>
                                        </Canvas>
                                    </Border>
                                </BulletDecorator.Bullet>
                                <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" Margin="{TemplateBinding Padding}" RecognizesAccessKey="True" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="Center"/>
                            </BulletDecorator>
                            <ControlTemplate.Triggers>
                                <Trigger Property="HasContent" Value="true">
                                    <Setter Property="FocusVisualStyle" Value="{StaticResource CheckRadioFocusVisual}"/>
                                    <Setter Property="Padding" Value="4,0,0,0"/>
                                </Trigger>
                                <Trigger Property="IsChecked" Value="True">
                                    <!--<Setter TargetName="ForegroundPanel" Property="Background" Value="{DynamicResource Accent}" />-->
                                    <Setter TargetName="ForegroundPanel" Property="Background" Value="Green" />
                                    <Trigger.EnterActions>
                                        <BeginStoryboard x:Name="BeginStoryboardCheckedTrue" Storyboard="{StaticResource StoryboardIsChecked}" />
                                        <RemoveStoryboard BeginStoryboardName="BeginStoryboardCheckedFalse" />
                                    </Trigger.EnterActions>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="false">
                                    <Setter Property="Foreground" Value="{DynamicResource {x:Static SystemColors.GrayTextBrushKey}}"/>
                                    <Setter TargetName="ForegroundPanel" Property="Background" Value="LightGray" />
                                </Trigger>
                                <Trigger Property="IsChecked" Value="False">
                                    <Setter TargetName="ForegroundPanel" Property="Background" Value="Gray" />
                                    <Trigger.EnterActions>
                                        <BeginStoryboard x:Name="BeginStoryboardCheckedFalse" Storyboard="{StaticResource StoryboardIsCheckedOff}" />
                                        <RemoveStoryboard BeginStoryboardName="BeginStoryboardCheckedTrue" />
                                    </Trigger.EnterActions>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="BlueRound" TargetType="{x:Type Button}">
                <Setter Property="Background" Value="#FF1D3245" />
                <Setter Property="Foreground" Value="#FFE8EDF9" />
                <Setter Property="FontSize" Value="15" />
                <Setter Property="SnapsToDevicePixels" Value="True" />

                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button" >

                            <Border Name="border"
                    BorderThickness="1"
                    Padding="4,2"
                    BorderBrush="#336891"
                    CornerRadius="6"
                    Background="#0078d7">
                                <ContentPresenter HorizontalAlignment="Center"
                                    VerticalAlignment="Center"
                                    TextBlock.TextAlignment="Center"
                                    />
                            </Border>

                            <ControlTemplate.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#FFE8EDF9" />
                                </Trigger>
                                <Trigger Property="IsPressed" Value="True">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#FF1D3245" />
                                    <Setter Property="Button.Foreground" Value="#FF1D3245" />
                                    <Setter Property="Effect">
                                        <Setter.Value>
                                            <DropShadowEffect ShadowDepth="0" Color="#FF1D3245" Opacity="1" BlurRadius="10"/>
                                        </Setter.Value>
                                    </Setter>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#336891" />
                                    <Setter Property="Button.Foreground" Value="#336891" />
                                </Trigger>
                                <Trigger Property="IsFocused" Value="False">
                                    <Setter TargetName="border" Property="BorderBrush" Value="#336891" />
                                    <Setter Property="Button.Background" Value="#336891" />
                                </Trigger>

                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <TextBlock x:Name="txtMainTitle" HorizontalAlignment="Left" Text="PowerShell Module Installer" VerticalAlignment="Top" FontSize="32" Margin="10,10,0,0" TextAlignment="Left" FontFamily="Segoe UI Light" Foreground="Black" Width="635"/>

        <TextBox x:Name="txtVersion" HorizontalAlignment="Right" Height="25" VerticalAlignment="Top" Width="67" IsEnabled="False" Margin="0,10,10,0" BorderThickness="0" HorizontalContentAlignment="Right" />

        <TabControl x:Name="tabControlMainAllModules" Margin="0,55,0,65" Style="{StaticResource TabControlLeftSide}" BorderThickness="0" >
            $($syncHash.TabControlItems)
        </TabControl>

        <StackPanel Margin="0,60,10,92" HorizontalAlignment="Right">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110*" />
                    <ColumnDefinition Width="135*"/>
                </Grid.ColumnDefinitions>
                <Label Content="Selected Modules:" Grid.Column="0" HorizontalAlignment="Left" FontSize="12" Foreground="Gray" Width="107" />
                <TextBox x:Name="txtTotalSelectedCount" Grid.Column="1" Text="0" HorizontalAlignment="Left" HorizontalContentAlignment="Left" Width="33" IsEnabled="False" BorderThickness="0" VerticalAlignment="Center" Height="15"/>
            </Grid>
            <ListBox x:Name="lbSelectedModules" HorizontalAlignment="Left" FontSize="12" Width="248" Height="565"/>
        </StackPanel>
        <CheckBox x:Name="chkModuleAutoUpdate"  Content="Auto update modules" HorizontalAlignment="Left" Margin="137,0,0,31" VerticalAlignment="Bottom" Width="183" Style="{DynamicResource SliderCheckBox}" />
        <CheckBox x:Name="chkModuleRemoveDuplicates"  Content="Remove duplicate modules" HorizontalAlignment="Left" Margin="303,0,0,31" VerticalAlignment="Bottom" Width="185" Style="{DynamicResource SliderCheckBox}" />
        <CheckBox x:Name="chkModuleRemoveAll"  Content="Remove all modules (except selected)" HorizontalAlignment="Left" Margin="495,716,0,31" VerticalAlignment="Bottom" Width="253" Style="{DynamicResource SliderCheckBox}" />
        <CheckBox x:Name="chkModuleRepairSelected" Content="Repair selected modules" HorizontalAlignment="Left" Margin="743,716,0,31" VerticalAlignment="Bottom" Width="186" Style="{DynamicResource SliderCheckBox}" />
        <CheckBox x:Name="chkModuleInstallUserContext" Content="Install under user context" HorizontalAlignment="Left" Margin="924,716,0,31" VerticalAlignment="Bottom" Width="186" Style="{DynamicResource SliderCheckBox}" />
        <CheckBox x:Name="chkModuleInstallForPS7" Content="Install for Powershell 7*" HorizontalAlignment="Left" Margin="1108,660,0,0" VerticalAlignment="Top" Width="167" Style="{DynamicResource SliderCheckBox}" />
        <Button x:Name="btnInstall" Style="{DynamicResource BlueRound}" Content="Install Selected Modules" HorizontalAlignment="Right" VerticalAlignment="Bottom"  Width="246" Height="66" FontSize="16" Margin="0,0,10,10" />
        <Button x:Name="btnExit" Content="Exit" HorizontalAlignment="Left" VerticalAlignment="Bottom"  Width="80" Height="46" FontSize="16" Margin="10,0,0,10" IsEnabled="False"/>
    </Grid>
</Window>
"@
    #$Code{
        #Load assembles to display UI
        [System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework') | out-null
        [System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')      | out-null
        #convert XAML to XML just to grab info using xml dot sourcing (Not used to process form)
        [string]$XML = $MainXaml -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window' -replace 'Click=".*','/>'
        #convert XAML to XML
        [xml]$xaml = $XML
        $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
        # Store Form Objects In hashtable
        #===========================================================================
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}
        #===========================================================================
        #build log name
        [string]$LogFilePath = $syncHash.LogPath
        #=================================
        #Functions inside runspace
        #=================================

        Function Test-IsISE {
            # try...catch accounts for:
            # Set-StrictMode -Version latest
            try {
                return ($null -ne $psISE);
            }
            catch {
                return $false;
            }
        }

        Function Test-PwshInstalled{
            param([switch]$Passthru)
            #$pwshPath = "C:\Program Files\PowerShell\7\pwsh.exe"
            #$pwshUserPath = "$env:LOCALAPPDATA\Microsoft\PowerShell\7\pwsh.exe"
            $PwshInstalled = $false
            $PwshPath = Get-Command pwsh -ErrorAction SilentlyContinue
            If($PwshPath){$PwshInstalled = $true}

            If($Passthru){Return $PwshPath}
            Else{Return $PwshInstalled}
        }

        function Show-ConfirmationPopup {
            param (
                [string]$Message
            )

            #build runspace
            $syncHash = [hashtable]::Synchronized(@{})
            $ASPRunSpace =[runspacefactory]::CreateRunspace()
            $syncHash.Runspace = $ASPRunSpace
            $syncHash.Message = $Message
            $syncHash.Decision = $null
            $ASPRunSpace.ApartmentState = "STA"
            $ASPRunSpace.ThreadOptions = "ReuseThread"
            $ASPRunSpace.Open() | Out-Null
            $ASPRunSpace.SessionStateProxy.SetVariable("syncHash",$syncHash)
            $PowerShellCommand = [PowerShell]::Create().AddScript({

            [string]$xaml = @"
                    <Window
                        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                        xmlns:local="clr-namespace:NotificationPopup"
                        Title="NotificationPopup"
                        WindowStyle="None"
                        WindowStartupLocation="CenterScreen"
                        Height="200" Width="400"
                        ResizeMode="NoResize"
                        ShowInTaskbar="False">
                        <Window.Resources>
                            <Style TargetType="{x:Type Button}">
                                <!-- This style is used for buttons, to remove the WPF default 'animated' mouse over effect -->
                                <Setter Property="OverridesDefaultStyle" Value="True"/>
                                <Setter Property="Foreground" Value="#FFEAEAEA"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button" >

                                            <Border Name="border"
                                                        BorderThickness="1"
                                                        Padding="4,2"
                                                        BorderBrush="#FFEAEAEA"
                                                        CornerRadius="2"
                                                        Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center"
                                                                    VerticalAlignment="Center"
                                                                    TextBlock.FontSize="10px"
                                                                    TextBlock.TextAlignment="Center"
                                                                    />
                                            </Border>
                                            <ControlTemplate.Triggers>
                                                <Trigger Property="IsMouseOver" Value="True">
                                                    <Setter TargetName="border" Property="BorderBrush" Value="#FF919191" />
                                                </Trigger>
                                            </ControlTemplate.Triggers>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </Window.Resources>
                        <Grid Background="#313130">
                            <StackPanel Margin="10">
                                <TextBox x:Name="txtMsg" HorizontalAlignment="Center" VerticalAlignment="Center" Width="332" Foreground="Red" FontSize="24" IsEnabled="False" Background="Transparent" BorderThickness="0" TextWrapping="Wrap"/>
                                <TextBox Text="Do you want to proceed?" HorizontalContentAlignment="Center" Width="332" Foreground="White" FontSize="18" IsEnabled="False" Background="Transparent" BorderThickness="0" TextWrapping="Wrap"/>
                            </StackPanel>
                            <Button x:Name="btnSubmit" Content="Yes" HorizontalAlignment="Left" Margin="286,140,0,0" VerticalAlignment="Top" Width="95" Height="50" Background="Black" Foreground="White"/>
                            <Button x:Name="btnCancel" Content="No" HorizontalAlignment="Left" Margin="32,140,0,0" VerticalAlignment="Top" Width="95" Height="50" Background="Black" Foreground="White"/>

                        </Grid>
                    </Window>
"@
                Add-Type -AssemblyName PresentationFramework
                # Load XAML
                [xml]$xaml = $xaml -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
                $reader=(New-Object System.Xml.XmlNodeReader $xaml)
                $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
                #===========================================================================
                # Store Form Objects In PowerShell
                #===========================================================================
                $xaml.SelectNodes("//*[@Name]") | %{ $syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}

                Function Close-ConFirmationPopup {
                    if ($syncHash.hadCritError) { Write-Host -Message "Background thread had a critical error" -ForegroundColor red }
                    #if runspace has not errored Dispose the UI
                    if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
                }

                $syncHash.txtMsg.Text = $syncHash.Message

                $syncHash.btnSubmit.Add_Click({
                    $syncHash.Decision = $true
                    Close-ConFirmationPopup
                })

                $syncHash.btnCancel.Add_Click({
                    $syncHash.Decision = $false
                    Close-ConFirmationPopup
                })

                #Allow UI to be dragged around screen
                $syncHash.Window.Add_MouseLeftButtonDown( {
                    $syncHash.Window.DragMove()
                })

                #Add smooth closing for Window
                $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
                $syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-ConFirmationPopup })
                $syncHash.Window.Add_Closed({ $syncHash.isClosed = $True })

                #make sure this display on top of every window
                $syncHash.Window.Topmost = $true

                $syncHash.window.ShowDialog()
                $syncHash.Error = $Error
            }) # end scriptblock

            #collect data from runspace
            $Data = $syncHash

            #invoke scriptblock in runspace
            $PowerShellCommand.Runspace = $ASPRunSpace
            $AsyncHandle = $PowerShellCommand.BeginInvoke()

            #wait until runspace is completed before ending
            do {
                Start-sleep -m 100 }
            while (!$AsyncHandle.IsCompleted)
            #end invoked process
            $null = $PowerShellCommand.EndInvoke($AsyncHandle)

            #cleanup registered object
            Register-ObjectEvent -InputObject $syncHash.Runspace `
                    -EventName 'AvailabilityChanged' `
                    -Action {

                            if($Sender.RunspaceAvailability -eq "Available")
                            {
                                $Sender.Closeasync()
                                $Sender.Dispose()
                                # Speed up resource release by calling the garbage collector explicitly.
                                # Note that this will pause *all* threads briefly.
                                [GC]::Collect()
                            }

                        } | Out-Null

            return $Data

        }#end runspace function

        Function Get-RunningPowershellApps {
            param([switch]$Passthru)

            $PoshObject = @()
            # Define the list of processes to check
            $processNames = @("Code", "powershell_ise", "powershell", "pwsh")

            # Get the current process ID
            $currentProcessId = $PID

            # Get the list of running processes that match the defined process names, excluding the current process
            $runningProcesses = Get-Process | Where-Object { $_.Name -in $processNames -and $_.Id -ne $currentProcessId }

            # If Passthru is specified, return the running processes object
            $runningProcesses | ForEach-Object {
                $PoshObject += [PSCustomObject]@{
                    Id = $_.Id
                    Name = $_.Name
                    Product = $_.MainModule.FileVersionInfo.ProductName
                    ProductVersion = $_.MainModule.FileVersionInfo.ProductVersion
                    Path = $_.Path
                    SessionId = $_.SessionId
                }
            }

            if ($Passthru) {
                return $runningProcesses
            }Else{
                return $PoshObject
            }
        }

        function Confirm-RunningApps {
            # Get all running PowerShell instances except the current one
            $runningPS = Get-RunningPowershellApps -Passthru

            if ($runningPS.count -gt 0) {
                # Prompt user with XAML Popup
                $msg = ("There are {0} applications currently running that will be closed." -f $runningPS.count)
                $continue = Show-ConfirmationPopup -Message $msg

                if ($continue.Decision) {
                    # Kill detected PowerShell processes
                    $runningPS | Stop-Process -Force
                    return $true
                } else {
                    return $false
                }
            }
            return $continue.Decision
        }

        Function Write-UILogEntry{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
                [ValidateNotNullOrEmpty()]
                [string]$Message,

                [Parameter(Mandatory=$false,Position=2)]
                [string]$Source,

                [parameter(Mandatory=$false)]
                [ValidateSet(0,1,2,3,4,5)]
                [int16]$Severity = 1,

                [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
                [ValidateNotNullOrEmpty()]
                [string]$OutputLogFile = $LogFilePath
            )
            ## Get the name of this function
            #[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
            if (-not $PSBoundParameters.ContainsKey('Verbose')) {
                $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
            }

            if (-not $PSBoundParameters.ContainsKey('Debug')) {
                $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference')
            }
            #get BIAS time
            [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
            [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
            [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
            [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias

            #  Get the file name of the source script
            If($Source){
                $ScriptSource = $Source
            }
            Else{
                Try {
                    If ($script:MyInvocation.Value.ScriptName) {
                        [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
                    }
                    Else {
                        [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
                    }
                }
                Catch {
                    $ScriptSource = ''
                }
            }

            #if the severity and preference level not set to silentlycontinue, then log the message
            $LogMsg = $true
            If( $Severity -eq 4 ){$Message='VERBOSE: ' + $Message;If(!$VerboseEnabled){$LogMsg = $false} }
            If( $Severity -eq 5 ){$Message='DEBUG: ' + $Message;If(!$DebugEnabled){$LogMsg = $false} }
            #If( ($Severity -eq 4) -and ($VerbosePreference -eq 'SilentlyContinue') ){$LogMsg = $false$Message='VERBOSE: ' + $Message}
            #If( ($Severity -eq 5) -and ($DebugPreference -eq 'SilentlyContinue') ){$LogMsg = $false;$Message='DEBUG: ' + $Message}

            #generate CMTrace log format
            $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$Severity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"

            # Add value to log file
            If($LogMsg)
            {
                try {
                    Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
                }
                catch {
                    Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
                }
            }
        }

        #Closes UI objects and exits (within runspace)
        Function Close-UIMainWindow
        {
            if ($syncHash.hadCritError) { Write-UILogEntry -Message ("Critical error occurred, closing UI: {0}" -f $syncHash.Error) -Source 'Close-UIMainWindow' -Severity 3 }
            #if runspace has not errored Dispose the UI
            if (!($syncHash.isClosing)) { $syncHash.Window.Close() | Out-Null }
        }

        Function Update-UIListView{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                $ListView,
                [Parameter(Mandatory = $true)]
                $Data
            )
            Try{
                $syncHash.Window.Dispatcher.Invoke([action]{
                    $ListView.ItemsSource = $null
                    $ListView.ItemsSource = $Data
                }.GetNewClosure())
                Write-UILogEntry -Message ("Added [{1}] data to list element [{0}]" -f $ListView.Name,$Data.count) -Source 'Update-UIListView' -Severity 0
            }Catch{
                Write-UILogEntry -Message ("Failed to add [{1}] data to list element [{0}]: {2}" -f $ListView.Name,$Data.count,$_.Exception.Message) -Source 'Update-UIListView' -Severity 3
            }
        }

        # Define a function to handle sorting
        function Sort-UIListViewColumn {
            param(
                [System.Windows.Controls.ListView]$ListView,
                [System.Windows.Controls.GridViewColumnHeader]$ColumnHeader
            )

            # Get the binding property name from the column
            $binding = $ColumnHeader.Column.DisplayMemberBinding.Path.Path
            if (-not $binding) { return }

            # Get the current view and sort direction
            $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($ListView.ItemsSource)

            if ($view.SortDescriptions.Count -gt 0 -and $view.SortDescriptions[0].PropertyName -eq $binding) {
                # Toggle sort direction
                if ($view.SortDescriptions[0].Direction -eq [System.ComponentModel.ListSortDirection]::Ascending) {
                    $direction = [System.ComponentModel.ListSortDirection]::Descending
                } else {
                    $direction = [System.ComponentModel.ListSortDirection]::Ascending
                }
                # Clear previous sorting
                $view.SortDescriptions.Clear()
            } else {
                $direction = [System.ComponentModel.ListSortDirection]::Ascending
            }

            # Apply new sorting
            $view.SortDescriptions.Clear()
            $view.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription($binding, $direction)))
        }

        Function Update-UIText{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                $TextBox,
                [Parameter(Mandatory = $true)]
                [string]$Text
            )
            Try{
                $syncHash.Window.Dispatcher.Invoke("Normal",[action]{
                    $TextBox.Text = $Text
                }.GetNewClosure())
                Write-UILogEntry -Message ("Set UI element [{0}] to value [{1}]" -f $TextBox.Name,$Text) -Source 'Update-UIText' -Severity 0
            }Catch{
                Write-UILogEntry -Message ("Failed to set UI element [{0}] to value [{1}]: {2}" -f $TextBox.Name,$Text,$_.Exception.Message) -Source 'Update-UIText' -Severity 3
            }
        }

        Function Update-UICheckBox{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                $CheckBox,
                [Parameter(Mandatory = $true)]
                [bool]$IsChecked
            )
            Try{
                $syncHash.Window.Dispatcher.Invoke("Normal",[action]{
                    $CheckBox.IsChecked = $IsChecked
                }.GetNewClosure())
                Write-UILogEntry -Message ("Set UI element [{0}] to value [{1}]" -f $CheckBox.Name,$IsChecked) -Source 'Update-UICheckBox' -Severity 0
            }Catch{
                Write-UILogEntry -Message ("Failed to set UI element [{0}] to value [{1}]: {2}" -f $CheckBox.Name,$IsChecked,$_.Exception.Message) -Source 'Update-UICheckBox' -Severity 3
            }
        }

        function Update-UICombobox{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                $ComboBox,
                [Parameter(Mandatory = $true)]
                [string[]]$List,
                [Parameter(Mandatory = $false)]
                [string]$SelectedItem
            )
            Try{
                $syncHash.Window.Dispatcher.Invoke("Normal",[action]{
                    $ComboBox.ItemsSource = @($List)
                    If($SelectedItem){
                        $ComboBox.SelectedItem = $SelectedItem
                    }
                }.GetNewClosure())
                Write-UILogEntry -Message ("Added [{1}] items to UI element property [{0}]" -f $ComboBox.Name,$List.Count) -Source 'Update-UICombobox' -Severity 0
            }Catch{
                Write-UILogEntry -Message ("Failed to add [{1}] items to UI element property [{0}]: {2}" -f $ComboBox.Name,$List.Count,$_.Exception.Message) -Source 'Update-UICombobox' -Severity 3
            }
        }

        Function Update-UIElementProperty{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory=$true)]
                $Element,

                [Parameter(Mandatory=$true)]
                [ValidateSet('Visibility','Text','Content','Foreground','Background','IsReadOnly','IsChecked','IsEnabled','Fill','BorderThickness','BorderBrush')]
                [String]$Property,

                [Parameter(,Mandatory=$true)]
                [String]$Value
            )
            Try{
                $syncHash.Window.Dispatcher.invoke([action]{
                    $Element.$Property=$Value
                }.GetNewClosure())
                Write-UILogEntry -Message ("Set UI element [{0}] property [{1}] to value [{2}]" -f $Element.Name,$Property,$Value) -Source 'Update-UIElementProperty' -Severity 0
            }Catch{
                Write-UILogEntry -Message ("Failed to set UI element [{0}] property [{1}] to value [{2}]: {3}" -f $Element.Name,$Property,$Value,$_.Exception.Message) -Source 'Update-UIElementProperty' -Severity 3
            }
        }

        Function Format-UIListViewData{
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
                $NewList,
                [Parameter(Mandatory = $true)]
                $ExistingList
            )
            Begin{
                $FormattedList = @()
            }
            Process{
                #TEST $module = $ModuleItem.ModuleList[0]
                Foreach ($NewItem in $NewList)
                {
                    If($null -eq $NewItem.Name){
                        Write-UILogEntry -Message ("Item name is null: {0}" -f $NewItem) -Source 'Format-UIListViewData' -Severity 2
                        Continue
                    }
                    Try{
                        Write-UILogEntry -Message ("Adding a formatted listview item for [{0}]" -f $NewItem.Name) -Source 'Format-UIListViewData' -Severity 0
                        $VersionList = @()

                        $ExistingItem = $ExistingList | Where-Object {$_.Name -eq $NewItem.Name}
                        If(-Not $ExistingItem){
                            $ExistingItem = @()
                            $Message = "Not Installed"
                            $Selected = $false
                        }ElseIf($ExistingItem.Count -gt 1){
                            $Message = "Multiple Installed"
                            $VersionList += $ExistingItem | Foreach { $_.Version.ToString() }
                            $Selected = $true
                        }ElseIf($ExistingItem.Version -ne $NewItem.Version){
                            $Message = "Update Available"
                            $VersionList += $ExistingItem.Version.ToString()
                            $Selected = $true
                        }Else{
                            $Message = "Installed"
                            $VersionList += $ExistingItem.Version.ToString()
                            $Selected = $true
                        }

                        $FormattedList += [PSCustomObject]@{
                            IsSelected       = $Selected
                            Name             = $NewItem.Name
                            cVersion         = $VersionList | Select -Last 1
                            lVersion         = $NewItem.Version.ToString()
                            Owner            = $NewItem.Author
                            Count            = $ExistingItem.count
                            Status           = $Message
                            Version          = $VersionList  # ComboBox options
                            SelectedVersion  = $VersionList[-1] # Default dropdown selection
                        }

                    }catch{
                        Write-UILogEntry -Message ("Failed to format listview item [{0}]: {1}" -f $NewItem.Name,$_.Exception.Message) -Source 'Format-UIListViewData' -Severity 3
                    }
                }

            }End{

                return $FormattedList
            }
        }

        If(Test-IsISE){
            $Windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
            $asyncWindow = Add-Type -MemberDefinition $Windowcode -name Win32ShowWindowAsync -namespace Win32Functions
            $null = $asyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
        }
        #*=============================================
        #* BUILD UI EVENTS AND ACTIONS
        #*=============================================
        [System.Windows.RoutedEventHandler]$Script:CheckBoxChecked = {
            #do something first
        }
        #Get all check boxes and add event handler
        $syncHash.Content.Children | Where {
            $_ -is [System.Windows.Controls.CheckBox]
        } | ForEach {
            $_.AddHandler([System.Windows.Controls.CheckBox]::CheckedEvent, $CheckBoxChecked)
        }

        #=================================
        # MAIN UI EVENTS
        #=================================
        #set do action to false until Install button is clicked
        $syncHash.OutputData.DoAction = $false
        $syncHash.txtMainTitle.Text = $syncHash.Config.Title
        $syncHash.txtVersion.Text = $syncHash.Config.Version
        #set default values
        $syncHash.chkModuleAutoUpdate.isChecked = [Boolean]::Parse($syncHash.Config.DefaultSettings.AutoUpdate)
        $syncHash.chkModuleRemoveDuplicates.isChecked = [Boolean]::Parse($syncHash.Config.DefaultSettings.RemoveDuplicates)
        $syncHash.chkModuleRemoveAll.isChecked = [Boolean]::Parse($syncHash.Config.DefaultSettings.RemoveAllModulesFirst)

        If([Boolean]::Parse($syncHash.Config.DefaultSettings.AllowUserContextInstall)){
            $syncHash.chkModuleInstallUserContext.Visibility = 'Visible'
        }
        Else{
            $syncHash.chkModuleInstallUserContext.Visibility = 'Hidden'
        }

        If($syncHash.Config.DefaultSettings.InstallMode -eq 'CurrentUser'){
            $syncHash.chkModuleInstallUserContext.IsChecked = $true
        }Else{
            $syncHash.chkModuleInstallUserContext.IsChecked = $false
        }

        #$syncHash.btnInstall.IsEnabled = $false
        $syncHash.btnExit.IsEnabled = $true

        $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        #default to install under user context no matter if running as SYSTEM or admin
        $syncHash.chkModuleInstallUserContext.IsChecked = $true
        if ($CurrentUser -eq "NT AUTHORITY\SYSTEM") {
            Write-UILogEntry -Message "UI is running as SYSTEM" -Source 'Show-UIMainWindow' -Severity 0
            $syncHash.chkModuleInstallUserContext.IsEnabled = $false
        } else {
            Write-UILogEntry -Message "UI is running as $CurrentUser" -Source 'Show-UIMainWindow' -Severity 0
            $syncHash.chkModuleInstallUserContext.IsEnabled = $true
        }

        If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
        {
            Write-UILogEntry -Message "UI is not running as Administrator" -Source 'Show-UIMainWindow' -Severity 0
            $syncHash.chkModuleInstallUserContext.IsEnabled = $false
            $syncHash.chkModuleInstallUserContext.IsChecked = $True
        } else {
            Write-UILogEntry -Message "UI is running as Administrator" -Source 'Show-UIMainWindow' -Severity 0
            $syncHash.chkModuleInstallUserContext.IsEnabled = $true
        }

        If(Test-PwshInstalled){
            $PwshVersion = (Test-PwshInstalled -Passthru).Version.ToString()
            $syncHash.chkModuleInstallForPS7.Visibility = 'Visible'
            Write-UILogEntry -Message ("PowerShell {0} is installed" -f $PwshVersion) -Source 'Show-UIMainWindow' -Severity 0
        }Else{
            $syncHash.chkModuleInstallForPS7.Visibility = 'Hidden'
            Write-UILogEntry -Message "PowerShell 7* is not installed" -Source 'Show-UIMainWindow' -Severity 0
        }

        #update all comboboxes
        $syncHash.Keys | Where-Object { $_ -match "^cmb[A-Z][a-zA-Z]+ModuleStatusType$" } | ForEach-Object {
            $comboBoxName = $_
            Update-UICombobox -ComboBox $syncHash[$comboBoxName] -List @("All","Installed","Update Available","Not Installed") -SelectedItem "All"
        }

        # Main ListBox for displaying selected items across all tabs
        $mainListBox = $syncHash.lbSelectedModules

        $updateUI = {
            #what to do to run on a timer
        }
        <#
        # use tab control data to populate moduledata list element

        TabElements:
        Name                 Type     Elements
        ----                 ----     --------
        Azure                Module   {tabItemAzure, txtAzureModuleNameSearch, cmbAzureModuleStatusType, btnAzureSelectAll...}

        ModuleData:
        GroupName            ModuleList
        ---------            ----------
        Azure                {@{Version=4.0.2; Name=Az.Accounts; Repository=PSGallery; Description=Microsoft Azure PowerShell - Accounts credential management cmdlets for Azure Resource Manager in Windows PowerShell and PowerShell Core....
        #>

        $syncHash.TabElements | ForEach-Object{

            Write-UILogEntry -Message ("Processing Group [{0}]" -f $_.Name) -Source 'EventHandler' -Severity 0
            If($_.Type -eq 'Module'){
                $ListData = ($syncHash.ModuleData | Where GroupName -eq $_.Name).ModuleList | Format-UIListViewData -ExistingList $syncHash.InstalledModules
            }Else{
                $ListData = ($syncHash.SolutionData | Where GroupName -eq $_.Name).ModuleList | Format-UIListViewData -ExistingList $syncHash.InstalledModules
            }

            $ListData = @( $ListData )  # Ensure it's a collection
            $syncHash.ModuleList += $ListData

            #get lst element
            $listview = $_.Elements | Where-Object { $_ -match "^lst(.+)List$" }
            $moduleCount = $_.Elements | Where-Object { $_ -match "^txt(.+)ModuleCount$" }
            $selectAll = $_.Elements | Where-Object { $_ -match "^btn(.+)SelectAll$" }
            $selectNone = $_.Elements | Where-Object { $_ -match "^btn(.+)SelectNone$" }
            $selectupdatesbtn = $_.Elements | Where-Object { $_ -match "^btn(.+)SelectUpdates$" }
            $selectedCount = $_.Elements | Where-Object { $_ -match "^txt(.+)SelectedCount$" }

            Write-UILogEntry -Message ("Adding [{0}] data to list element [{1}]" -f $ListData.count,$listview) -Source 'EventHandler' -Severity 0

            Try{
                Update-UIListView -ListView $syncHash[$listview] -Data $ListData
                Update-UIText -TextBox $syncHash[$moduleCount] -Text $ListData.Count
            }catch{
                Write-UILogEntry -Message ("Failed to add [{0}] data to list element [{1}]: {2}" -f $ListData.count,$listview,$_.Exception.Message) -Source 'EventHandler' -Severity 3
            }

            #if listview has 'Updates Available' in column Status, then show select updates button
            If($ListData | Where-Object { $_.Status -eq 'Update Available' }){
                $syncHash[$selectupdatesbtn].Visibility = 'Visible'
            }Else{
                $syncHash[$selectupdatesbtn].Visibility = 'Hidden'
            }

            #make search text box work
            If($_.Type -eq 'Module'){
                $searchInput = $_.Elements | Where-Object { $_ -match "^txt(.+)NameSearch$" }
                $statusType = $_.Elements | Where-Object { $_ -match "^cmb(.+)StatusType$" }

                Write-UILogEntry -Message ("Adding event handlers for [{0}]" -f $searchText) -Source 'EventHandler' -Severity 0
                $syncHash[$searchInput].Add_TextChanged({
                    param($sender, $eventArgs)
                    $searchText = $sender.Text.Trim()
                    Write-UILogEntry -Message ("Filtering {0} for search criteria: {1}" -f $listView,$searchText) -Source $searchInput -Severity 0

                    # Filter ListView items based on search text
                    if ([string]::IsNullOrEmpty($searchText)) {
                        $syncHash[$listview].Items.Filter = $null  # Reset filter if empty
                    } else {
                        $syncHash[$listview].Items.Filter = { param($item)
                            $item.Name -match [regex]::Escape($searchText)
                        }
                    }
                }.GetNewClosure())  # Ensures each event has the correct variables

                Write-UILogEntry -Message ("Adding event handlers for [{0}]" -f $statusType) -Source 'EventHandler' -Severity 0
                $syncHash[$statusType].Add_SelectionChanged({
                    param($sender, $eventArgs)
                    $status = $sender.SelectedItem
                    Write-UILogEntry -Message ("Filtering {0} for status type: {1}" -f $listView,$status) -Source $statusType -Severity 0

                    # Filter ListView items based on status type
                    if ($status -eq 'All') {
                        $syncHash[$listview].Items.Filter = $null  # Reset filter if empty
                    } else {
                        $syncHash[$listview].Items.Filter = { param($item)
                            $item.Status -eq $status
                        }
                    }
                }.GetNewClosure())  # Ensures each event has the correct variables
            }

            #make selectall and selectnone buttons work
            Write-UILogEntry -Message ("Adding event handlers for select all button [{0}] and select none button [{1}]" -f $selectAll,$selectNone) -Source 'EventHandler' -Severity 0
            $syncHash[$selectAll].Add_Click({
                If ($syncHash[$listview].items.Count -gt 0) {
                    # Ensure event assignment runs on the UI thread
                    $syncHash.Window.Dispatcher.Invoke([action]{
                        Write-UILogEntry -Message ("Selecting all items for {0}" -f $listview) -Source $selectAll -Severity 0
                        $syncHash[$listview].SelectAll();

                        #If solution type check all chlAddDownloaded items if is required
                        #require setting is in $syncHash.Config.SolutionGroupedModules.AdditionalDownloads
                        If($_.Type -eq 'Solution'){
                            $adddownloads = $_.Elements | Where-Object { $_ -match "^chkAddDownload(.+)$" }
                            Foreach($adddownload in $adddownloads){
                                $checkBoxName = $adddownload
                                $downloadName = $adddownload -replace 'chkAddDownload','' -replace ($_.Name -replace '\W+')
                                If( ($syncHash.Config.SolutionGroupedModules.Additionaldownloads | Where-Object { ($_.Name -replace '\W+') -eq $downloadName}).Required ){
                                    Write-UILogEntry -Message ("Checking item [{0}] for download [{1}]" -f $checkBoxName,$downloadName) -Source $selectAll -Severity 0
                                    $syncHash[$checkBoxName].IsChecked = $true
                                    $syncHash.AdditionalDownloads += $downloadName
                                }
                            }
                        }
                    })
                }
            }.GetNewClosure())

            # Add event handler for select none button
            $syncHash[$selectNone].Add_Click({
                If ($syncHash[$listview].items.Count -gt 0) {
                    # Ensure event assignment runs on the UI thread
                    $syncHash.Window.Dispatcher.Invoke([action]{
                        Write-UILogEntry -Message ("Deselecting all items for {0}" -f $listview) -Source $selectNone -Severity 0
                        $syncHash[$listview].SelectedItems.Clear()
                    })
                }
            }.GetNewClosure())

            # Add event handler for select updates button
            $syncHash[$selectupdatesbtn].Add_Click({
                If ($syncHash[$listview].items.Count -gt 0) {
                    # Ensure event assignment runs on the UI thread
                    $syncHash.Window.Dispatcher.Invoke([action]{
                        Write-UILogEntry -Message ("Selecting all items with updates for {0}" -f $listview) -Source $selectupdatesbtn -Severity 0
                        $syncHash[$listview].Items | Where-Object { $_.Status -eq 'Update Available' } | ForEach-Object {
                            $syncHash[$listview].SelectedItems.Add($_)
                        }
                    })
                }
            }.GetNewClosure())

            # Attach SelectionChanged event: add selected to main list
            $syncHash[$listview].Add_SelectionChanged({
                param($sender, $eventArgs)

                # Get selected items
                $selectedItems = $sender.SelectedItems | ForEach-Object { $_.Name }

                #update selected count
                #Update-UIText -TextBox $syncHash[$selectedCount] -Text $selectedItems.Count
                $syncHash[$selectedCount].Text = $selectedItems.Count

                # Ensure unique addition
                foreach ($item in $selectedItems) {
                    if ($mainListBox.Items -notcontains $item) {
                        Write-UILogEntry -Message ("Adding item [{0}] to main ListBox [{1}]" -f $item,$mainListBox.Name) -Source $listview -Severity 0
                        $mainListBox.Items.Add($item)
                    }
                }

                # Remove items if deselected
                $deselectedItems = $eventArgs.RemovedItems | ForEach-Object { $_.Name }
                foreach ($item in $deselectedItems) {
                    Write-UILogEntry -Message ("Removing item [{0}] from main ListBox [{1}]" -f $item,$mainListBox.Name) -Source $listview -Severity 0
                    $mainListBox.Items.Remove($item)
                }

                #update install button if list has items
                If($mainListBox.Items.Count -gt 0){
                    $syncHash.btnInstall.Content = "Install Selected Modules"
                    $syncHash.chkModuleRepairSelected.IsEnabled = $true
                }Else{
                    $syncHash.chkModuleRepairSelected.IsChecked = $false
                    $syncHash.chkModuleRepairSelected.IsEnabled = $false
                    # if no checkboxes are checked, change button text to "Update Modules"
                    If($syncHash.chkModuleRemoveDuplicates.IsChecked -or $syncHash.chkModuleRemoveAll.IsChecked -or $syncHash.chkModuleAutoUpdate.IsChecked){
                        $syncHash.btnInstall.Content = "Update Modules"
                    }
                }

                #update total selected count
                $syncHash.txtTotalSelectedCount.Text = $mainListBox.Items.Count
            }.GetNewClosure())
        }

        #onload; update button text if checkboxes are checked
        If( ($syncHash.chkModuleRemoveDuplicates.IsChecked -or $syncHash.chkModuleRemoveAll.IsChecked -or $syncHash.chkModuleAutoUpdate.IsChecked) -and $mainListBox.Items.Count -eq 0){
            $syncHash.btnInstall.Content = "Update Modules"
        }

        #get all checkboxes named chkAddDownload) and add the name to a variable if checked
        $syncHash.Keys | Where-Object { $_ -match "^chkAddDownload(.+)$" } | ForEach-Object {
            $checkBoxName = $_
            $downloadName = $matches[1]
            $syncHash[$checkBoxName].Add_Checked({
                $syncHash.AdditionalDownloads += $downloadName
            }.GetNewClosure())
        }

        #if repair selected is checked, unchecked conflicting buttons
        $syncHash.chkModuleRepairSelected.Add_Checked({
            If($syncHash.chkModuleRepairSelected.IsChecked){
                $syncHash.chkModuleRemoveAll.IsChecked = $false
                $syncHash.chkModuleRemoveDuplicates.IsChecked = $false
                #change install button text
                $syncHash.btnInstall.Content = "Repair Selected Modules"
            }
        }.GetNewClosure())

        #detect when repair selected is unchecked
        $syncHash.chkModuleRepairSelected.Add_Unchecked({
            If(-Not $syncHash.chkModuleRepairSelected.IsChecked){
                $syncHash.btnInstall.Content = "Install Selected Modules"
            }
        }.GetNewClosure())

        #if remove all is checked, unchecked conflicting buttons
        $syncHash.chkModuleRemoveAll.Add_Checked({
            If($syncHash.chkModuleRemoveAll.IsChecked){
                $syncHash.chkModuleRepairSelected.IsChecked = $false
                If( -not($syncHash.chkModuleRemoveDuplicates.IsChecked) -and -not($syncHash.chkModuleAutoUpdate.IsChecked) -and ($mainListBox.Items.Count -gt 0) ){
                    $syncHash.btnInstall.Content = "Remove Selected Modules"
                }ElseIf($mainListBox.Items.Count -gt 0){
                    $syncHash.btnInstall.Content = "Install Selected Modules"
                }
            }
            
        }.GetNewClosure())

        #if remove duplicates is checked, unchecked conflicting buttons
        $syncHash.chkModuleRemoveDuplicates.Add_Checked({
            If($syncHash.chkModuleRemoveDuplicates.IsChecked){
                $syncHash.chkModuleRepairSelected.IsChecked = $false
            }
        }.GetNewClosure())


        #Allow UI to be dragged around screen
        $syncHash.Window.Add_MouseLeftButtonDown( {
            $syncHash.Window.DragMove()
        })

        #action for exit button
        $syncHash.btnExit.Add_Click({
            Close-UIMainWindow
        })


        #action for exit button
        $syncHash.btnInstall.Add_Click({
            $syncHash.btnInstall.IsEnabled = $false
            If( ([Boolean]::Parse($syncHash.Config.DefaultSettings.IgnorePoshProcessCheck)) -or ($syncHash.DisableProcessCheck) ){
                $result = $true
            }Else{
                $result = Confirm-RunningApps
            }
           
            If($result){
                #Build OutputData OBJECT for external installer or Sequence Window
                $syncHash.OutputData = @{
                    DoAction = $true
                    RemoveAll = $syncHash.chkModuleRemoveAll.IsChecked
                    AutoUpdate = $syncHash.chkModuleAutoUpdate.IsChecked
                    DuplicateCleanup = $syncHash.chkModuleRemoveDuplicates.IsChecked
                    RepairSelected = $syncHash.chkModuleRepairSelected.IsChecked
                    PS7install = $syncHash.chkModuleInstallForPS7.IsChecked
                    InstallUserContext = $syncHash.chkModuleInstallUserContext.IsChecked
                    #Currently selected modules are a list meaning they will be the latest version to install
                    SelectedModules = $syncHash.lbSelectedModules.Items
                    InstalledModules = $syncHash.InstalledModules
                    AdditionalDownloads = $syncHash.AdditionalDownloads | Select -Unique
                }
                Close-UIMainWindow
            }
            Else{
                $syncHash.btnInstall.IsEnabled = $true
            }
        })

        $syncHash.Window.Add_KeyDown({
            #allow window in back if ESC is hit
            if ( ($_.Key -match 'Esc') -and $syncHash.Window.Topmost ) {
                $syncHash.Window.Topmost = $false
            }
            #and set inf front if hit again
            ElseIf ( ($_.Key -match 'Esc') ) {
                $syncHash.Window.Topmost = $true
            }
        })

        Write-UILogEntry -Message "UI is ready" -Source 'Show-UIMainWindow' -Severity 0
        # Before the UI is displayed
        # Create a timer dispatcher to watch for value change externally on regular interval
        # update those values when found using scriptblock ($updateblock)
        $syncHash.Window.Add_SourceInitialized({
            ## create a timer
            $timer = new-object System.Windows.Threading.DispatcherTimer
            ## set to fire 4 times every second
            $timer.Interval = [TimeSpan]"0:0:0.01"
            ## invoke the $updateBlock after each fire
            $timer.Add_Tick( $updateUI )
            ## start the timer
            $timer.Start()

            if( -Not($timer.IsEnabled) ) {
               $syncHash.Error = "Timer didn't start"
            }
        })

        #Add smooth closing for Window
        $syncHash.Window.Add_Loaded({ $syncHash.isLoaded = $True })
        $syncHash.Window.Add_Closing({ $syncHash.isClosing = $True; Close-UIMainWindow })
        $syncHash.Window.Add_Closed({ $syncHash.isClosed = $True })

        #make sure this display on top of every window
        $syncHash.Window.Topmost = $true
        If($syncHash.TopPosition){
            $syncHash.Window.WindowStartupLocation = 'CenterOwner'
            $syncHash.Window.Top = $syncHash.TopPosition
        }
        If($syncHash.LeftPosition){
            $syncHash.Window.WindowStartupLocation = 'CenterOwner'
            $syncHash.Window.Left = $syncHash.LeftPosition
        }
        $syncHash.window.ShowDialog()
        $syncHash.Error = $Error
    }) # end scriptblock
    #collect data from runspace
    $Data = $syncHash
    #invoke scriptblock in runspace
    $PowerShellCommand.Runspace = $PSRunSpace
    $AsyncHandle = $PowerShellCommand.BeginInvoke()

    If($Wait){
        #wait until runspace is completed before ending
        do {
            Start-sleep -m 100 }
        while (!$AsyncHandle.IsCompleted)
        #end invoked process
        $null = $PowerShellCommand.EndInvoke($AsyncHandle)
    }

    #cleanup registered object
    Register-ObjectEvent -InputObject $syncHash.Runspace `
            -EventName 'AvailabilityChanged' `
            -Action {
                    if($Sender.RunspaceAvailability -eq "Available")
                    {
                        $Sender.Closeasync()
                        $Sender.Dispose()
                        # Speed up resource release by calling the garbage collector explicitly.
                        # Note that this will pause *all* threads briefly.
                        [GC]::Collect()
                    }
                } | Out-Null
    return $Data
}
#endregion

##=============================================
## BUILD MAIN UI
##=============================================
Function Add-UIModuleSubTabItem{
    Param(
        [String]$Name,
        [String]$Description,
        [switch]$Passthru
    )

    If($null -eq $Description){
        $Description = "Available $($Name) PowerShell Module"
    }
    #remove special characters from name and spaces
    $FriendlyName = $Name -replace '[^a-zA-Z0-9]',''

    If($Name -match "beta"){
        $HeaderName = "Beta"
    }
    Else{
        $HeaderName = "Released"
    }
    $TabItemContentXaml = @"
            <TabItem x:Name="tabItem$($FriendlyName)" Header="$HeaderName" Style="{StaticResource LineTabStyle}">
                <Grid Margin="0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="938*"/>
                        <ColumnDefinition Width="269*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Margin="10,10,10,10" Grid.Column="0">
                        <Label Content="$Name PowerShell Modules" HorizontalAlignment="Left" FontSize="20" Foreground="Black" FontWeight="Bold"/>
                        <Label Content="$Description" HorizontalAlignment="Left" FontSize="10" Foreground="Gray" Width="439"/>
                        <Separator Margin="15,10,15,0"/>
                        <Separator Margin="15,0,15,10"/>

                        <Grid Height="31">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="51*" />
                                <ColumnDefinition Width="178*"/>
                                <ColumnDefinition Width="63*" />
                                <ColumnDefinition Width="166*"/>
                            </Grid.ColumnDefinitions>
                            <Label Content="Search by Module :" Grid.Column="0" FontSize="12" Foreground="Gray" HorizontalAlignment="Right" VerticalAlignment="Center" Height="25" Margin="0,0,345,0" Grid.ColumnSpan="2"/>
                            <TextBox x:Name="txt$($FriendlyName)ModuleNameSearch" Grid.Row="0" Grid.Column="1" HorizontalAlignment="Center" Width="337" Margin="0,4,0,5" />
                            <Label Content="Filter by Status :" Grid.Column="2" HorizontalAlignment="Left" FontSize="12" Foreground="Gray" VerticalAlignment="Center" Height="25" Width="96"/>
                            <ComboBox x:Name="cmb$($FriendlyName)ModuleStatusType" Grid.Column="2" HorizontalAlignment="Left" Margin="96,0,0,0" VerticalAlignment="Center" Width="348" Grid.ColumnSpan="2"/>

                        </Grid>

                        <Grid Height="31">
                           <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120*" />
                                <ColumnDefinition Width="120*"/>
                                <ColumnDefinition Width="275*"/>
                                <ColumnDefinition Width="100*"/>
                                <ColumnDefinition Width="50*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="btn$($FriendlyName)SelectAll" Grid.Column="0" Content="Select All" HorizontalAlignment="Left" Width="141" Height="20"  />
                            <Button x:Name="btn$($FriendlyName)SelectNone" Grid.Column="1" Content="Select None" HorizontalAlignment="Left" Width="141" Height="20"  />
                            <Label Content="Modules Selected:" Grid.Column="3" HorizontalAlignment="Right" VerticalAlignment="Center" HorizontalContentAlignment="Right" Height="25" Margin="2,0,0,0" />
                            <TextBox x:Name="txt$($FriendlyName)SelectedCount" Grid.Column="4" Text="0" HorizontalAlignment="Left" Width="33" IsEnabled="False" BorderThickness="0" VerticalAlignment="Center" Height="15"/>
                        </Grid>
                        <ListView x:Name="lst$($FriendlyName)ModuleList" HorizontalAlignment="Center" Height="371" Margin="0,10,0,0" Width="916" SelectionMode="Multiple">
                            <ListView.View>
                                <GridView>

                                    <GridViewColumn Header="Check" DisplayMemberBinding="{Binding Check}" Width="40">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <CheckBox IsChecked="{Binding IsSelected}" />
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="250" />
                                    <GridViewColumn Header="Current Version" DisplayMemberBinding="{Binding cVersion}" Width="90" />
                                    <GridViewColumn Header="Latest Version" DisplayMemberBinding="{Binding lVersion}" Width="90" />
                                    <GridViewColumn Header="Owner" DisplayMemberBinding="{Binding Owner}" Width="140" />
                                    <GridViewColumn Header="Count" DisplayMemberBinding="{Binding Count}" Width="50"/>
                                    <GridViewColumn Header="Status" DisplayMemberBinding="{Binding Status}" Width="120" />
                                    <GridViewColumn Header="Installed Versions" Width="120">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <ComboBox ItemsSource="{Binding Version}" SelectedItem="{Binding SelectedVersion}" Width="80"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>

                        <Grid Height="43">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="110*" />
                                <ColumnDefinition Width="50*"/>
                                <ColumnDefinition Width="300*"/>
                                <ColumnDefinition Width="175*" />
                                <ColumnDefinition Width="150*" />
                            </Grid.ColumnDefinitions>
                            <Label Content="Total Module Count:" Grid.Column="0" HorizontalAlignment="Center" VerticalAlignment="Center" HorizontalContentAlignment="Right" Height="25" />
                            <TextBox x:Name="txt$($FriendlyName)ModuleCount" Grid.Column="1" Text="0" HorizontalAlignment="Left" Width="33" IsEnabled="False" BorderThickness="0" VerticalAlignment="Center"/>
                            <Button x:Name="btn$($FriendlyName)SelectUpdates" Grid.Column="4" Content="Select Update Available" HorizontalAlignment="Left" VerticalAlignment="Center" Width="141" Height="35" Margin="9,0,0,0" />

                        </Grid>
                    </StackPanel>
                </Grid>
            </TabItem>
"@

    #get elements by matching x:name="elementName"
    $matches = [regex]::Matches($TabItemContentXaml, 'x:Name="([^"]+)"')
    $Elements = $matches | ForEach-Object { $_.Groups[1].Value }

    #build psobject output with xaml, names and types
    $TabItemData = [PSCustomObject]@{
        Name = $Name
        Type = "Module"
        Elements = $Elements
        #Tab = "tabItem$($FriendlyName)"
        #ListView = "lst$($FriendlyName)Module"
        #Counter = "txt$($FriendlyName)ModuleCount"
        #SelectAllButton = "btn$($FriendlyName)SelectAll"
        #SelectNoneButton = "btn$($FriendlyName)SelectNone"
        #SearchTextBox = "txt$($FriendlyName)ModuleNameSearch"
        #StatusDropdown = "cmb$($FriendlyName)ModuleStatusType"
        #UpdateButton = "btn$($FriendlyName))SolutionUpdates"
        Xaml = $TabItemContentXaml
    }

    If($Passthru){ return $TabItemContentXaml }
    Else{ return $TabItemData }
}

#region FUNCTION: Loader for modern checkbox style
Function Add-UIModuleTabItem{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        $ModuleConfig,
        [Parameter(Mandatory=$true)]
        [bool]$Selected,
        [Parameter(Mandatory=$false)]
        [switch]$Passthru
    )

    #remove special characters from name and spaces
    $FriendlyName = $ModuleConfig.Name -replace '[^a-zA-Z0-9]',''
    $ShowBetaTabItem = [Boolean]::Parse($ModuleConfig.ShowBetaTab)

    Write-Verbose '$ModuleConfig is:'
    Write-Verbose $ModuleConfig

    #process tabs
    $TabObjectData = @()
    $ReleaseTabItemXaml = (Add-UIModuleSubTabItem -Name $ModuleConfig.Name -Description $ModuleConfig.Description)
    $TabObjectData += $ReleaseTabItemXaml

    If($ShowBetaTabItem){
        $BetaTabItemXaml = (Add-UIModuleSubTabItem -Name ($ModuleConfig.Name + ' Beta') -Description $ModuleConfig.Description)
        $TabObjectData += $BetaTabItemXaml
    }

    #build tab item xaml
    $TabItemXaml = @()
    $TabItemXaml += "<TabItem Header=`"$($ModuleConfig.Name)`" Style=`"{StaticResource WhiteTabItems}`" Width=`"150`" IsSelected=`"$Selected`">"
    $TabItemXaml += "   <Grid HorizontalAlignment=`"Left`">"
    $TabItemXaml += "       <TabControl x:Name=`"tabControlSub$($FriendlyName)`" HorizontalAlignment=`"Left`" Height=`"656`" VerticalAlignment=`"Top`" BorderThickness=`"0`" Width=`"1210`">"
    $TabItemXaml += $ReleaseTabItemXaml.Xaml
    If($ShowBetaTabItem){
        $TabItemXaml += $BetaTabItemXaml.Xaml
    }
    $TabItemXaml += "       </TabControl>"
    $TabItemXaml += "   </Grid>"
    $TabItemXaml += "</TabItem>"

    If($Passthru){ return $TabItemXaml }
    Else{ return $TabObjectData }
}



Function Add-UISolutionSubTabItem{
    Param(
        [String]$Name,
        [String]$Description,
        $AdditionalXaml,
        [switch]$Passthru
    )

    If($null -eq $Description){
        $Description = "Available $($Name) PowerShell Solution"
    }
    #remove special characters from name and spaces
    $FriendlyName = $Name -replace '[^a-zA-Z0-9]',''

    $TabItemContentXaml = @"
            <TabItem x:Name="tab$($FriendlyName)Solution" Header="$($FriendlyName)" Style="{StaticResource LineTabStyle}" >
                <Grid Margin="0">

                    <StackPanel Margin="10,10,264,10">
                        <Label Content="$($FriendlyName) Modules" HorizontalAlignment="Left" FontSize="20" Foreground="Black" FontWeight="Bold"/>
                        <Label Content="$Description" HorizontalAlignment="Left" FontSize="10" Foreground="Gray" Width="439"/>
                        <Separator Margin="15,0,15,10"/>

                        <Grid Height="31">
                           <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="120*" />
                                <ColumnDefinition Width="120*"/>
                                <ColumnDefinition Width="275*"/>
                                <ColumnDefinition Width="100*"/>
                                <ColumnDefinition Width="50*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="btn$($FriendlyName)SelectAll" Grid.Column="0" Content="Select Solution" HorizontalAlignment="Left" Width="141" Height="20"  />
                            <Button x:Name="btn$($FriendlyName)SelectNone" Grid.Column="1" Content="Remove Solution" HorizontalAlignment="Left" Width="141" Height="20"  />
                            <Label Content="Modules Selected:" Grid.Column="3" HorizontalAlignment="Right" VerticalAlignment="Center" HorizontalContentAlignment="Right" Height="25" Margin="2,0,0,0" />
                            <TextBox x:Name="txt$($FriendlyName)SelectedCount" Grid.Column="4" Text="0" HorizontalAlignment="Left" Width="33" IsEnabled="False" BorderThickness="0" VerticalAlignment="Center" Height="15"/>
                        </Grid>
                        <ListView x:Name="lst$($FriendlyName)SolutionList" HorizontalAlignment="Center" Height="230" Margin="0,10,0,0" Width="916" SelectionMode="Multiple"
                                        IsHitTestVisible="True"
                                        ScrollViewer.VerticalScrollBarVisibility="Auto"
                                        ScrollViewer.CanContentScroll="True">
                            <ListView.View>
                                <GridView>

                                    <GridViewColumn Header="Check" DisplayMemberBinding="{Binding Check}" Width="40">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <CheckBox IsChecked="{Binding IsSelected}" />
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="250" />
                                    <GridViewColumn Header="Current Version" DisplayMemberBinding="{Binding cVersion}" Width="90" />
                                    <GridViewColumn Header="Latest Version" DisplayMemberBinding="{Binding lVersion}" Width="90" />
                                    <GridViewColumn Header="Owner" DisplayMemberBinding="{Binding Owner}" Width="150" />
                                    <GridViewColumn Header="Count" DisplayMemberBinding="{Binding Count}" Width="50"/>
                                    <GridViewColumn Header="Status" DisplayMemberBinding="{Binding Status}" Width="120" />
                                    <GridViewColumn Header="Install Version" Width="120">
                                        <GridViewColumn.CellTemplate>
                                            <DataTemplate>
                                                <ComboBox ItemsSource="{Binding Version}" SelectedItem="{Binding SelectedVersion}" Width="80"/>
                                            </DataTemplate>
                                        </GridViewColumn.CellTemplate>
                                    </GridViewColumn>
                                </GridView>
                            </ListView.View>
                        </ListView>

                        <Grid Height="43">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="100*" />
                                <ColumnDefinition Width="50*"/>
                                <ColumnDefinition Width="300*"/>
                                <ColumnDefinition Width="175*" />
                                <ColumnDefinition Width="150*" />
                            </Grid.ColumnDefinitions>
                            <Label Content="Total Module Count:" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center" HorizontalContentAlignment="Right" Height="25" Margin="-2,0,0,0" />
                            <TextBox x:Name="txt$($FriendlyName)ModuleCount" Grid.Column="1" Text="0" HorizontalAlignment="Left" Width="33" IsEnabled="False" BorderThickness="0" VerticalAlignment="Center" Height="15"/>
                            <Button x:Name="btn$($FriendlyName)SelectUpdates" Grid.Column="4" Content="Select Update Available" HorizontalAlignment="Center" VerticalAlignment="Center" Width="135" Height="35"  />

                        </Grid>
                        $AdditionalXaml
                    </StackPanel>


                </Grid>
            </TabItem>
"@

    #get elements by matching x:name="elementName"
    $matches = [regex]::Matches($TabItemContentXaml, 'x:Name="([^"]+)"')
    $Elements = $matches | ForEach-Object { $_.Groups[1].Value }

    #build psobject output with xaml, names and types
    $TabItemData = [PSCustomObject]@{
        Name = $Name
        Type = 'Solution'
        Elements = $Elements
        Xaml = $TabItemContentXaml
    }

    If($Passthru){ return $TabItemContentXaml }
    Else{ return $TabItemData }
}

#region FUNCTION: Loader for modern checkbox style
Function Add-UISolutionTabItem{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $SolutionConfig,
        [Parameter(Mandatory=$false)]
        [switch]$Passthru
    )

    Begin{


        #process tabs
        $TabObjectData = @()
        #build tab item xaml
        $TabItemXaml = @()
        $TabItemXaml += "<TabItem Header=`"Solutions`" Style=`"{StaticResource WhiteTabItems}`" Width=`"150`">"
        $TabItemXaml += "   <Grid HorizontalAlignment=`"Left`">"
        $TabItemXaml += "       <TabControl HorizontalAlignment=`"Left`" Height=`"656`" VerticalAlignment=`"Top`" BorderThickness=`"0`" Width=`"1210`">"
    }
    Process{
         Write-Verbose '$SolutionConfig is:'
         Write-Verbose $SolutionConfig
        <#
        "<Label x:Name="lbl$($FriendlyName)Additional" Content="$($FriendlyName) Additional Items" HorizontalAlignment="Left" FontSize="20" Foreground="Black" FontWeight="Bold"/>
        "<CheckBox x:Name=`"chk$($FriendlyName)OPAInstall`" Content=`"Open Policy Agent (Executable)`" IsChecked=`"True`" Margin=`"10`" FontSize=`"14`" />"
        #>
        If($SolutionConfig.AdditionalDownloads){
            #build additional items in xaml
            $AdditionalItemXaml = @()
            $AdditionalItemXaml += "<Label Content=`"$($SolutionConfig.Name) Additional Items`" HorizontalAlignment=`"Left`" FontSize=`"20`" Foreground=`"Black`" FontWeight=`"Bold`"/>"
            ForEach($Item in $SolutionConfig.AdditionalDownloads){
                $FriendlyItemName = $SolutionConfig.Name + $Item.Name -replace '[^a-zA-Z0-9]',''
                $AdditionalItemXaml += "<CheckBox x:Name=`"chkAddDownload$($FriendlyItemName)`" Content=`"$($Item.Name) [$($Item.Type)]`" Margin=`"10`" FontSize=`"14`" />"
            }
            $SolutionTabItemXaml = (Add-UISolutionSubTabItem -Name $SolutionConfig.Name -Description $SolutionConfig.Description -AdditionalXaml $AdditionalItemXaml)
        }Else{
            $SolutionTabItemXaml = (Add-UISolutionSubTabItem -Name $SolutionConfig.Name -Description $SolutionConfig.Description)
        }

        $TabObjectData += $SolutionTabItemXaml
        $TabItemXaml += $SolutionTabItemXaml.Xaml

    }
    End{
        $TabItemXaml += "       </TabControl>"
        $TabItemXaml += "   </Grid>"
        $TabItemXaml += "</TabItem>"

        If($Passthru){ return  $TabItemXaml }
        Else{ return $TabObjectData }
    }
}


# Function to get window coordinates
function Get-RunspacePosition {
    param(
        [Parameter(Mandatory = $true)]
        $Runspace
    )

    $Runspace.Window.Dispatcher.Invoke([action]{
        return @{
            Left   = $Runspace.Window.Left
            Top    = $Runspace.Window.Top
            #Width  = $Runspace.Window.Width
            #Height = $Runspace.Window.Height
        }
    })
}


Function Find-ModuleRetry{
    [CmdletBinding()]
    param(
        [string[]]$Name,
        [string]$RequiredVersion,
        [string]$Repository = 'PSGallery',
        [int]$RetryCount = 3
    )

    # Ensure Name is not null
    If (-not $Name -or $Name.Count -eq 0) {
       # Write-Error "ERROR: Module Name is NULL or EMPTY!"
        return @()
    }

    # Run find-module with retry
    $Retry = 0
    $FoundModule = @()

    $FindModule = @{
        Name = $Name
        Repository = $Repository
    }

    # Add version if provided
    If ($RequiredVersion) {
        $FindModule["RequiredVersion"] = $RequiredVersion
    }

    # Debug output before calling Find-Module
    Write-Verbose "Searching with Parameters:"
    Write-Verbose ("  Name: {0}" -f ($FindModule.Name -join ','))
    If ($FindModule.ContainsKey("RequiredVersion")) {
        Write-Verbose ("  RequiredVersion: {0}" -f $FindModule.RequiredVersion)
    }

    do {
        Write-LogEntry ("RUNNING: Find-Module [{0}]. Retry Count [{1}]" -f ($FindModule.Name -join ','), $Retry) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
        $FoundModule += Find-Module @FindModule -ErrorAction SilentlyContinue  | Select-Object Version, Name, Repository, Description, Author, CompanyName, Dependencies
        $Retry++
    } until (($FoundModule.Count -gt 0) -or ($Retry -eq $RetryCount))

    Write-LogEntry ("Found [{0}] modules" -f $FoundModule.Count) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
    return $FoundModule
}

Function Get-ModuleListData{

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $ModuleGroupItem,
        [switch]$Passthru
    )

    Begin{
        $ModuleObject = @()
    }
    Process{
        # If included modules are specified, only search for those
        $ModuleList = @()
        If($null -ne $ModuleGroupItem.ModuleSearch){
            $ModuleParam = @{
                Name = $ModuleGroupItem.ModuleSearch
            }
            $SearchMessage = ("Searching for modules with criteria [{0}] from the PowerShell Gallery..." -f $ModuleGroupItem.ModuleSearch)

            If($ModuleGroupItem.ModuleVersion -gt 0){
                $ModuleParam += @{
                    RequiredVersion = $ModuleGroupItem.ModuleVersion
                }
                $SearchMessage += ("Searching for modules with criteria [{0}] and version [{1}] from the PowerShell Gallery..." -f $ModuleGroupItem.ModuleSearch,$ModuleGroupItem.ModuleVersion)
            }

            Write-Verbose $SearchMessage
            Write-LogEntry -Message $SearchMessage -Source $MyInvocation.MyCommand.Name -Severity 0
            $FoundModules = Find-ModuleRetry @ModuleParam
            If($ModuleGroupItem.ModuleAuthors.count -gt 0){
                $ModuleListCurrentCount = $FoundModules.Count
                $FoundModules = $FoundModules | Where-Object { $_.Author -in $ModuleGroupItem.ModuleAuthors }
                Write-LogEntry -Message ("Only including modules with specified authors from [{0}]. Removed [{1}]" -f $ModuleGroupItem.Name,($ModuleListCurrentCount - $FoundModules.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
            }
        }

        #add to module list
        $ModuleList += $FoundModules

        If($ModuleGroupItem.IncludedModules.count -gt 0){
            Foreach($Module in $SolutionGroupItem.IncludedModules){
                If($Module.ModuleName){
                    $ModuleParam = @{
                        Name = $Module.ModuleName
                    }
                }Else{
                    $ModuleParam = @{
                        Name = $Module
                    }
                }

                If($Module.ModuleVersion){
                    $ModuleParam += @{
                        RequiredVersion = $Module.ModuleVersion
                    }
                }
                Write-LogEntry -Message ("Searching for included modules [{0}] for module group [{1}]" -f $Module.ModuleName,$ModuleGroupItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                $IncludedModules = Find-ModuleRetry @ModuleParam
                If($Module.ModuleAuthors.count -gt 0){
                    $ModuleListCurrentCount = $FoundModules.Count
                    $IncludedModules = $IncludedModules | Where-Object { $_.Author -in $Module.ModuleAuthors }
                    Write-LogEntry -Message ("Only including modules with specified authors from [{0}]. Removed [{1}]" -f $ModuleGroupItem.Name,($ModuleListCurrentCount - $IncludedModules.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                }
            }

            #add to module list
            $ModuleList += $IncludedModules
        }
        #if the module is not beta, add beta to exclusions just in case of similar names
        #EXAMPLE: Microsoft.Graph.Beta vs Microsoft.Graph

        if ( $ModuleGroupItem.Name -notmatch "Beta|Preview" ) {
            #exclude beta from elements
            $ModuleListCurrentCount = $ModuleList.Count
            $ModuleList = $ModuleList | Where-Object { $_.Name -notmatch "Beta|Preview" }
            Write-LogEntry -Message ("Excluding Beta modules from [{0}]. Removed [{1}]" -f $ModuleGroupItem.Name,($ModuleListCurrentCount - $ModuleList.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
        }

        #filter out excluded modules
        If($ModuleGroupItem.ExcludedModules.count -gt 0){
            $ModuleListCurrentCount = $ModuleList.Count
            $ModuleList = $ModuleList | Where-Object { $_.Name -notmatch $ModuleGroupItem.ExcludedModules }
            Write-LogEntry ("Excluding modules from [{0}]. Removed [{1}]" -f $ModuleGroupItem.Name,($ModuleListCurrentCount - $ModuleList.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0
        }

        Write-Verbose ("Adding [{0}] modules to module group [{1}]" -f $ModuleList.Count,$ModuleGroupItem.Name)

        $ModuleObject += New-Object PSObject -Property @{
            GroupName = $ModuleGroupItem.Name
            KeyName = $ModuleGroupItem.Name -replace '[^a-zA-Z0-9]',''
            ModuleList = @($ModuleList | Sort-Object -Property Name -Unique)
        }
    }
    End{
        If($Passthru){ return $ModuleList }Else{ return $ModuleObject }
    }

}

Function Get-SolutionListData{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $SolutionGroupItem,
        [switch]$Passthru
    )

    Begin{
        $SolutionObject = @()

    }
    Process{

        $ModuleList = @()
        $ModulesListUrl = @()

        #retrieve module list from url
        If($SolutionGroupItem.ModulesListUrl.count -gt 0)
        {
            Try{
                Write-LogEntry -Message ("Downloading module list from [{0}]" -f $SolutionGroupItem.ModulesListUrl) -Source $MyInvocation.MyCommand.Name -Severity 0
                $ProgressPreference = 'SilentlyContinue'    # Subsequent calls do not display UI.
                $ModulesListUrl = Invoke-WebRequest -Uri $SolutionGroupItem.ModulesListUrl

                If($ModulesListUrl.StatusCode -eq 200){
                    #check for extension
                    $Filename = (Split-Path -Path $SolutionGroupItem.ModulesListUrl -Leaf)
                    switch($Filename.split('.')[-1]){
                        'json'{
                            Write-LogEntry -Message ("Converting JSON: {0}" -f $Filename) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                            $ModulesListUrl = $ModulesListUrl | ConvertFrom-Json
                        }
                        'xml'{
                            Write-LogEntry -Message ("Converting XML: {0}" -f $Filename) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                            $ModulesListUrl = [xml]$ModulesListUrl.Content
                        }
                        'txt'{
                            Write-LogEntry -Message ("Converting TXT: {0}" -f $Filename) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                            $ModulesListUrl = $ModulesListUrl.Content -split "`n"
                        }
                        default{
                            Write-LogEntry -Message ("Converting JSON: {0}" -f $Filename) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
                            $ModulesListUrl = $ModulesListUrl | ConvertFrom-Json
                        }
                    }
                }
            }catch{
                Write-LogEntry -Message ("Failed to download from [{0}]: {1}" -f $SolutionGroupItem.ModulesListUrl,$_.Exception.Message) -Source $MyInvocation.MyCommand.Name -Severity 3
                Write-Warning ("Failed to download from [{0}]: {1}" -f $SolutionGroupItem.ModulesListUrl,$_.Exception.Message)
            }
        }Else{
            Write-LogEntry -Message ("No module list url provided for [{0}]. Skipping" -f $SolutionGroupItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 2
        }

        If($SolutionGroupItem.ModulesQuery.length -gt 0){
            #$ModuleQuery = [ScriptBlock]::Create('$ModulesListUrl | Select-Object -ExpandProperty poshModule -Unique')
            Write-LogEntry -Message ("Executing query: [{0}]" -f $SolutionGroupItem.ModulesQuery) -Source $MyInvocation.MyCommand.Name -Severity 0 -Verbose
            $ModuleQuery = [ScriptBlock]::Create($SolutionGroupItem.ModulesQuery)
            $ModulesListUrl = Invoke-Command -ScriptBlock $ModuleQuery
        }


        #TEST  $ModulesListUrl = $UIConfig.SolutionGroupedModules[0].IncludedModules
        #TEST  $ModulesListUrl = $UIConfig.SolutionGroupedModules[1].IncludedModules
        If( $ModulesListUrl.count -gt 0)
        {
            $FoundModules = @()
            Foreach($Module in  $ModulesListUrl){
                If($Module.ModuleName){
                    $ModuleParam = @{
                        Name = $Module.ModuleName
                    }
                }Else{
                    $ModuleParam = @{
                        Name = $Module
                    }
                }

                If($Module.ModuleVersion){
                    $ModuleParam += @{
                        RequiredVersion = $Module.ModuleVersion
                    }
                }
                $FoundModules += Find-ModuleRetry @ModuleParam
            }
        }

        If($SolutionGroupItem.ModuleAuthors.count -gt 0)
        {
            $ModuleListCurrentCount = $FoundModules.Count
            $FoundModules = $FoundModules | Where-Object { $_.Author -in $SolutionGroupItem.ModuleAuthors }
            Write-LogEntry ("Only including modules with specified authors from [{0}]. Removed [{1}]" -f $SolutionGroupItem.Name,( $ModuleListCurrentCount - $FoundModules.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0
        }

        #filter out excluded modules
        If($SolutionGroupItem.ExcludedModules.count -gt 0){
            $ModuleListCurrentCount = $FoundModules.Count
            $FoundModules = $FoundModules | Where-Object { $_.Name -notmatch $SolutionGroupItem.ExcludedModules }
            Write-LogEntry ("Excluding modules from [{0}]. Removed [{1}]" -f $SolutionGroupItem.Name,($ModuleListCurrentCount - $FoundModules.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0
        }

        #add found modules to module list
        $ModuleList += $FoundModules

        If($SolutionGroupItem.IncludedModules.count -gt 0){
            <#
            Data looks like
            "IncludedModules": [
                {"ModuleName" : "$($FriendlyName)","ModuleVersion" : "1.4.0","ModuleAuthors" : "CISA"}
            ],
            #>
            #TEST $Module = $SolutionGroupItem.IncludedModules[0]
            Foreach($Module in $SolutionGroupItem.IncludedModules){
                $ModuleParam = @{
                    Name            = $Module.ModuleName
                }
                $message = "Searching for included Module: $($Module.ModuleName)"
                If($Module.ModuleVersion.length -gt 0){
                    $ModuleParam += @{
                        RequiredVersion = $Module.ModuleVersion
                    }
                    $message = "Searching for included Module: $($Module.ModuleName), Version: $($Module.ModuleVersion)"
                }

                Write-LogEntry ("{0} in solution group [{1}] " -f $message,$SolutionGroupItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 0
                $IncludedModules = Find-ModuleRetry @ModuleParam -Verbose

                If($Module.ModuleAuthors.count -gt 0){
                    $ModuleListCurrentCount = $IncludedModules.Count
                    $IncludedModules = $IncludedModules | Where-Object { $_.Author -in $Module.ModuleAuthors }
                    Write-LogEntry ("Only including modules with specified authors from [{0}]. Removed [{1}]" -f $SolutionGroupItem.Name,($ModuleListCurrentCount - $IncludedModules.Count)) -Source $MyInvocation.MyCommand.Name -Severity 0
                }
            }

            #add to module list
            $ModuleList += $IncludedModules
        }
        Write-Verbose ("Adding [{0}] modules for solution group [{1}]" -f $ModuleList.Count,$SolutionGroupItem.Name)

        $SolutionObject += New-Object PSObject -Property @{
            GroupName = $SolutionGroupItem.Name
            KeyName = $SolutionGroupItem.Name -replace '[^a-zA-Z0-9]',''
            ModuleList = @($ModuleList | Sort-Object -Property Name -Unique)
        }
    }
    End{
        If($Passthru){ return $ModuleList }Else{ return $SolutionObject }
    }
}



Function Merge-ModuleObject{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $ModuleObject,
        $MergeData
    )
    Begin {
        $newObjects = @()
    }
    Process{
        #Get the elements for the module
        $Elements = $MergeData | Where Name -eq "$($ModuleObject.Name)" | Select -ExpandProperty Elements
        Write-LogEntry -Message "Merging module object: $($ModuleObject.Name)" -Source $MyInvocation.MyCommand.Name -Severity 1

        # Always include the original object, adding "Elements"
        $primaryObject = [PSCustomObject]@{
            Name                     = $ModuleObject.Name
            Description              = $ModuleObject.Description
            PowerShellVersionSupport = $ModuleObject.PowerShellVersionSupport
            ModuleSearch             = $ModuleObject.ModuleSearch
            ModuleAuthors            = $ModuleObject.ModuleAuthors
            ModuleVersion            = $ModuleObject.ModuleVersion
            IncludePrereleaseVersion = $ModuleObject.IncludePrereleaseVersion
            IncludedModules          = $ModuleObject.IncludedModules
            ExcludedModules          = $BetaExclusions
            Elements                 = $Elements
        }

        # Add the primary object to the list
        $newObjects += $primaryObject

        # If ShowBetaTab is True, create the beta object
        if ($ModuleObject.ShowBetaTab -eq $true) {

            #Get the elements for the module
            $Elements = $MergeData | Where Name -eq "$($ModuleObject.Name + " Beta")" | Select -ExpandProperty Elements
            Write-LogEntry -Message "Merging module object: $($ModuleObject.Name) Beta" -Source $MyInvocation.MyCommand.Name -Severity 1


            $betaObject = [PSCustomObject]@{
                Name                     = $ModuleObject.Name + " Beta"
                Description              = $ModuleObject.Description.replace('v1.0','Beta')
                PowerShellVersionSupport = $ModuleObject.PowerShellVersionSupport
                ModuleSearch             = $ModuleObject.ModuleSearch -replace '.\*|\*','.Beta.*'
                ModuleAuthors            = $ModuleObject.ModuleAuthors
                ModuleVersion            = $ModuleObject.ModuleVersion
                IncludePrereleaseVersion = $ModuleObject.IncludePrereleaseVersion
                IncludedModules          = $ModuleObject.IncludedModules
                ExcludedModules          = $ModuleObject.ExcludedModules
                Elements                 = $Elements
            }

            # Add the beta object to the list
            $newObjects += $betaObject
        }

    }
    End{
        $newObjects
    }
}

Function Merge-SolutionObject{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $SolutionObject,
        $MergeData
    )
    Begin {
        $newObjects = @()
    }
    Process{
        #Get the elements for the module
        $Elements = $MergeData | Where Name -eq "$($SolutionObject.Name)" | Select -ExpandProperty Elements
        Write-LogEntry -Message "Merging solution object: $($SolutionObject.Name)" -Source $MyInvocation.MyCommand.Name -Severity 1

        # Always include the original object, adding "Elements"
        $primaryObject = [PSCustomObject]@{
            Name                     = $SolutionObject.Name
            Description              = $SolutionObject.Description
            ModulesListUrl           = $SolutionObject.ModulesListUrl
            ModulesQuery             = $SolutionObject.ModulesQuery
            ModuleAuthors            = $SolutionObject.ModuleAuthors
            IncludedModules          = $SolutionObject.IncludedModules
            ExcludedModules          = $SolutionObject.ExcludedModules
            AdditionalDownloads      = $SolutionObject.AdditionalDownloads
            Elements                 = $Elements
        }

        # Add the primary object to the list
        $newObjects += $primaryObject
    }
    End{
        $newObjects
    }
}



Function Write-LogEntry{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false,Position=2)]
        [string]$Source,

        [parameter(Mandatory=$false)]
        [ValidateSet(0,1,2,3,4,5)]
        [int16]$Severity = 1,

        [parameter(Mandatory=$false, HelpMessage="Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$OutputLogFile = $LogFilePath
    )
    ## Get the name of this function
    #[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
        $VerbosePreference = $PSCmdlet.SessionState.PSVariable.GetValue('VerbosePreference')
    }

    if (-not $PSBoundParameters.ContainsKey('Debug')) {
        $DebugPreference = $PSCmdlet.SessionState.PSVariable.GetValue('DebugPreference')
    }
    #get BIAS time
    [string]$LogTime = (Get-Date -Format 'HH:mm:ss.fff').ToString()
    [string]$LogDate = (Get-Date -Format 'MM-dd-yyyy').ToString()
    [int32]$script:LogTimeZoneBias = [timezone]::CurrentTimeZone.GetUtcOffset([datetime]::Now).TotalMinutes
    [string]$LogTimePlusBias = $LogTime + $script:LogTimeZoneBias

    #  Get the file name of the source script
    If($Source){
        $ScriptSource = $Source
    }
    Else{
        Try {
            If ($script:MyInvocation.Value.ScriptName) {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.Value.ScriptName -Leaf -ErrorAction 'Stop'
            }
            Else {
                [string]$ScriptSource = Split-Path -Path $script:MyInvocation.MyCommand.Definition -Leaf -ErrorAction 'Stop'
            }
        }
        Catch {
            $ScriptSource = ''
        }
    }

    #if the severity and preference level not set to silentlycontinue, then log the message
    $LogMsg = $true
    If( $Severity -eq 4 ){$Message='VERBOSE: ' + $Message;If(!$VerboseEnabled){$LogMsg = $false} }
    If( $Severity -eq 5 ){$Message='DEBUG: ' + $Message;If(!$DebugEnabled){$LogMsg = $false} }
    #If( ($Severity -eq 4) -and ($VerbosePreference -eq 'SilentlyContinue') ){$LogMsg = $false$Message='VERBOSE: ' + $Message}
    #If( ($Severity -eq 5) -and ($DebugPreference -eq 'SilentlyContinue') ){$LogMsg = $false;$Message='DEBUG: ' + $Message}

    #generate CMTrace log format
    $LogFormat = "<![LOG[$Message]LOG]!>" + "<time=`"$LogTimePlusBias`" " + "date=`"$LogDate`" " + "component=`"$ScriptSource`" " + "context=`"$([Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + "type=`"$Severity`" " + "thread=`"$PID`" " + "file=`"$ScriptSource`">"

    # Add value to log file
    If($LogMsg)
    {
        try {
            Out-File -InputObject $LogFormat -Append -NoClobber -Encoding Default -FilePath $OutputLogFile -ErrorAction Stop
        }
        catch {
            Write-Host ("[{0}] [{1}] :: Unable to append log entry to [{1}], error: {2}" -f $LogTimePlusBias,$ScriptSource,$OutputLogFile,$_.Exception.ErrorMessage) -ForegroundColor Red
        }
    }

    # Write to console
    If($VerbosePreference){
        Write-Verbose ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message)
    }
    Else{
        Write-Host ("[{0}] [{1}] :: {2}" -f $LogTimePlusBias,$ScriptSource,$Message) -ForegroundColor Cyan
    }
}

Function Get-LoggedOnUser{
    Param(
        [switch]$Passthru
    )

    $CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    #$UserSID = (Get-WmiObject Win32_UserAccount | Where-Object { $_.Name -eq ($LoggedOnUser -split '\\')[-1] }).SID
    if ($CurrentUser -eq "NT AUTHORITY\SYSTEM") {
        $LoggedOnUser = (Get-Process -IncludeUserName -Name explorer -ErrorAction SilentlyContinue).UserName
    }else {
        $LoggedOnUser = $env:USERNAME
    }
    Write-LogEntry -Message "Logged on user: $LoggedOnUser" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
    $UserSID = (New-Object System.Security.Principal.NTAccount($LoggedOnUser)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    Write-LogEntry -Message "Logged on user: $LoggedOnUser" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
    $UserProfilePath = (Get-CimInstance Win32_UserProfile | Where-Object { $_.SID -eq $UserSID }).LocalPath
    Write-LogEntry -Message "User Profile Path: $UserProfilePath" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose

    #create objkect of username and sid and profile path
    $UserObject = [PSCustomObject]@{
        UserName = ($LoggedOnUser -split '\\')[-1]
        SID = $UserSID
        ProfilePath = $UserProfilePath
    }
    If($Passthru){Return $UserObject}
    Else{Return $LoggedOnUser}
}




Function Get-UserModulePath{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $LoggedOnUser,
        [ValidateSet('WindowsPowerShell','PowerShell')]
        $PowershellPath = 'WindowsPowerShell'
    )
    Begin{
        $ModulePaths = @()
        #$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        Write-LogEntry -Message "Logged on user: $LoggedOnUser" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
    }
    Process{

        #get path of user modules
        New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -ErrorAction SilentlyContinue | Out-Null

        $ModulePaths += "C:\Users\$($LoggedOnUser.UserName)\Documents\$PowershellPath\Modules"
        #$env:PSModulePath = $env:PSModulePath +';' + $UserModulePath
        $OneDriveKey = "HKU:\$($LoggedOnUser.SID)\Software\Microsoft\OneDrive"
        if (Test-Path $OneDriveKey) {
            $OneDrivePath = Get-ChildItem -Path "$($LoggedOnUser.ProfilePath)\" -Directory | Where-Object { $_.Name -match "OneDrive - " } | Select-Object -ExpandProperty FullName
            $OneDriveModulePath = "$OneDrivePath\Documents\$PowershellPath\Modules"
            If(($env:PSModulePath -split ';') -contains $OneDriveModulePath){
                $ModulePaths += $OneDriveModulePath
            }Else{
                Write-LogEntry -Message "OneDrive Module Path not in env:PSModulePath: $OneDriveModulePath. Ignoring" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
                #$env:PSModulePath = $env:PSModulePath +';' + $ODModulePath
            }
        }
    }
    End{
        Write-LogEntry -Message "User Module Paths: $($ModulePaths -join ', ')" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
        return $ModulePaths
    }
}


Function Get-UserInstalledModule {
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string[]]$ModulePath,
        [switch]$GetDetails,
        [switch]$PassThru # New switch to return all versions
    )

    Begin {
        $InstalledModules = @()
    }
    Process {
        Foreach ($Path in $ModulePath) {
            If (Test-Path $Path -PathType Container) {
                Write-LogEntry -Message "Searching for modules in [$Path]" -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose

                # Get module names (first-level folders inside $ModulePath)
                $ModuleFolders = Get-ChildItem -Path $Path -Directory

                Foreach ($Module in $ModuleFolders) {
                    # Get available versions (subfolders inside the module folder)
                    $Versions = Get-ChildItem -Path $Module.FullName -Directory |
                                Where-Object { $_.Name -match '^\d+(\.\d+)*$' } # Ensure it's a version number

                    If ($Versions) {
                        # Sort versions numerically (latest first)
                        $SortedVersions = $Versions | Sort-Object { [version]$_.Name } -Descending

                        # Determine which versions to return
                        $SelectedVersions = If ($PassThru) { $SortedVersions } Else { $SortedVersions | Select-Object -First 1 }

                        Foreach ($Version in $SelectedVersions) {
                            $ModuleName = $Module.Name
                            $ModuleVersion = $Version.Name

                            Write-LogEntry -Message ("Processing module [{0}] Version [{1}]" -f $ModuleName, $ModuleVersion) -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose

                            # Skip duplicates (same module & version in different paths)
                            If ($InstalledModules | Where-Object { $_.Name -eq $ModuleName -and $_.Version -eq $ModuleVersion }) {
                                Write-LogEntry -Message ("Skipping duplicate module [{0}] Version [{1}]" -f $ModuleName, $ModuleVersion) -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
                                Continue
                            }

                            If ($GetDetails) {
                                Write-LogEntry -Message ("Getting details for module: [{0}] Version: [{1}]" -f $ModuleName, $ModuleVersion) -Source $MyInvocation.MyCommand.Name -Severity 1 -Verbose
                                $ModuleDetails = Find-Module -Name $ModuleName -RequiredVersion $ModuleVersion -Repository PSGallery -ErrorAction SilentlyContinue | Select-Object Version, Name, Repository, Description, Author, CompanyName, Dependencies

                                $InstalledModules += [PSCustomObject]@{
                                    Version = $ModuleVersion
                                    Name = $ModuleName
                                    Repository = $ModuleDetails.Repository
                                    Description = $ModuleDetails.Description
                                    Author = $ModuleDetails.Author
                                    CompanyName = $ModuleDetails.CompanyName
                                    Dependencies = $ModuleDetails.Dependencies
                                    InstalledLocation = $Path
                                }
                            } Else {
                                $InstalledModules += [PSCustomObject]@{
                                    Version = $ModuleVersion
                                    Name = $ModuleName
                                    InstalledLocation = ($Path + "\" + $ModuleName + "\" + $ModuleVersion)
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    End {
        return $InstalledModules
    }
}

Function Get-AllModules{
    $isAdmin = [bool](([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole] "Administrator"))

    $ModuleList = Get-Module -ListAvailable

    $ExludeModules = @(
        'PackageManagement'
        'Pester'
        'PowerShellGet'
        'PSReadLine'
    )

    #iif not an admin no need to get modules in all users:
    $ExludeAllUsersPaths = @(
        'C:\\Program Files\\WindowsPowerShell\\Modules'
        'C:\\Program Files\\PowerShell\\Modules'
        'C:\\Program Files\\PowerShell\\7\\Modules'
        'C:\\Program Files \(x86\)\\WindowsPowerShell\\Modules'
        'C:\\Program Files \(x86\)\\PowerShell\\Modules'
        'C:\\Program Files \(x86\)\\PowerShell\\7\\Modules'
        'C:\\Program Files\\PowerShell\\Modules'
    )

    $SystemPath = @(
        'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\Modules'
    )
    #always exclude system path
    $ModuleList = $ModuleList | Where-Object { $_.Path -notmatch ($SystemPath -join '|') }

    #exclude modules
    $ModuleList = $ModuleList | Where-Object { $_.Name -notmatch ($ExludeModules -join '|') }

    #exclude All Users paths if not admin
    If(-not $isAdmin){
        $ModuleList = $ModuleList | Where-Object { $_.Path -notmatch ($ExludeAllUsersPaths -join '|') }
    }

    return $ModuleList
}

Function Get-AdditionalDownloadsMapping{
    Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        $MappingConfig,
        $MappingId
    )
    $CheckboxMap = @{}
    # Loop through solution modules
    foreach ($Module in $MappingConfig) {
        $ModuleName = $Module.Name -replace '\W+', '' # Remove special char spaces from module name

        if ($Module.AdditionalDownloads) {
            foreach ($Download in $Module.AdditionalDownloads) {
                $DownloadName = $Download.Name -replace '\W+', '' # Remove special char and spaces from download name
                #$CheckboxID = "chkAddDownloads$ModuleName$DownloadName"
                $CheckboxID = "$ModuleName$DownloadName"

                # Store mapping
                $CheckboxMap[$CheckboxID] = @{
                    ModuleName = $Module.Name
                    DownloadName = $Download.Name
                    DownloadType = $Download.Type
                    DownloadUrl = $Download.DownloadUrl
                    DestinationPath = $Download.DestinationPath
                }
            }
        }
    }

    If($MappingId){
        $CheckboxMap[$MappingId]
    }Else{
        $CheckboxMap
    }
}


Function Test-IsISE {
    # try...catch accounts for:
    # Set-StrictMode -Version latest
    try {
        return ($null -ne $psISE);
    }
    catch {
        return $false;
    }
}
#endregion

#region FUNCTION: Check if running in Visual Studio Code
Function Test-VSCode{
    if($env:TERM_PROGRAM -eq 'vscode') {
        return $true;
    }
    Else{
        return $false;
    }
}
#endregion
#region FUNCTION: Get-PSDWizardScriptPath
Function Get-ScriptPath {
    <#
    .SYNOPSIS
        Finds the current script path even in ISE or VSC
    .LINK
        Test-InVSC
        Test-InISE
    #>
    param(
        [switch]$Parent
    )

    Begin {}
    Process {
        Try {
            if ($PSScriptRoot -eq "") {
                if (Test-IsISE) {
                    $ScriptPath = $psISE.CurrentFile.FullPath
                }
                elseif (Test-VSCode) {
                    $context = $psEditor.GetEditorContext()
                    $ScriptPath = $context.CurrentFile.Path
                }
                Else {
                    $ScriptPath = (Get-location).Path
                }
            }
            else {
                $ScriptPath = $PSCommandPath
            }
        }
        Catch {
            $ScriptPath = ($PWD.ProviderPath, $PSScriptRoot)[[bool]$PSScriptRoot]
        }
    }
    End {

        If ($Parent) {
            Split-Path $ScriptPath -Parent
        }
        Else {
            $ScriptPath
        }
    }

}
#endregion
##*=============================================
##* VARIABLES
##*=============================================
#[string]$scriptPath = ($PWD.ProviderPath, $PSScriptRoot)[[bool]$PSScriptRoot]

#build log name
[string]$scriptFullPath = Get-ScriptPath
If(Test-Path $scriptFullPath -PathType Leaf ){
    $scriptPath = Split-Path $scriptFullPath -Parent
    [string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptFullPath)
}else{
    $scriptPath = $scriptFullPath
    [string]$scriptName = 'PowerShellGalleryModuleInstaller'
}

$FileName = "$($scriptName)_$(Get-Date -Format 'yyyy-MM-dd_Thh-mm-ss-tt').log"
#build global log fullpath
If($LogFilePath){
    $LogFilePath = $LogFilePath
}else{
    $LogFilePath = Join-Path "$scriptPath\Logs" -ChildPath $FileName
}
Write-Host "logging to file: $LogFilePath" -ForegroundColor Cyan

$ConfigFilePath = "$scriptPath\UIConfig.json"
$UIConfig = Get-Content $ConfigFilePath | ConvertFrom-Json
Write-LogEntry -Message "Loading UI Config file: $ConfigFilePath" -Source $MyInvocation.MyCommand.Name -Severity 1

$CurrentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

#set the data stored path
If($StoredDataPath){
    New-Item -Path $StoredDataPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    Write-LogEntry -Message "Using stored data path: $StoredDataPath" -Source $MyInvocation.MyCommand.Name -Severity 1
    $ModuleDataPath = "$StoredDataPath\ExportedModuleData.xml"
    $SolutionDataPath = "$StoredDataPath\ExportedSolutionData.xml"
}else{
    Write-LogEntry -Message "Using default data path: $scriptPath" -Source $MyInvocation.MyCommand.Name -Severity 1
    $ModuleDataPath = "$scriptPath\ExportedModuleData.xml"
    $SolutionDataPath = "$scriptPath\ExportedSolutionData.xml"
}

#code can't be killed as it can have many processes running
If(Test-VSCode){
    Write-LogEntry -Message "Running in VS code, disabling process check" -Source $MyInvocation.MyCommand.Name -Severity 2
    $DisableProcessKill = $true
}else{
    $DisableProcessKill = $false
}

#build array
$TabItemData = @()
$TabControlXaml = @()
$AllModuleData = @()
$AllSolutionData = @()
##*=============================================
##* BUILD UI
##*=============================================
Write-Host ("Launching {0} [ver: {1}]..." -f $UIConfig.Title,$UIConfig.version ) -ForegroundColor Cyan
Write-Host "=========================================================================" -ForegroundColor Cyan
$totalSteps = 7
$buildSequence = Show-SequenceWindow -Config $UIConfig -Message "Building Selection UI, please wait..."
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Indeterminate -Message ("Load data for menu...")
Write-LogEntry -Message ("Loading UI, please wait...") -Source $MyInvocation.MyCommand.Name -Severity 0
Start-Sleep 3

Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 1 -MaxStep $totalSteps
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Indeterminate -Message ("Retrieving all installed modules on: {0}..." -f $env:Computername)
Write-LogEntry -Message ("Retrieving all installed modules on: {0}..." -f $env:Computername) -Source $MyInvocation.MyCommand.Name -Severity 0
#only get installed module:
$Global:InstalledModules = @()
#this will get all installed modules (if running under SYSTEM, only modules for SYSTEM will be included)
#$Global:InstalledModules += Get-InstalledModule | Select Name, Version, Description, Author, CompanyName, InstalledLocation
$Global:InstalledModules += Get-AllModules |
        Select Name, Version, Description, Author, CompanyName, @{Name='InstalledLocation';Expression={Split-Path $_.Path -Parent}}

#this will get all modules for the logged on user (even if running under SYSTEM)
$UserModulePath = Get-LoggedOnUser -Passthru | Get-UserModulePath | Get-UserInstalledModule
Foreach($UserModule in $UserModulePath){
    if($UserModule.Name -notin $Global:InstalledModules.Name){
        Write-LogEntry -Message ("Adding user module: {0}" -f $UserModule.Name) -Source $MyInvocation.MyCommand.Name -Severity 0
        $Global:InstalledModules += $UserModule
    }
}
#remove PowerShellGet and PackageManagement modules from list
$Global:InstalledModules = $Global:InstalledModules | Where-Object {$_.Name -NotMatch "PowerShellGet|PackageManagement"}
Write-LogEntry -Message ("Found [{0}] installed modules on: {1}" -f $Global:InstalledModules.Count,$env:Computername) -Source $MyInvocation.MyCommand.Name -Severity 0

#build UI for tab items

$i=0
#TEST $ModuleGroupItem = $UIConfig.ModuleGroups[0]
#TEST $ModuleGroupItem = $UIConfig.ModuleGroups[1]
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 2 -MaxStep $totalSteps
Foreach($ModuleGroupItem in $UIConfig.ModuleGroups){
    $i++
    Write-LogEntry -Message ("Building menu for module: {0}..." -f $ModuleGroupItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 1
    Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $UIConfig.ModuleGroups.Count -Message ("Building menu for module: {0}..." -f $ModuleGroupItem.Name)
    #if first make it selected
    $Selected = $false
    If($TabItemData.count -eq 0){ $Selected = $true}
    #build data for each tab item
    $TabItemData += Add-UIModuleTabItem -Selected:$Selected -ModuleConfig $ModuleGroupItem | Select Name, Type, Elements
    #build xaml for each tab item
    $TabControlXaml += Add-UIModuleTabItem -Selected:$Selected -ModuleConfig $ModuleGroupItem -Passthru
    #$TabControlXaml += Add-UIModuleTabItem -Selected:$Selected -ModuleConfig $ModuleGroupItem -CloseTabControl -Passthru
    Start-Sleep 1
}

#build UI for solution tab items
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 3 -MaxStep $totalSteps
$SolutionTabData = $UIConfig.SolutionGroupedModules | Add-UISolutionTabItem
$i=0
Foreach($SolutionTabItem in $SolutionTabData){
    $i++
    Write-LogEntry -Message ("Building menu for solution: {0}..." -f $SolutionTabItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 1
    Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $SolutionTabData.Count -Message ("Building menu for solution: {0}..." -f $SolutionTabItem.Name)
    Start-Sleep 1
}
#build data for each tab item
$TabItemData += $UIConfig.SolutionGroupedModules | Add-UISolutionTabItem | Select Name, Type, Elements
#build xaml for solution tab items
$TabControlXaml += $UIConfig.SolutionGroupedModules | Add-UISolutionTabItem -Passthru
Write-LogEntry -Message ("Built [{0}] menu items" -f $TabItemData.Count) -Source $MyInvocation.MyCommand.Name -Severity 1

# update module group; split beta from released and add elements
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 4 -MaxStep $totalSteps
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Indeterminate -Message ("Building Module Groups for search criteria...")
Write-LogEntry -Message ("Building Module Groups for search criteria...") -Source $MyInvocation.MyCommand.Name -Severity 1
$ModuleSearchData = $UIConfig.ModuleGroups | Merge-ModuleObject -MergeData ($TabItemData | Where Type -eq 'Module')

If(-not $PSBoundParameters.ContainsKey('SkipSolutionData')){
    Start-Sleep 1
    Write-LogEntry -Message ("Building Solution Groups for search criteria...") -Source $MyInvocation.MyCommand.Name -Severity 1
    Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Indeterminate -Message ("Building Solution Groups for search criteria...")
    $SolutionSearchData = $UIConfig.SolutionGroupedModules | Merge-SolutionObject -MergeData ($TabItemData | Where Type -eq 'Solution')
}

#=======================================================================================
# BUILD SEQUENCE module data
#=======================================================================================


#determine if we are using preloaded data
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 5 -MaxStep $totalSteps
If($PSBoundParameters.ContainsKey('ForceNewModuleData')){
    Write-LogEntry -Message ("Forcing new module data to be downloaded from PowerShell Gallery") -Source $MyInvocation.MyCommand.Name -Severity 1
    $UseExportedModuleData = $false
    $SourceLocation = 'PowerShell Gallery'
}Else{
    If(Test-Path $ModuleDataPath -ErrorAction SilentlyContinue){
        $UseExportedModuleData = $true
        $MockedModuleData = Import-Clixml $ModuleDataPath
        Write-LogEntry -Message ("Using preloaded module data: {0}" -f $ModuleDataPath) -Source $MyInvocation.MyCommand.Name -Severity 1
        $SourceLocation = 'ExportedModuleData.xml'
    }Else{
        $UseExportedModuleData = $false
        $SourceLocation = 'PowerShell Gallery'
    }
}

$i=0

#search for modules. This is the longest part of the process
Foreach($ModuleItem in $ModuleSearchData){
    $i++
    Write-LogEntry -Message ("Searching for module [{0}] using criteria [{1}] in {2}..." -f $ModuleItem.Name,$ModuleItem.ModuleSearch,$SourceLocation) -Source 'Get-ModuleListData' -Severity 1
    Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $ModuleSearchData.Count -Message ("Searching for module [{0}] in the PowerShell Gallery..." -f $ModuleItem.Name)
    If($UseExportedModuleData){
        $ModuleGroupItems = $MockedModuleData | Where-Object { $_.GroupName -eq $ModuleItem.Name }
        Start-Sleep 1
    }Else{
        $ModuleGroupItems = Get-ModuleListData -ModuleGroupItem $ModuleItem -Verbose
    }
    Write-LogEntry -Message ("Found [{0}] modules for module group [{1}]" -f $ModuleGroupItems.ModuleList.Count,$ModuleItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 1

    #COMBINE ALL MODULE DATA
    $AllModuleData += $ModuleGroupItems
}
#export module data for next time
$AllModuleData | Export-Clixml -Path $ModuleDataPath -Force


# populate solution data
#=======================================================================================
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 6 -MaxStep $totalSteps

#determine if we are using preloaded data
If($PSBoundParameters.ContainsKey('ForceNewSolutionData')){
    Write-LogEntry -Message ("Forcing new solution data to be downloaded from PowerShell Gallery") -Source $MyInvocation.MyCommand.Name -Severity 1
    $UseExportedSolutionData = $false
    $SourceLocation = 'PowerShell Gallery'
}Else{
    If(Test-Path $SolutionDataPath -ErrorAction SilentlyContinue){
        $UseExportedSolutionData = $true
        $MockedSolutionData = Import-Clixml $SolutionDataPath
        Write-LogEntry -Message ("Using preloaded solution data: {0}" -f $SolutionDataPath) -Source $MyInvocation.MyCommand.Name -Severity 1
        $SourceLocation = 'ExportedSolutionData.xml'
    }Else{
        $UseExportedSolutionData = $false
        $SourceLocation = 'PowerShell Gallery'
    }
}


If(-not $PSBoundParameters.ContainsKey('SkipSolutionData')){
    $i=0
    #TEST $SolutionItem = $SolutionSearchData[0]
    #TEST $SolutionItem = $SolutionSearchData[1]
    Foreach($SolutionItem in $SolutionSearchData){
        $i++
        Write-LogEntry -Message ("Searching for module list of module for solution group [{0}] in {1}..." -f $SolutionItem.Name, $SourceLocation) -Source $MyInvocation.MyCommand.Name -Severity 1
        Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $SolutionSearchData.Count -Message ("Searching for module list for solution group [{0}]..." -f $SolutionItem.Name)
        If($UseExportedSolutionData){
            $SolutionGroupItems = $MockedSolutionData | Where-Object { $_.GroupName -eq $SolutionItem.Name }
            Start-Sleep 1
        }Else{
            $SolutionGroupItems = $SolutionItem | Get-SolutionListData -Verbose
        }
        Write-LogEntry -Message ("Found [{0}] modules for solution group [{1}]" -f $SolutionGroupItems.ModuleList.Count,$SolutionItem.Name) -Source $MyInvocation.MyCommand.Name -Severity 1
        #COMBINE ALL SOLUTION DATA
        $AllSolutionData += $SolutionGroupItems
    }
    #export solution data
    $AllSolutionData | Export-Clixml -Path $SolutionDataPath -Force
}


Write-LogEntry -Message ("Found [{0}] modules and [{1}] solutions" -f $AllModuleData.Count,$AllSolutionData.Count) -Source $MyInvocation.MyCommand.Name -Severity 1
Update-SequenceProgressBar -Runspace $buildSequence -ProgressBar 'ProgressBarMain' -Step 7 -MaxStep $totalSteps -Message ("Launching {0} UI..." -f $UIConfig.Title)
Write-LogEntry -Message ("Launching {0} UI..." -f $UIConfig.Title) -Source $MyInvocation.MyCommand.Name -Severity 1

Start-Sleep 3
Close-SequenceWindow -Runspace $buildSequence

##=============================================
## LAUNCH SELECTOR UI
##=============================================
Write-Host "Loading UI, please wait..." -ForegroundColor White
$Global:UI = Show-UIMainWindow `
                -TabContents $TabControlXaml `
                -TabElements $TabItemData `
                -Config $UIConfig `
                -ModuleData $AllModuleData `
                -SolutionData $AllSolutionData `
                -InstalledModules $Global:InstalledModules `
                -LogPath ($LogFilePath -replace "\.log$","_UI.log") `
                -TopPosition $buildSequence.Window.Top `
                -LeftPosition $buildSequence.Window.Left `
                -DisableProcessCheck:$DisableProcessKill `
                -Wait


# after the UI is closed, check for outputdata to do work
If($Global:UI.OutputData.DoAction -eq $False){
    Write-LogEntry -Message "No action selected. Exiting" -Source $MyInvocation.MyCommand.Name -Severity 2
    If($Global:UI.Error){
        Write-LogEntry -Message ("UI HAS {0} ERRORS" -f $Global:UI.Error.count) -Source $MyInvocation.MyCommand.Name -Severity 3
        Write-Host "UI ERRORS:" -ForegroundColor Red
        Write-Host "==================================================================" -ForegroundColor Red
        $Global:UI.Error
    }
    Exit
}ElseIf($Global:UI.OutputData.UseExternalInstaller -and -not $SimulateInstall){
    #use external installer
    Write-LogEntry = "Using external installer..." -Source $MyInvocation.MyCommand.Name -Severity 1
    Return $Global:UI.OutputData
}Else{
    Write-LogEntry -Message "Starting module install sequence..." -Source $MyInvocation.MyCommand.Name -Severity 1
    #launch Sequence Window
    $Global:installSequence = Show-SequenceWindow -Config $UIConfig -Message "Working on modules, this can take some time..." -TopPosition $Global:UI.Window.Top -LeftPosition $Global:UI.Window.Left
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Indeterminate -Message "Starting module install sequence..."
}
<#
INSTALL SEQUENCE
- For each item add to the outputData object; this will reopen the Sequence Window or pass the data to the external installer
    - if UseExternalInstaller is checked, use external installer (only OutputData for external installer)
    - If "Remove All Except selected" is checked, remove all modules from all locations and install selected modules
        Do i need to remove all modules from all locations? or just use remove-module and uninstall-module?
        - C:\Program Files\WindowsPowerShell\Modules
        - C:\Program Files(x86)\WindowsPowerShell\Modules
        - C:\Program Files\PowerShell\Modules
        - C:\Program Files(x86)\PowerShell\Modules
        - C:\Users\CurrentUser\Documents\WindowsPowerShell\Modules
        - C:\Windows\System32\config\systemprofile\Documents\WindowsPowerShell\Modules
    - If "Auto Update" is checked, update all modules that have updates available
    - If "Repair" is checked, reinstall all modules that are equal versions
    - If "Remove Duplicate" is checked, remove the oldest duplicate modules
    - If "Install for PS7" is checked, install modules in PS7 module directory (should it open a pwsh.exe window to install or move module to PS7 location?)
        - Module section in uiconfig has "PowerShellVersionSupport". 
            - If "Install for PS7" is selected; only install those that support it, install the rest for 5.1
            - If "Install for PS7" is not selected and powershell support 5.1, install those for 5.1
            - If "Install for PS7" is not selected and powershell only support 7, install only the 5.1 support
    - If "Install under user context" is checked, install modules under user context
        - If running as SYSTEM, install under SYSTEM context, then move to user location and update ACL
        - If running as user, install under user context
        - otherwise install module using -scope AllUsers
    - Install selected modules from $Global:UI.OutputData.SelectedModules
    - if additional downloads are checked required, add to OutputData
        -Config will be used as instruction to download and install files


OUTPUT DATA EXAMPLE:
#================================================================================================
$Global:UI.OutputData = @{
    DoAction = $true
    RemoveAll = $syncHash.chkModuleRemoveAll.IsChecked
    AutoUpdate = $syncHash.chkModuleAutoUpdate.IsChecked
    DuplicateCleanup = $syncHash.chkModuleRemoveDuplicates.IsChecked
    RepairSelected = $syncHash.chkModuleRepairSelected.IsChecked
    PS7install = $syncHash.chkModuleInstallForPS7.IsChecked
    InstallUserContext = $syncHash.chkModuleInstallUserContext.IsChecked
    SelectedModules = $syncHash.lbSelectedModules.Items
    InstalledModules = $syncHash.InstalledModules
    AdditionalDownloads = $syncHash.AdditionalDownloads | Select -Unique
}

#>
If($SimulateInstall){$WhatIfPreference = $true}
If($Global:UI.OutputData.PS7install){
    $PowerShellFolderPath = 'PowerShell'
}else{
    $PowerShellFolderPath = 'WindowsPowerShell'
}

If($Global:UI.OutputData.InstallUserContext){
    $InstallContext = 'CurrentUser'
}Else{
    $InstallContext = 'AllUsers'
}

#Lets get all checked itemd to determine max steps
$startStep = 0
$TotalSteps = 0
If($Global:UI.OutputData.DoAction){$TotalSteps++}
If($Global:UI.OutputData.RemoveAll){$TotalSteps++}
If($Global:UI.OutputData.AutoUpdate){$TotalSteps++}
If($Global:UI.OutputData.DuplicateCleanup){$TotalSteps++}
If($Global:UI.OutputData.RepairSelected){$TotalSteps++}
If($Global:UI.OutputData.AdditionalDownloads){$TotalSteps++}

#you can't have a 1 becuse any step plus DoAction will be 2
If($TotalSteps -le 1){
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Indeterminate -Message "No action selected. Exiting"
    Write-LogEntry -Message "No action selected. Exiting" -Source $MyInvocation.MyCommand.Name -Severity 2
    Start-Sleep 3
    Close-SequenceWindow -Runspace $installSequence
    Exit 0
}


Write-LogEntry -Message ("Starting module install sequence...") -Source $MyInvocation.MyCommand.Name -Severity 1

Start-Sleep 2
##================================================================================================
## INSTALL SEQUENCE - REPAIR SELECTED MODULES
##================================================================================================
If($Global:UI.OutputData.RepairSelected)
{
    $startStep++
    Write-LogEntry -Message ("[{0} of {1}] Repairing selected modules" -f $startStep,$totalSteps) -Source 'Installer' -Severity 0
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    #Repair selected modules
    $i=0
    foreach($RepairModule in $Global:UI.OutputData.SelectedModules)
    {
        $i++
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $Global:UI.OutputData.SelectedModules.Count -Message "Repairing module [$($RepairModule.Name)]"
        Write-LogEntry -Message ("Repairing module [{0}]" -f $RepairModule.Name) -Source 'Installer' -Severity 0
        try {
            Get-Module -Name $RepairModule.Name -ListAvailable | Remove-Module -Force -ErrorAction Stop
            Get-Module -Name $RepairModule.Name -ListAvailable | Uninstall-Module -Force -ErrorAction Stop
            Install-Module -Name $RepairModule.Name -Scope $InstallContext -Force -ErrorAction Stop
        }
        catch {
            Write-LogEntry -Message ("Failed to repair module [{0}]: {1}" -f $RepairModule.Name,$_.Exception.Message) -Source 'Installer' -Severity 3
        }
    }
}

##================================================================================================
## INSTALL SEQUENCE - REMOVE ALL MODULES EXCEPT SELECTED
##================================================================================================
$RemovedModules = @()
If($Global:UI.OutputData.RemoveAll)
{
    $startStep++
    Write-LogEntry -Message ("[{0} of {1}] Removing all modules (except selected)" -f $startStep,$totalSteps) -Source 'Installer' -Severity 0
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    #remove all modules
    $i=0
    $maxCount = $Global:UI.InstalledModules.count

    Foreach($RemoveModule in $Global:UI.InstalledModules)
    {
        $i++
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $maxCount -Message "Removing module [$($RemoveModule.Name)]"
        If($Global:UI.OutputData.SelectedModules -contains $RemoveModule.Name){
            Write-LogEntry -Message ("Module [{0}] was selected, skipping removal" -f $RemoveModule.Name) -Source 'Installer' -Severity 0
        }Else{
            Write-LogEntry -Message ("Removing module [{0}]" -f $RemoveModule.Name) -Source 'btnInstall' -Severity 0
            try {
                Get-Module -Name $RemoveModule.Name -ListAvailable | Remove-Module -Force -ErrorAction Stop
                Get-Module -Name $RemoveModule.Name -ListAvailable | Uninstall-Module -Force -ErrorAction Stop
                $RemovedModules += $RemoveModule.Name
            }
            catch {
                Write-LogEntry -Message ("Failed to remove module [{0}]: {1}" -f $RemoveModule.Name,$_.Exception.Message) -Source 'Installer' -Severity 3
            }
        }
    }
}

##================================================================================================
## INSTALL SEQUENCE - AUTO UPDATE MODULES
##================================================================================================
If($Global:UI.OutputData.AutoUpdate)
{
    $startStep++
    #update all modules
    Write-LogEntry -Message ("[{0} of {1}] Updating all modules" -f $startStep,$totalSteps) -Source 'Installer' -Severity 0
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    #TODO: Update all modules
    $i=0
    Foreach($ModuleUpdate in $Global:UI.InstalledModules)
    {
        $i++
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $Global:UI.InstalledModules.Count -Message "Updating module [$($ModuleUpdate.Name)]"

        try {
            #install module if it was removed
            If( ($RemovedModules | Select -Unique) -contains $ModuleUpdate.Name){
                Write-LogEntry -Message ("Module was recently removed, installing module [{0}]" -f $ModuleUpdate.Name) -Source 'Installer' -Severity 0
                Install-Module -Name $ModuleUpdate.Name -Scope $InstallContext -Force -ErrorAction Stop
            }Else{
                Write-LogEntry -Message ("Updating module [{0}]" -f $ModuleUpdate.Name) -Source 'Installer' -Severity 0
                Update-Module -Name $ModuleUpdate.Name -Scope $InstallContext -Force -ErrorAction Stop
            }
        }
        catch {
            Write-LogEntry -Message ("Failed to update module [{0}]: {1}" -f $ModuleUpdate.Name,$_.Exception.Message) -Source 'Installer' -Severity 3
        }
    }
}

##================================================================================================
## INSTALL SEQUENCE - REMOVE DUPLICATE MODULES
##================================================================================================
If($Global:UI.OutputData.DuplicateCleanup)
{
    $startStep++
    #Remove duplicate modules
    Write-LogEntry -Message ("[{0} of {1}] Cleaning up old modules" -f $startStep,$totalSteps) -Source 'Installer' -Severity 0
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    #TODO: Remove duplicate modules. Check if each list has update available are any old versions of modules and remove them.
    $i=0
    Foreach($ModuleItem in $Global:UI.InstalledModules)
    {
        $i++
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxStep $Global:UI.InstalledModules.Count -Message "Cleaning up module [$($ModuleItem.Name)]"
        Write-LogEntry -Message ("Cleaning up module [{0}]" -f $ModuleItem.Name) -Source 'Installer' -Severity 0

        #get all versions of module
        $ModuleVersions = $Global:UI.InstalledModules | Where Name -eq $ModuleItem.Name
        If($ModuleVersions.count -gt 1){
            try {
                #get latest version
                $LatestVersion = $ModuleVersions | Sort-Object -Descending | Select-Object -First 1
                #remove all versions except latest
                Foreach($VersionItem in $ModuleVersions | Where-Object { $ModuleItem -ne $LatestVersion })
                {
                    Write-LogEntry -Message ("Cleaning up duplicate module [{0}] version [{1}]" -f $ModuleItem.Name,$VersionItem.Version) -Source 'Installer' -Severity 0
                    Uninstall-Module -Name $ModuleItem.Name -RequiredVersion $VersionItem.Version -Force -ErrorAction Stop
                }
            }
            catch {
                Write-LogEntry -Message ("Failed to clean up module [{0}]: {1}" -f $ModuleItem.Name,$_.Exception.Message) -Source 'Installer' -Severity 3
            }
        }Else{
            Write-LogEntry -Message ("Module [{0}] has only one version, skipping" -f $ModuleItem.Name) -Source 'Installer' -Severity 0
        }
    }
}

##================================================================================================
## INSTALL SEQUENCE - INSTALL SELECTED MODULES
##================================================================================================
If($Global:UI.OutputData.SelectedModules.count -gt 0)
{
    $startStep++
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    Write-LogEntry -Message ("Installing {0} selected modules" -f $Global:UI.OutputData.SelectedModules.count) -Source 'Installer' -Severity 0
    #TODO: Install selected modules
    $i=0
    Foreach($SelectedModule in $Global:UI.OutputData.SelectedModules)
    {
        $i++
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxSteps $Global:UI.OutputData.SelectedModules.Count -Message ("Installing module [{0}]" -f $SelectedModule)
        Write-LogEntry -Message ("Installing module [{0}]" -f $SelectedModule) -Source 'Installer' -Severity 0

        #install under user context
        If($Global:UI.OutputData.InstallUserContext -and ($CurrentUser -eq "NT AUTHORITY\SYSTEM"))
        {

            Write-LogEntry -Message "Installing under user context while running under SYSTEM..." -Source 'Installer' -Severity 0
            <#
            Retrieve current user logged in or last user logged in
            when running as SYSTEM; install-Module -Scope CurrentUser points to C:\Windows\System32\config\systemprofile\Documents\WindowsPowerShell\Modules
            which is not the correct location for the user. To fix this:
                - Install module under System
                - move module to user location
                - update ACL for module
            #>
            Try{
                Install-Module -Name $SelectedModule -Force -Scope CurrentUser -ErrorAction Stop
                #GEt user SID
                $UserInfo = Get-LoggedOnUser -Passthru
                #$OneDrivePath = Get-ItemProperty -Path "HKCU:\Software\Microsoft\OneDrive\Accounts\Business1" -Name "UserFolder" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty UserFolder
                $Destination = $UserInfo | Get-UserModulePath -PowershellPath $PowerShellFolderPath | Where-Object { $SelectedModule -match "OneDrive" } | Select-Object -ExpandProperty FullName
                If(-Not $Destination){
                    Write-LogEntry -Message "OneDrive path not found" -Source 'Installer' -Severity 2
                    $Destination = $UserInfo | Get-UserModulePath -PowershellPath $PowerShellFolderPath | Select -First 1
                }
                #$Source = (Get-InstalledModule Az.Accounts | Select -ExpandProperty InstalledLocation) -replace '\d+\.\d+\.\d+'
                $Source = "C:\WINDOWS\system32\config\systemprofile\Documents\$PowerShellFolderPath\Modules\$($SelectedModule)"

                If(-Not (Test-Path -LiteralPath $Destination)){
                    New-Item -ItemType Directory -LiteralPath $Destination -Force
                }
                Move-Item -Path $Source -Destination $Destination -Force
                #update ACL
                $Acl = Get-Acl -LiteralPath "$Destination\$($SelectedModule)"
                $Account = New-Object -TypeName System.Security.Principal.NTAccount -ArgumentList $userInfo.UserName
                $Acl.SetOwner($Account); # Update the in-memory ACL
                Set-Acl -Path $Destination -AclObject $Acl
            }
            catch {
                Write-LogEntry -Message ("Failed to install module [{0}]: {1}" -f $SelectedModule,$_.Exception.Message) -Source 'Installer' -Severity 3
            }

        }
        ElseIf($Global:UI.OutputData.PS7install){
            Write-LogEntry -Message "Installing for PowerShell 7 under context [$InstallContext]" -Source 'Installer' -Severity 0
            Try{
                & pwsh -NoProfile -Command "Install-Module -Name $SelectedModule -Scope $InstallContext -Force"
            }
            catch {
                Write-LogEntry -Message ("Failed to install module [{0}] for PowerShell 7: {1}" -f $SelectedModule,$_.Exception.Message) -Source 'Installer' -Severity 3
            }
        }Else{
            Write-LogEntry -Message "Installing for Windows PowerShell 5.1 under context [$InstallContext]" -Source 'Installer' -Severity 0
            Try{
                Install-Module -Name $SelectedModule -Force -Scope $InstallContext -ErrorAction Stop
            }
            catch {
                Write-LogEntry -Message ("Failed to install module [{0}] for Windows PowerShell 5.1: {1}" -f $SelectedModule,$_.Exception.Message) -Source 'Installer' -Severity 3
            }
        }
    }
}

##================================================================================================
## INSTALL SEQUENCE - ADDITIONAL DOWNLOADS
##================================================================================================
If($Global:UI.OutputData.AdditionalDownloads)
{
    $startStep++
    #additional downloads
    Write-LogEntry -Message ("[{0} of {1}] Processing additional downloads" -f $startStep,$totalSteps) -Source 'Installer' -Severity 0
    Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -Step $startStep -MaxStep $totalSteps
    Foreach($DownloadID in $Global:UI.OutputData.AdditionalDownloads)
    {
        $i++
        #get solution but only where additional download is avaialble
        $DownloadItem = Get-AdditionalDownloadsMapping -MappingConfig $UIConfig.SolutionGroupedModules -MappingId $DownloadID
        If(-Not $DownloadItem){
            Write-LogEntry -Message ("Download item [{0}] not found" -f $DownloadID) -Source 'Installer' -Severity 3
            Continue
        }
        Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -Step $i -MaxSteps $Global:UI.OutputData.AdditionalDownloads.count -Message ("Installing [{0}]..." -f $DownloadItem.DownloadName)
        Write-LogEntry -Message ("Downloading [{0}] from [{1}]" -f $DownloadItem.DownloadName,$DownloadItem.DownloadUrl) -Source 'Installer' -Severity 0
        #DO ACTION

        #build destination path
        #destination path has "$env:Temp\\Downloads", need to replace any environment variables
       # $DestinationPath = $DownloadItem.DestinationPath -replace '\$env:(\w+)\\', ([System.Environment]::GetEnvironmentVariable($matches[1]) + '\')
        $DestinationPath = $ExecutionContext.InvokeCommand.ExpandString($DownloadItem.DestinationPath)
        $FileName = Split-Path $DownloadItem.DownloadUrl -Leaf
        New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

        #download file
        $DownloadPath = Join-Path -Path $env:temp -ChildPath $FileName
        Invoke-WebRequest -Uri $DownloadItem.DownloadUrl -OutFile $DownloadPath -ErrorAction Stop

        #install File
        Write-LogEntry -Message ("Installing [{0}]]" -f $DownloadItem.DownloadName,$DownloadPath) -Source 'Installer' -Severity 0
        switch($DownloadItem.DownloadType){
            'MSI'{
                Write-LogEntry -Message ("RUNNING MSI [msiexec /i `"{0}`" /qn /norestart]" -f $DownloadPath) -Source 'Installer' -Severity 0 -Verbose
                Start-Process -FilePath msiexec -ArgumentList "/i `"$DownloadPath`" /qn /norestart" -Wait
            }
            'Executable'{
                Write-LogEntry -Message ("RUNNING EXECUTABLE [{0} /quiet]" -f $DownloadPath) -Source 'Installer' -Severity 0 -Verbose
                Start-Process -FilePath $DownloadPath -ArgumentList "/quiet" -Wait
            }
            'Zip'{
                Write-LogEntry -Message ("Extracting ZIP [{0}] to [{1}]" -f $DownloadPath,$DestinationPath) -Source 'Installer' -Severity 0 -Verbose
                Expand-Archive -Path $DownloadPath -DestinationPath $DestinationPath -Force
            }
            'MSIX'{
                Write-LogEntry -Message ("RUNNING MSIX [Add-AppxPackage `"{0}`"]" -f $DownloadPath) -Source 'Installer' -Severity 0 -Verbose
                Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command Add-AppxPackage `"$DownloadPath`"" -Wait
            }
            'File'{
                Write-LogEntry -Message ("Copying file [{0}] to [{1}]" -f $FileName,$DestinationPath) -Source 'Installer' -Severity 0 -Verbose
                New-Item -Path $DestinationPath -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                Copy-Item -Path $DownloadPath -Destination $DestinationPath -Force
            }
            default{
                Write-LogEntry -Message ("Unknown file type [{0}] for [{1}]" -f $DownloadItem.DownloadType,$DownloadItem.DownloadName) -Source 'Installer' -Severity 3
            }
        }
    }
}

If($SimulateInstall){
    $WhatIfPreference = $false
}
Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarMain' -PercentComplete 100 -Message "Install steps complete, Exiting..."
Update-SequenceProgressBar -Runspace $installSequence -ProgressBar 'ProgressBarSub' -PercentComplete 100 -Message "Install steps complete, Exiting..."
Write-LogEntry -Message "Installation complete" -Source $MyInvocation.MyCommand.Name -Severity 1
Start-Sleep 5
Close-SequenceWindow -Runspace $installSequence

#Tag for detection of completion
#=======================================================================================
If($TagDetectionPath){
    $TagName = $UIconfig.Title -replace '\W+',''
    $TagVersion = $UIconfig.Version
    $TagFullPath = "$TagDetectionPath\$TagName\$TagName.$TagVersion.tag"
    New-Item -Path "$TagDetectionPath\$TagName" -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
    #remove old tag
    Get-ChildItem -Path $TagFullPath -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    Set-Content -Path $TagFullPath -Value $Global:UI.OutputData.SelectedModules -Force
    Write-LogEntry -Message ("Created tagged file [{0}]" -f $TagFullPath) -Source $MyInvocation.MyCommand.Name -Severity 1
}
Write-Host "==================================================================" -ForegroundColor Cyan