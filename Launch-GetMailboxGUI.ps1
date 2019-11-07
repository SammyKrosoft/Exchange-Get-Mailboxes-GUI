
<#PSScriptInfo

.VERSION 1.1.1

.GUID b5592ee1-dc40-4e2e-aa14-a819b595ceb4

.AUTHOR sammy

.DESCRIPTION This is a script to demonstrate the usage of Graphical User Interface (GUI) with PowerShell. Just launch the script, and click on the buttons to get your mailbox information, whether you're On-Premises, on a Hybrid mode, or fully on the Cloud.

.NOTES
    This script requires:
        - PowerShell v3 minium
        - Exchange Management Tools (or being logged on Exchange Online on the current PowerShell session)
    
    These above dependencies are checked when the script launches.

.LINK
 https://github.com/SammyKrosoft/Exchange-Get-Mailboxes-GUI

.COMPANYNAME SCO - Sam Corp Ottawa

.COPYRIGHT Free to copy, inspire, etc...

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/SammyKrosoft/Exchange-Get-Mailboxes-GUI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#>


$Version = "1.1.1"
<#Version History
v1.1.1 - Fixed RecipientTypeDetails : forgot to select REcipientTypeDetails on the main  Get-Mailbox
v1.1.0 - Fixed test whether mailbox is OnPrem or on the Cloud : forgot to select OrganizationalUnit on the main Get-Mailbox
v1.0.5 - tested different comments for PSScriptInfo
v1.00 - added PSSCriptInfo for publishing on PSGallery
v0.97 - made window bigger - added Copy to Clipboard button for mailboxes list - added ServerName and database info
v0.96 - added icon, made window a bit bigger
v0.95 - fixed Arbitration mailbox SIR and Quota functions
v0.92 - Added Arbitration mailbox check box
v0.91 - Added ability to sort each columns in List quota action ...
#>
#region FUNCTIONS other than Form events
Function IsPSV3 {
    <#
    .DESCRIPTION
    Just printing Powershell version and returning "true" if powershell version
    is Powershell v3 or more recent, and "false" if it's version 2.
    .OUTPUTS
    Returns $true or $false
    .EXAMPLE
    IsPSVersionV3
    #>
    $PowerShellMajorVersion = $PSVersionTable.PSVersion.Major
    $msgPowershellMajorVersion = "You're running Powershell v$PowerShellMajorVersion"
    Write-Host $msgPowershellMajorVersion -BackgroundColor blue -ForegroundColor yellow
    If($PowerShellMajorVersion -le 2){
        Write-Host "Sorry, PowerShell v3 or more is required. Exiting."
        Return $false
        Exit
    } Else {
        Write-Host "You have PowerShell v3 or later, great !" -BackgroundColor blue -ForegroundColor yellow
        Return $true
        }
}

Function Test-ExchTools(){
    <#
    .SYNOPSIS
    This small function will just check if you have Exchange tools installed or available on the
    current PowerShell session.

    .DESCRIPTION
    The presence of Exchange tools are checked by trying to execute "Get-ExBanner", one of the basic Exchange
    cmdlets that runs when the Exchange Management Shell is called.

    Just use Test-ExchTools in your script to make the script exit if not launched from an Exchange
    tools PowerShell session...

    .EXAMPLE
    Test-ExchTools
    => will exit the script/program si Exchange tools are not installed
    #>
    Try
    {
        Get-command Get-MAilbox -ErrorAction Stop
        $ExchInstalledStatus = $true
        $Message = "Exchange tools are present !"
        Write-Host $Message -ForegroundColor Blue -BackgroundColor Red
    }
    Catch [System.SystemException]
    {
        $ExchInstalledStatus = $false
        $Message = "Exchange Tools are not present ! This script/tool need these. Exiting..."
        Write-Host $Message -ForegroundColor red -BackgroundColor Blue
        # Add-Type -AssemblyName presentationframework, presentationcore
        # Option #4 - a message, a title, buttons, and an icon
        # More info : https://msdn.microsoft.com/en-us/library/system.windows.messageboximage.aspx
        $msg = "You must run this tool from an Exchange-enabled PowerShell console like Exchange Management Console or a PowerShell session where you imported an Exchange session."
        $Title = "Error - No Exchange Tools available !"
        $Button = "Ok"
        $Icon = "Error"
        [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
        Exit
    }
    Return $ExchInstalledStatus
}

Function Run-Action{
    $SelectedAction = $wpf.comboSelectAction.SelectedItem.Content
    Switch ($SelectedAction) {
        #Region Disable Mailbox ***********************************************************
        "Disable Mailbox"  {
            Write-host "Displaying Info"
            Write-Host "Listing selected mailbox names:"
            $SelectedITems = $wpf.GridView.SelectedItems
            $List = @()
            $SelectedItems | Foreach{
                $List += ("""") + $($_.Alias) + ("""")
            }
            $List = $List -join ","
            $Command = "$List | Disable-Mailbox"
            WRite-Host "About to execute action on $($SelectedItems.Count) mailboxes..."
            Write-Host "About to run $Command"
        }
        #endregion
        #End of the Disabled Mailbox region
        #region List Quotas
        "List Quotas"   {
            Write-host "Displaying Mailbox SIR and retention settings status"
            $SelectedITems = $wpf.GridView.SelectedItems
            Write-host "Displaying Mailbox Single Item Recovery and retention settings status for $($SelectedItems.count) items..."
            $List = @()
            $SelectedItems | Foreach {
                $List += $_.primarySMTPAddress.tostring()
            }
            #$List = $List -join ","
            Function Get-MailboxQuotas {
                [CmdLetBinding()]
                Param(
                    [Parameter(Mandatory = $False, Position = 1)][string[]]$List,
                    [Parameter(Mandatory = $False, Position = 2)][switch]$Arbitration
                )
                #Initiating stopwatch to measure the time it takes to retrieve mailboxes
                $stopwatch = [system.diagnostics.stopwatch]::StartNew()
                If ($Arbitration){
                    $QueryMailboxFeaturesStd = $List | get-mailbox -Arbitration | Select DisplayName,Name,ServerName,Database,*quota*,OrganizationalUnit
                } Else {
                    $QueryMailboxFeaturesStd = $List | get-mailbox | Select DisplayName,Name,ServerName,Database,*quota*,OrganizationalUnit
                }
                
                $QueryMailboxFeatures = @()
                Foreach ($mailbox in $QueryMailboxFeaturesStd){
                    $objTemp = $mailbox
                    if ($($objTemp.OrganizationalUnit) -match "prod.outlook.com"){
                        $CloudMailbox = $True
                    }else{
                        $CloudMailbox = $false
                    }
                    #<optional> - Removing OrganizationalUnit information ...
                    $ObjTemp = $ObjTemp | Select DisplayName, Database, UseDatabaseQuotaDefaults, ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, RulesQuota, CalendarLoggingQuota, RecoverableItemsQuota, RecoverableItemsWarningQuota, ArchiveQuota, ArchiveWarningQuota
                    if ($CloudMailbox){
                        $objTemp | add-Member -NotePropertyName DatabaseProhibitSRQuota -NotePropertyValue "Cloud mailbox (no DB Quota info)"
                        $ObjTemp | Add-Member -NotePropertyName DatabaseSendQuota -NotePropertyValue "Cloud mailbox (no DB Quota info)"
                        $ObjTemp | Add-Member -NotePropertyName DatabaseWarningQuota -NotePropertyValue "Cloud mailbox (no DB Quota info)"
                    } Else {
                        $objTemp | add-Member -NotePropertyName DatabaseProhibitSRQuota -NotePropertyValue $((Get-MailboxDatabase $($Mailbox.Database)).ProhibitSendReceiveQuota)
                        $ObjTemp | Add-Member -NotePropertyName DatabaseSendQuota -NotePropertyValue $((Get-MailboxDatabase $($Mailbox.Database)).ProhibitSendQuota)
                        $ObjTemp | Add-Member -NotePropertyName DatabaseWarningQuota -NotePropertyValue $((Get-MailboxDatabase $($Mailbox.Database)).IssueWarningQuota)
                    }
                    $QueryMailboxFeatures += $objTemp
                }
                [System.Collections.IENumerable]$MailboxFeatures = @($QueryMailboxFeatures)
                Write-host $($MailboxFeatures | ft | out-string)
                
                #Stopping stopwatch
                $stopwatch.Stop()
                $msg = "`n`nInstruction took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds ..."
                Write-Host $msg
                $msg = $null
                $StopWatch = $null

                #region Get-MailboxFeaturesView Form definition
                # Load a WPF GUI from a XAML file build with Visual Studio
                Add-Type -AssemblyName presentationframework, presentationcore
                $wpf = @{ }
                # NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
                #$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
                $inputXML = @"
                <Window x:Name="frmMbxQuotaStatus" x:Class="Get_CASMAilboxFeaturesWPF.MainWindow"
                                        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                                        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                        xmlns:local="clr-namespace:Get_CASMAilboxFeaturesWPF"
                                        mc:Ignorable="d"
                                        Title="Mailboxes Quotas settings status" Height="450" Width="872.145" ResizeMode="NoResize">
                    <Grid>
                        <DataGrid x:Name="DataGrid" HorizontalAlignment="Left" Height="326" Margin="10,10,-59,0" VerticalAlignment="Top" Width="844" IsReadOnly="True"/>
                        <Button x:Name="btnClose" Content="Close" HorizontalAlignment="Left" Margin="748,352,0,0" VerticalAlignment="Top" Width="106" Height="46"/>
                        <Button x:Name="btnClipboard" Content="Copy to clipboard" HorizontalAlignment="Left" Margin="10,352,0,0" VerticalAlignment="Top" Width="174" Height="46"/>

                    </Grid>
                </Window>   
"@

                $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
                [xml]$xaml = $inputXMLClean
                $reader = New-Object System.Xml.XmlNodeReader $xaml
                $tempform = [Windows.Markup.XamlReader]::Load($reader)
                $namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
                $namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

                #Get the form name to be used as parameter in functions external to form...
                $FormName = $NamedNodes[0].Name


                #Define events functions
                #region Load, Draw (render) and closing form events
                #Things to load when the WPF form is loaded aka in memory
                $wpf.$FormName.Add_Loaded({
                    #Update-Cmd
                    $wpf.DataGrid.ItemsSource = $MailboxFeatures
                    $wpf.DataGrid.Columns | Foreach {
                        $_.CanUserSort = $true
                    }            
                })
                #Things to load when the WPF form is rendered aka drawn on screen
                $wpf.$FormName.Add_ContentRendered({
                    #Update-Cmd
                })
                $wpf.$FormName.add_Closing({
                    $msg = "Closed the MBX SIR and retention settings status list window"
                    write-host $msg
                })
                $wpf.btnClipboard.add_Click({
                    $CSVClip = $mailboxFeatures | ConvertTo-CSV -NoTypeInformation
                    $CSVClip | clip.exe
                    $title = "Copied !"
                    $msg = "Data copied to the clipboard ! `n`rUse CTRL+V on Notepad or on Excel !"
                    [System.Windows.MessageBox]::Show($msg,$title, "OK","Asterisk")
                })
                $wpf.btnClose.add_Click({
                    $wpf.$FormName.Close()
                })

                #endregion Load, Draw and closing form events
                #End of load, draw and closing form events

                #HINT: to update progress bar and/or label during WPF Form treatment, add the following:
                # ... to re-draw the form and then show updated controls in realtime ...
                $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


                # Load the form:
                # Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
                # Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
                # USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
                $async = $wpf.$FormName.Dispatcher.InvokeAsync({
                    $wpf.$FormName.ShowDialog() | Out-Null
                })
                $async.Wait() | Out-Null

                #endregion
                # end of Form definition for Get-MailboxFeaturesView
                
            }

            if ($wpf.chkArbitrationOnly.IsChecked){
                Get-MailboxQuotas $List -Arbitration
            } Else {
                Get-MailboxQuotas $List
            }
        }
        #endregion
        #End of the List Quotas region
        #region Single Item Recovery Status
        "List Single Item Recovery status" {
            Write-host "Displaying Mailbox SIR and retention settings status"
            $SelectedITems = $wpf.GridView.SelectedItems
            Write-host "Displaying Mailbox Single Item Recovery and retention settings status for $($SelectedItems.count) items..."
            $List = @()
            $SelectedItems | Foreach {
                $List += $_.primarySMTPAddress.tostring()
            }
            #$List = $List -join ","
            Function Get-MailboxSIRView {
                [CmdLetBinding()]
                Param(
                    [Parameter(Mandatory = $False, Position = 1)][string[]]$List,
                    [Parameter(Mandatory = $False, Position = 2)][switch]$Arbitration
                )
                #Initiating stopwatch to measure the time it takes to retrieve mailboxes
                $stopwatch = [system.diagnostics.stopwatch]::StartNew()
                If ($Arbitration){
                    $QueryMailboxFeatures = $List | get-mailbox -Arbitration | Select DisplayName,Name,*item*,OrganizationalUnit
                } Else {
                    $QueryMailboxFeatures = $List | Get-Mailbox | Select DisplayName,Name, *item*, OrganizationalUnit
                }
                [System.Collections.IENumerable]$MailboxFeatures = @($QueryMailboxFeatures)
                Write-host $($MailboxFeatures | ft | out-string)
                
                #Stopping stopwatch
                $stopwatch.Stop()
                $msg = "`n`nInstruction took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds ..."
                Write-Host $msg
                $msg = $null
                $StopWatch = $null

                #region Get-MailboxFeaturesView Form definition
                # Load a WPF GUI from a XAML file build with Visual Studio
                Add-Type -AssemblyName presentationframework, presentationcore
                $wpf = @{ }
                # NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
                #$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
                $inputXML = @"
                <Window x:Name="frmMbxSIRStatus" x:Class="Get_CASMAilboxFeaturesWPF.MainWindow"
                                        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                                        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                        xmlns:local="clr-namespace:Get_CASMAilboxFeaturesWPF"
                                        mc:Ignorable="d"
                                        Title="Mailboxes Single Item Recovery and Retention settings status" Height="450" Width="872.145" ResizeMode="NoResize">
                    <Grid>
                        <DataGrid x:Name="DataGridCASMbx" HorizontalAlignment="Left" Height="326" Margin="10,10,-59,0" VerticalAlignment="Top" Width="844" IsReadOnly="True"/>
                        <Button x:Name="btnClose" Content="Close" HorizontalAlignment="Left" Margin="748,352,0,0" VerticalAlignment="Top" Width="106" Height="46"/>
                        <Button x:Name="btnClipboard" Content="Copy to clipboard" HorizontalAlignment="Left" Margin="10,352,0,0" VerticalAlignment="Top" Width="174" Height="46"/>

                    </Grid>
                </Window>   
"@

                $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
                [xml]$xaml = $inputXMLClean
                $reader = New-Object System.Xml.XmlNodeReader $xaml
                $tempform = [Windows.Markup.XamlReader]::Load($reader)
                $namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
                $namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

                #Get the form name to be used as parameter in functions external to form...
                $FormName = $NamedNodes[0].Name


                #Define events functions
                #region Load, Draw (render) and closing form events
                #Things to load when the WPF form is loaded aka in memory
                $wpf.$FormName.Add_Loaded({
                    #Update-Cmd
                    $wpf.DataGridCASMbx.ItemsSource = $MailboxFeatures
                })
                #Things to load when the WPF form is rendered aka drawn on screen
                $wpf.$FormName.Add_ContentRendered({
                    #Update-Cmd
                })
                $wpf.$FormName.add_Closing({
                    $msg = "Closed the MBX SIR and retention settings status list window"
                    write-host $msg
                })
                $wpf.btnClipboard.add_Click({
                    $CSVClip = $mailboxFeatures | ConvertTo-CSV -NoTypeInformation
                    $CSVClip | clip.exe
                    $title = "Copied !"
                    $msg = "Data copied to the clipboard ! `n`rUse CTRL+V on Notepad or on Excel !"
                    [System.Windows.MessageBox]::Show($msg,$title, "OK","Asterisk")
                })
                $wpf.btnClose.add_Click({
                    $wpf.$FormName.Close()
                })

                #endregion Load, Draw and closing form events
                #End of load, draw and closing form events

                #HINT: to update progress bar and/or label during WPF Form treatment, add the following:
                # ... to re-draw the form and then show updated controls in realtime ...
                $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


                # Load the form:
                # Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
                # Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
                # USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
                $async = $wpf.$FormName.Dispatcher.InvokeAsync({
                    $wpf.$FormName.ShowDialog() | Out-Null
                })
                $async.Wait() | Out-Null

                #endregion
                # end of Form definition for Get-MailboxFeaturesView
                
            }

            If ($wpf.chkArbitrationOnly.IsChecked){
                Get-MailboxSIRView $List -Arbitration
            } Else {
                Get-MailboxSIRView $List
            }
            
        }
        #endregion
        #End of the Single Item Recovery status region
        #region List Mailbox Features
        "List Mailbox Features"  {
            if ($wpf.chkArbitrationOnly.IsChecked){
                # Option #4 - a message, a title, buttons, and an icon
                # More info : https://msdn.microsoft.com/en-us/library/system.windows.messageboximage.aspx
                $msg = "Arbitration mailboxes is checked - cannot get mailbox features for Arbitration mailboxes"
                $Title = "Error - Arbitration mailboxes don't get features"
                $Button = "Ok"
                $Icon = "Error"
                [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
                Return
            }
            Write-host "Displaying Mailbox Features"
            $SelectedITems = $wpf.GridView.SelectedItems
            Write-host "Displaying Mailbox Features for $($SelectedItems.count) items..."
            $List = @()
            $SelectedItems | Foreach {
                $List += $_.primarySMTPAddress.tostring()
            }
            #$List = $List -join ","
            Function Get-MailboxFeaturesView {
                [CmdLetBinding()]
                Param(
                    [Parameter(Mandatory = $False, Position = 1)][string[]]$List
                )

                #Initiating stopwatch to measure the time it takes to retrieve mailboxes
                $stopwatch = [system.diagnostics.stopwatch]::StartNew()

                $QueryMailboxFeatures = $List | Get-CASMAilbox | Select DisplayName, *enabled, *MAPIblock*
                [System.Collections.IENumerable]$MailboxFeatures = @($QueryMailboxFeatures)
                Write-host $($MailboxFeatures | ft DisplayName, ActiveSyncEnabled,OWAEnabled,ECPEnabled,MAPIEnabled,MAPIBlockOutlookRpcHttp,MapiHttpEnabled  -a | out-string)

                #Stopping stopwatch
                $stopwatch.Stop()
                $msg = "`n`nInstruction took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds..."
                Write-Host $msg
                $msg = $null
                $StopWatch = $null

                #region Get-MailboxFeaturesView Form definition
                # Load a WPF GUI from a XAML file build with Visual Studio
                Add-Type -AssemblyName presentationframework, presentationcore
                $wpf = @{ }
                # NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
                #$inputXML = Get-Content -Path ".\WPFGUIinTenLines\MainWindow.xaml"
                $inputXML = @"
                <Window x:Name="frmCASMBOXProps" x:Class="Get_CASMAilboxFeaturesWPF.MainWindow"
                                        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                                        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
                                        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
                                        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                        xmlns:local="clr-namespace:Get_CASMAilboxFeaturesWPF"
                                        mc:Ignorable="d"
                                        Title="Mailbox features enabled and blocked status" Height="450" Width="872.145" ResizeMode="NoResize">
                    <Grid>
                        <DataGrid x:Name="DataGridCASMbx" HorizontalAlignment="Left" Height="326" Margin="10,10,-59,0" VerticalAlignment="Top" Width="844" IsReadOnly="True"/>
                        <Button x:Name="btnClose" Content="Close" HorizontalAlignment="Left" Margin="748,352,0,0" VerticalAlignment="Top" Width="106" Height="46"/>
                        <Button x:Name="btnClipboard" Content="Copy to clipboard" HorizontalAlignment="Left" Margin="10,352,0,0" VerticalAlignment="Top" Width="174" Height="46"/>

                    </Grid>
                </Window>         
"@

                $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
                [xml]$xaml = $inputXMLClean
                $reader = New-Object System.Xml.XmlNodeReader $xaml
                $tempform = [Windows.Markup.XamlReader]::Load($reader)
                $namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
                $namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

                #Get the form name to be used as parameter in functions external to form...
                $FormName = $NamedNodes[0].Name

                #Define events functions
                #region Load, Draw (render) and closing form events
                #Things to load when the WPF form is loaded aka in memory
                $wpf.$FormName.Add_Loaded({
                    #Update-Cmd
                    $wpf.DataGridCASMbx.ItemsSource = $MailboxFeatures
                })
                #Things to load when the WPF form is rendered aka drawn on screen
                $wpf.$FormName.Add_ContentRendered({
                    #Update-Cmd
                })
                $wpf.$FormName.add_Closing({
                    $msg = "Closed the MBX features list window"
                    write-host $msg
                })
                $wpf.btnClipboard.add_Click({
                    $CSVClip = $mailboxFeatures | ConvertTo-CSV -NoTypeInformation
                    $CSVClip | clip.exe
                    $title = "Copied !"
                    $msg = "Data copied to the clipboard ! `n`rUse CTRL+V on Notepad or on Excel !"
                    [System.Windows.MessageBox]::Show($msg,$title, "OK","Asterisk")
                })
                $wpf.btnClose.add_Click({
                    $wpf.$FormName.Close()
                })

                #endregion Load, Draw and closing form events
                #End of load, draw and closing form events

                #HINT: to update progress bar and/or label during WPF Form treatment, add the following:
                # ... to re-draw the form and then show updated controls in realtime ...
                $wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})


                # Load the form:
                # Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
                # Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
                # USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
                $async = $wpf.$FormName.Dispatcher.InvokeAsync({
                    $wpf.$FormName.ShowDialog() | Out-Null
                })
                $async.Wait() | Out-Null

                #endregion
                # end of Form definition for Get-MailboxFeaturesView
                
            }

            Get-MailboxFeaturesView $List
        }
        #endregion
        #End of the List Mailbox Features region
    }
}

Function Update-Label ($msg) {
    $wpf.lblStatus.Content = $msg
    $Wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function Working-Label {
        # Trick to enable a Label to update during work :
    # Follow with "Dispatcher.Invoke("Render",[action][scriptblobk]{})" or [action][scriptblock]::create({})
    $wpf.$FormName.IsEnabled = $False
    $wpf.lblStatus.Content = "Working ..."
    $wpf.lblStatus.ForeGround = [System.Windows.Media.Brushes]::Red
    $wpf.lblStatus.BackGround = [System.Windows.Media.Brushes]::Blue
    $Wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function Ready-Label{
    $wpf.$FormName.IsEnabled = $True
    $wpf.lblStatus.Content = "Ready !"
    $wpf.lblStatus.ForeGround = [System.Windows.Media.Brushes]::Green
    $wpf.lblStatus.BackGround = [System.Windows.Media.Brushes]::Yellow
    $Wpf.$FormName.Dispatcher.Invoke("Render",[action][scriptblock]{})
}

Function Update-MainCommandLine {
    If ($wpf.chkArbitrationOnly.IsChecked){
        $commandLine = "Get-Mailbox -ResultSize Unlimited -Arbitration -ErrorAction Stop"
    } Else {
        If ($wpf.txtMailboxString.text -eq ""){
            $SearchSubstring = ("*")
        } Else {
            $SearchSubstring = ("*") + ($wpf.txtMailboxString.text) + ("*")
        }
        If ($wpf.chkUnlimited.IsChecked){
            $ResultSize = "Unlimited"
        } Else {
            $ResultSize = $wpf.txtResultSize.Text
        }
        $chkIncludeDiscovery = $false
        If ($chkIncludeDiscovery){
            $commandLine = "Get-Mailbox -ResultSize $ResultSize -Identity $SearchSubstring -ErrorAction Stop| Select Name,Alias,DisplayName,primarySMTPAddress,ServerName,Database,RecipientTypeDetails,OrganizationalUnit"
        } Else {
            $commandLine = "Get-Mailbox -ResultSize $ResultSize -Identity $SearchSubstring -Filter {RecipientTypeDetails -ne `"DiscoveryMailbox`"} -ErrorAction Stop| Select Name,Alias,DisplayName,primarySMTPAddress,ServerName,Database,REcipientTypeDetails,OrganizationalUnit"
        }
    }
    $wpf.txtMainCommand.Text = $CommandLine
}

$lblabout_Click = {
    $Language = "EN"
    switch ($Language)
    {
        "EN"
        {
            $systemst = "QXV0aG9yOiBTYW0gRHJleQ0Kc2FtZHJleUBtaWNyb3NvZnQuY29tDQpzYW1teUBob3RtYWlsLmZyDQpNaWNyb3NvZnQgRW`
        5naW5lZXIgc2luY2UgT2N0IDE5OTkNCjE5OTktMjAwMDogUHJlc2FsZXMgRW5naW5lZXIgKEZyYW5jZSkNCjIwMDAtMjAwMzogU3VwcG9yd`
        CBFbmdpbmVlciAoRnJhbmNlKQ0KMjAwMy0yMDA2OiB2ZXJ5IGZpcnN0IFBGRSBpbiBGcmFuY2UNCjIwMDYtMjAwOTogTUNTIENvbnN1bHRhb`
        nQgKEZyYW5jZSkNCjIwMDktMjAxMDogVEFNIChGcmFuY2UpDQoyMDEwLW5vdyA6IENvbnN1bHRhbnQgKENhbmFkYSkNCk11c2ljaWFuLCBjb`
        21wb3NlciAoS2V5Ym9hcmQsIEd1aXRhcikNClBsYW5lIHBpbG90IHNpbmNlIDE5OTUNCkZvciBTaGFyZWQgU2VydmljZXMgQ2FuYWRh"
        } 
        "FR"
        {
            $systemst = "QXV0ZXVyOiBTYW0gRHJleQ0Kc2FtZHJleUBtaWNyb3NvZnQuY29tDQpzYW1teUBob3RtYWlsLmZyDQpJbmfDqW5pZXVyIGNo`
        ZXogTWljcm9zb2Z0IGRlcHVpcyBPY3QgMTk5OQ0KMTk5OS0yMDAwOiBJbmfDqW5pZXVyIEF2YW50LVZlbnRlIChGcmFuY2UpDQoyMDAwLTIwMD`
        M6IFNww6ljaWFsaXN0ZSBUZWNobmlxdWUgKEZyYW5jZSkNCjIwMDMtMjAwNjogUHJlbWllciBQRkUgZW4gRnJhbmNlDQoyMDA2LTIwMDk6IENv`
        bnN1bHRhbnQgTUNTIChGcmFuY2UpDQoyMDA5LTIwMTA6IFJlc3BvbnNhYmxlIFRlY2huaXF1ZSBkZSBDb21wdGUgKEZyYW5jZSkNCjIwMTAtMjA`
        xNiA6IENvbnN1bHRhbnQgKENhbmFkYSkNCk11c2ljaWVuLCBjb21wb3NpdGV1ciAoQ2xhdmllciwgR3VpdGFyZSkNCkJyZXZldCBkZSBQaWxvdGU`
        gUHJpdsOpIGRlcHVpcyAxOTk1DQpQb3VyIFNlcnZpY2VzIFBhcnRhZ8OpcyBDYW5hZGE="
        }
    }
    $systemst = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($systemst))
    # Option #4 - a message, a title, buttons, and an icon
    # More info : https://msdn.microsoft.com/en-us/library/system.windows.messageboximage.aspx
    $msg = $systemst
    $Title = $wpf.$FormName.Title
    $Button = "Ok"
    $Icon = "Information"
    [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
}


Function Get-Mailboxes {
    If ($wpf.chkUnlimited.IsChecked){
        Write-host "Unlimited specified ... ignoring Resultsize number..."
    } Else {
        If ($([int]$wpf.txtResultSize.Text) -gt 1000) {Write-Host "$($wpf.txtResultSize.Text) is greater than 1000 ..."} Else {write-host "$($wpf.txtResultSize.Text) is less than 1000"}
    }

    If ($([int]$wpf.txtResultSize.Text) -gt 1000 -or $wpf.chkUnlimited.IsChecked){
        # Option #4 - a message, a title, buttons, and an icon
        # More info : https://msdn.microsoft.com/en-us/library/system.windows.messageboximage.aspx
        if ($wpf.chkUnlimited.IsChecked) {
            $Specified = $wpf.chkUnlimited.Content
        } Else {
            $Specified = "$($wpf.txtResultSize.Text), which is more than 1000"
        }
        $msg = "WARNING: You specified -> $Specified <- for the Resultsize, mailbox collection can take a LOT of time, Continue ? (Y/N)"
        $Title = "Question..."
        $Button = "YesNo"
        $Icon = "Question"
        $Answer = [System.Windows.MessageBox]::Show($msg,$Title, $Button, $icon)
        If($Answer -eq "No"){Return}
    }
    Try {
        #Initiating stopwatch to measure the time it takes to retrieve mailboxes
        $stopwatch = [system.diagnostics.stopwatch]::StartNew()
        #Getting the command line from the text box where it's generated
        $commandLine = $wpf.txtMainCommand.text
        #Invoking the command line and storing in a variable
        $Mailboxes = invoke-expression $CommandLine
        $NewMailboxesObj = @()

        Foreach ($objTemp in $Mailboxes){
            If ($($objTemp.OrganizationalUnit) -match "prod.outlook.com"){
                $objtemp | Add-Member -NotePropertyName Location -NotePropertyValue "Cloud"
            } Else {
                $objtemp | Add-Member -NotePropertyName Location -NotePropertyValue "On-prem"
            }
            $NewMailboxesObj += $objtemp
        }

        $Mailboxes =  $NewMailboxesObj | Select Name,Alias,DisplayName,primarySMTPAddress, RecipientTypeDetails,Location,ServerName,Database
        #Stopping stopwatch
        $stopwatch.Stop()
        $msg = "`n`nInstruction took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to retrieve $($Mailboxes.count) mailboxes..."
        Write-Host $msg
        $msg = $null
        $StopWatch = $null

        #Populating the GridView
        [System.Collections.IENumerable]$Results = @($Mailboxes)
        $wpf.GridView.ItemsSource = $Results
        $wpf.GridView.Columns | Foreach {
            $_.CanUserSort = $true
        }
        $wpf.lblNbItemsInGrid.Content = $($Results.Count)
    } Catch {
        $Mailboxes = $null
        $wpf.GridView.ItemsSource = $null
        write-host "ZERO MAILBOXES"
        $wpf.lblNbItemsInGrid.Content = 0
    }
}

#endregion

#========================================================
#region WPF form definition and load controls
#========================================================

# Load a WPF GUI from a XAML file build with Visual Studio
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{}
# NOTE: Either load from a XAML file or paste the XAML file content in a "Here String"
# $inputXML = Get-Content -Path "C:\Users\Kamehameha\Documents\GitHub\PowerShell\Get-EventsFromEventLog\VisualStudio2017WPFDesign\Launch-EventsCollector-WPF\Launch-EventsCollector-WPF\MainWindow.xaml"
$inputXML = @"
<Window x:Name="WForm" x:Class="GridView_WPF.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:GridView_WPF"
        mc:Ignorable="d"
        Title="Search Mailboxes" Height="722.39" Width="1026.5" ResizeMode="NoResize">
    <Grid>
        <DataGrid x:Name="GridView" HorizontalAlignment="Left" Height="425" Margin="353,10,0,0" VerticalAlignment="Top" Width="641"/>
        <TextBox x:Name="txtMailboxString" HorizontalAlignment="Left" Height="23" Margin="10,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="338"/>
        <Label Content="Search for mailbox (substring of alias, e-mail address, &#xD;&#xA;display name, ...)" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,11,0,0" Height="51" Width="302"/>
        <Button x:Name="btnRun" Content="Search" HorizontalAlignment="Left" Margin="10,95,0,0" VerticalAlignment="Top" Width="75" Height="32">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Label x:Name="lblStatus" Content="Ready !" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="10,658,0,0" Width="1010" FontStyle="Italic" FontWeight="Bold">
        </Label>
        <Button x:Name="btnAction" Content="Run on selected" Margin="360,596,0,0" IsEnabled="False" HorizontalAlignment="Left" VerticalAlignment="Top">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <ComboBox x:Name="comboSelectAction" HorizontalAlignment="Left" Margin="360,534,0,0" VerticalAlignment="Top" Height="30" SelectedIndex="0" IsEnabled="False" TextOptions.TextFormattingMode="Display">
            <ComboBox.Effect>
                <DropShadowEffect/>
            </ComboBox.Effect>
            <ComboBoxItem Content="List Mailbox Features"/>
            <ComboBoxItem Content="List Single Item Recovery status"/>
            <ComboBoxItem Content="List Quotas"/>
            <ComboBoxItem Content="Disable Mailbox"/>
        </ComboBox>
        <Label x:Name="lblNbItemsInGrid" Content="0" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="506,440,0,0" Width="66"/>
        <Label Content="Number of Items in Grid:" HorizontalAlignment="Left" Margin="353,440,0,0" VerticalAlignment="Top" Width="148"/>
        <Label Content="Selected:" HorizontalAlignment="Left" Margin="353,471,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblNumberItemsSelected" Content="0" HorizontalAlignment="Left" Margin="460,471,0,0" VerticalAlignment="Top" Width="67"/>
        <TextBox x:Name="txtResultSize" HorizontalAlignment="Left" Height="23" Margin="224,98,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="124" Text="100"/>
        <TextBlock HorizontalAlignment="Left" Margin="95,95,0,0" TextWrapping="Wrap" Text="ResultSize (aka Nb of mailboxes to display):" VerticalAlignment="Top" Width="124"/>
        <Label Content="Status:" HorizontalAlignment="Left" Margin="10,627,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtMainCommand" HorizontalAlignment="Left" Height="132" Margin="10,200,0,0" TextWrapping="Wrap" Text="Get-Mailbox command to be run..." VerticalAlignment="Top" Width="338" IsReadOnly="True"/>
        <Rectangle HorizontalAlignment="Left" Height="26" Margin="353,440,0,0" VerticalAlignment="Top" Width="232">
            <Rectangle.Stroke>
                <SolidColorBrush Color="{DynamicResource {x:Static SystemColors.ActiveBorderColorKey}}"/>
            </Rectangle.Stroke>
        </Rectangle>
        <Rectangle HorizontalAlignment="Left" Height="26" Margin="353,471,0,0" VerticalAlignment="Top" Width="174">
            <Rectangle.Stroke>
                <SolidColorBrush Color="{DynamicResource {x:Static SystemColors.ActiveBorderColorKey}}"/>
            </Rectangle.Stroke>
        </Rectangle>
        <Label Content="The command run when clicking on the Search button is:" HorizontalAlignment="Left" Margin="10,174,0,0" VerticalAlignment="Top" Width="338" FontStyle="Italic"/>
        <CheckBox x:Name="chkUnlimited" Content="Unlimited" HorizontalAlignment="Left" Margin="223,126,0,0" VerticalAlignment="Top"/>
        <Label x:Name="lblAbout" Content="." HorizontalAlignment="Left" Margin="999,0,0,0" VerticalAlignment="Top" Height="22" Width="21"/>
        <CheckBox x:Name="chkArbitrationOnly" Content="Arbitration only" HorizontalAlignment="Left" Margin="10,149,0,0" VerticalAlignment="Top"/>
        <Button x:Name="btnClipboard" Content="Copy all above items to clipboard" HorizontalAlignment="Left" Margin="590,442,0,0" VerticalAlignment="Top" Width="202" IsEnabled="False" Height="22">
            <Button.Effect>
                <DropShadowEffect/>
            </Button.Effect>
        </Button>
        <Label Content="#1 - Choose what you wish to see about the selected" HorizontalAlignment="Left" Margin="353,508,0,0" VerticalAlignment="Top"/>
        <Label Content="#2 - Run the above selected action on selected items" HorizontalAlignment="Left" Margin="353,570,0,0" VerticalAlignment="Top"/>
        <Border BorderBrush="{DynamicResource {x:Static SystemColors.ActiveBorderBrushKey}}" BorderThickness="1" HorizontalAlignment="Left" Height="126" Margin="353,508,0,0" VerticalAlignment="Top" Width="304"/>

    </Grid>
</Window>
"@

$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
[xml]$xaml = $inputXMLClean
$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

#Get the form name to be used as parameter in functions external to form...
$FormName = $NamedNodes[0].Name

#Cloud icon
#$Base64icon = "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"
#Exchange Management Shell icon
$Base64Icon = "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAYdEVYdFNvZnR3YXJlAHBhaW50Lm5ldCA0LjEuNBNAaMQAAAT/SURBVEhLnZV5TNNnGMfBYwOnmLA5TcTFRDz2xw4TDE6NDvHgCttUnM5MZIkynTcCRbEgyKmIZhDEDpkOSGBTBrI5dYCb8OtBixQo0BZb2tKWaltBBJSj3z1lYOSwjr3JNynh7ffzvs/zfZ/ajV4A7Ivq211ZRWrX5Wn1rk4sxnXpOZFb/G0NNmU3YFGcEC5RjPPQ9tcvDqN2TvxTFfwZpzF4Lrsq2InF27/352asv1iHBXEiOIQxsD/Ghf1RBnYhDN4IZeB0ZAKA5DuaQrfUGrwXK8S0cB4mkYk9yWo2nqwAu4kAEu9obBqO1v8EVI5rNp6GARn3tNodVxr1X10dqx2k038o9NnlSocRAPsQLqbQ5xlhXMxlC/Fh8n14fF8Lh1DeGEBQnswy+RiDSa/QVBLtcxwFYFCv68Tzfgt6+wfQNzCAZ339mBtVNQZQLjdLV1wQt71JIVgQK0K2QI8bEiPpEXw4Eky1Hmo8gNrcjfaePvxUZUBSWSsSSzWYGcEfA9hf0Bzsx5Fkr6EbXhe3IV/UhiyeHgpjFw4VKih9NgAtph6suFBLf1M8h4xHA3bnyS2L46tBEUfkb0rMYfMx66QAJfXG1wOMT3uRTCffnS/HslTxCMgwYFeuzLL8vBh7C2SYzRYM7pkXXYUymRkBV5qoBzYAFgvQ3TuAJ8/6cOqmarBpowGRJUrLuyd4mBXJR2COFNt+bIBrjBAHf5HD+QSf3hP31YBOMr4jNeFmgwl76Bb24wDCC5XzfS9JBo0cQyvBYXTI4urwDplbfWwC/u1B3QvTlzUM+LZAZlmWch9uKdX4hMroSOPkYoUGOVV6zKepMNkW4GFnLyJuKOGVKcFKSok178MAR4rfknX7jm7PUlnej2lCxDUtlidK4RIhxgen6nD5rgHbM2VwCRdhyaqvD4wCcKEyd6F/wAJT13O6STeqNU/oodFphgDzjtdjS3oPNqc/xaGcDmw9b4JPQhs2xrZiQ0wL/E8/gHe0FB60LyC+EWNu8MXlJrBKWpB2T4vCWiPuykwjAEtj1fA/34Odl7pQrXyKmAIj/BP08I7VwPuUAhvYUqw7UYe1rBocyJCPBFg1hcwcaKq+FcHDTErJ25Rt64Qd/B9pddIj+Kd24vPUDnBKH0Pa2ol9ma3wIfONUXIylwyafxoqRNYt/ViALU0Pq4LnGSP8Ujrgd8aEbal6lIofo1HVgS+TpNhwshGe1AuPMBE86R3wGtsnBphNX16fbITvWTN8kwzwidNiT7oKyrYuCOVGMq/F2vBqrCVzr3ABtlhTlPaXFtNZ3BdlsKWF7CZ4EcAvwQDfOB32ZGhRzDejVmFGWlEzndxqLoDHkUocTpfgSL7a0S5XoA9yiRYETQ3hBu3Kk5q86fEsShBhGo3sl3/drLn+KLoZPokP6eQ6HM8xgC/rwE2aogczGuDJuk/mQjKnPh38G5d/VyOfIcDLK4dv+DjkV4X7qtQa9xksrvu5Mm11WLESfj9IsDi+BitjlfCJ1+HSbRNaDN3ILVWBfVWO6xUGlPAf4hpVY2ecgAB3kZAr2xwFTBqyHn8x6nbnNL5hTmCBck5Ynirwm3QdxVGFvHI9kmjI+bLrEHS2AZySB8i51QxOsRQB7Er4RTDYFMldOGTz39Z3mW3+W1N02BitgBeVan0klSWcmhpKdQ/hUWkqsOZQOVYfKMfhNDGSciUEsLP7B/agGS/3eWk6AAAAAElFTkSuQmCC"

# Create a streaming image by streaming the base64 string to a bitmap streamsource
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($Base64icon)
$bitmap.EndInit()
$bitmap.Freeze()
 
# This is the icon in the upper left hand corner of the app
$wpf.$FormName.Icon = $bitmap
#========================================================
# END of WPF form definition and load controls
#endregion
#========================================================

#========================================================
#region WPF EVENTS definition
#========================================================

#region Buttons
$wpf.btnRun.add_Click({
    Working-Label
    Get-Mailboxes
    if (($wpf.lblNbItemsInGrid.Content) -eq 0){
        $wpf.btnClipboard.IsEnabled = $false
    } Else {
        $wpf.btnClipboard.IsEnabled = $true
    }
    Ready-Label
})

$wpf.btnAction.add_Click({
    Working-Label
    Run-Action
    Ready-Label
})

$wpf.btnClipboard.add_Click({
    $CSVClip = $wpf.GridView.ItemsSource | ConvertTo-CSV -NoTypeInformation
    $CSVClip | clip.exe
    $title = "Copied !"
    $msg = "Data copied to the clipboard ! `n`rUse CTRL+V on Notepad or on Excel !"
    [System.Windows.MessageBox]::Show($msg,$title, "OK","Asterisk")
})

# End of Buttons region
#endregion

#region Load, Draw (render) and closing form events
#Things to load when the WPF form is loaded aka in memory
$Wpf.$FormName.Add_Loaded({
    Ready-Label
    Update-MainCommandLine
    $wpf.$FormName.Title += " - v$Version"
})
#Things to load when the WPF form is rendered aka drawn on screen
$Wpf.$FormName.Add_ContentRendered({
    write-host "rendered"
})

$Wpf.$FormName.add_Closing({
    $msg = "bye bye !"
    write-host $msg
})

$wpf.lblAbout.Add_MouseLeftButtonDown($lblabout_Click)

# End of load, draw and closing form events
#endregion

#region Text Changed events

$wpf.GridView.add_SelectionChanged({
    $Selected = $wpf.GridView.SelectedItems.count
    If ($Selected -eq 0) {
        $wpf.btnAction.IsEnabled = $false
        $wpf.comboSelectAction.IsEnabled = $false
    } ElseIf ($Selected -gt 0) {
        $wpf.btnAction.IsEnabled = $true
        $wpf.comboSelectAction.IsEnabled = $true
    }
    $wpf.lblNumberItemsSelected.Content = $Selected
})

$wpf.txtMailboxString.add_TextChanged({
    Update-MainCommandLine
})

$wpf.txtResultSize.add_TextChanged({
    Update-MainCommandLine
})

$wpf.chkUnlimited.add_Click({
    Update-MainCommandLine
    If ($wpf.chkUnlimited.IsChecked){
        $wpf.txtResultSize.IsEnabled = $false
    } Else {
        $wpf.txtResultSize.IsEnabled = $true
    }
})

$wpf.chkArbitrationOnly.add_Click({
    Update-MainCommandLine
    If ($Wpf.chkArbitrationOnly.IsChecked){
        $wpf.txtMailboxString.IsEnabled = $false
        $wpf.txtResultSize.IsEnabled = $false
        $wpf.chkUnlimited.IsEnabled = $false
    } Else {
        $wpf.txtMailboxString.IsEnabled = $true
        $wpf.txtResultSize.IsEnabled = $true
        $wpf.chkUnlimited.IsEnabled = $true
    }
})
#End of Text Changed events
#endregion


#endregion

#=======================================================
#End of Events from the WPF form
#endregion
#=======================================================

IsPSV3 | out-null

Test-ExchTools | out-null

# Load the form:
# Older way >>>>> $wpf.MyFormName.ShowDialog() | Out-Null >>>>> generates crash if run multiple times
# Newer way >>>>> avoiding crashes after a couple of launches in PowerShell...
# USing method from https://gist.github.com/altrive/6227237 to avoid crashing Powershell after we re-run the script after some inactivity time or if we run it several times consecutively...
$async = $wpf.$FormName.Dispatcher.InvokeAsync({
    $wpf.$FormName.ShowDialog() | Out-Null
})
$async.Wait() | Out-Null