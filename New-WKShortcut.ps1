<#
Application to generate shortcuts in WrapperKing package format.
Author: Simon Mellergård | It-center, Värnamo kommun
Version: 1.0 | 2024-04-18
#>

function Initialize-WPF {

    [CmdletBinding()]

    param ()
    
    begin {
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    }
    
    process {

        #region Variable declaration
        $PathToContent = '.\Content'

        $Script:Icon     = Get-ChildItem -Path "$PathToContent\VMO-Icon.ico" | Select-Object -ExpandProperty Fullname
        $Script:Logo     = Get-ChildItem -Path "$PathToContent\VK_logotyp.png" | Select-Object -ExpandProperty Fullname
        $Script:Resource = Get-ChildItem -Path "$PathToContent\ResourceDictionary.xaml" | Select-Object -ExpandProperty Fullname
        #$Wave          = Get-ChildItem -Path "$PathToContent\VK_vag_RGB-lila_Edit_LoginScreen.png" | Select-Object -ExpandProperty Fullname
        #$Wave2         = Get-ChildItem -Path "$PathToContent\VK_vag_RGB-ljusbla_2x.png" | Select-Object -ExpandProperty Fullname
        #$Script:Tick   = Get-ChildItem -Path "$PathToContent\tick.png" | Select-Object -ExpandProperty Fullname
        #$Script:Cross  = Get-ChildItem -Path "$PathToContent\cross.png" | Select-Object -ExpandProperty Fullname #>

        $InputXML = Get-Content -Path "$PathToContent\MainWindow.xaml" -Encoding UTF8

        #clean up xml there is syntax which Visual Studio 2015 creates which PoSH can't understand
        $inputXMLClean = $inputXML -replace 'mc:Ignorable="d"',''`
                                -replace "x:Na",'Na'`
                                -replace 'x:Class=".*?"',''`
                                -replace 'd:DesignHeight="\d*?"',''`
                                -replace 'd:DesignWidth="\d*?"',''`
                                -replace "VK_logotyp.png", $Logo`
                                -replace "ResourceDictionary.xaml", $Resource`
                                -replace 'VMO-Icon.ico', $Icon

        #change string variable into xml
        [xml]$xaml = $inputXMLClean

        $Script:WPF = @{}

        #read xml data into xaml node reader object
        $reader = New-Object System.Xml.XmlNodeReader $xaml
    
        #create System.Windows.Window object
        $tempform = [Windows.Markup.XamlReader]::Load($reader)
    
        #select each named node using an Xpath expression.
        $namedNodes = $XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
    
        #add all the named nodes as members to the $wpf variable, this also adds in the correct type for the objects.
        $namedNodes | ForEach-Object {
            $WPF.Add($_.Name, $tempform.FindName($_.Name))
        }
    }
    
    end {
        
    }
}
# End function.

function Get-MessageBox {
    param (
        [ValidateSet('YesNo', 'YesNoCancel', 'Ok', 'OkCancel')]
        [Parameter(Mandatory = $true)]
        [string]$Type,

        [ValidateSet('Asterisk', 'Error', 'Exclamation', 'Hand', 'Information', 'None', 'Question', 'Stop', 'Warning')]
        [Parameter(Mandatory = $true)]
        [string]$Icon,

        [Parameter(Mandatory = $true)]
        [string]$Body,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    $ButtonType   = [System.Windows.MessageBoxButton]::$Type
    $MessageIcon  = [System.Windows.MessageBoxImage]::$Icon
    $MessageBody  = $Body
    $MessageTitle = $Title
    return $Script:MessageResult = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
}
# End function.

function Show-NewWKShortcut {

    [CmdletBinding()]

    param (
        
    )
    
    begin {
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
        #region Setting pre populated values
        $WPF.Version.Text = "1"
        $WPF.LocalContent.Text = "`$env:SystemDrive\eKlient"
        $WPF.ReturnType.SelectedItem = $WPF.ReturnType.Items | Where-Object {$_.Content -eq 'Computer'}
        
        #endregion Setting pre populated values
    }
    
    process {

        #region Setting textbox options
        $WPF.ShortcutName.Add_LostFocus({
            if ($WPF.ShortcutName.Text) {
                switch ($env:USERDOMAIN) {
                    VARNAMO {$WPF.AppName.Text = "$($WPF.ShortcutName.Text) VK x64 R01 1.0"}
                    Default {$WPF.AppName.Text = "$($WPF.ShortcutName.Text) x64 R01 1.0"}
                }
            }
        })

        $WPF.Version.Add_LostFocus({
            if ($WPF.AppName.Text) {
                $NameVersion = $WPF.AppName.Text -split " " | Select-Object -Last 1
                $WPF.AppName.Text = $WPF.AppName.Text -replace $NameVersion, $WPF.Version.Text
            }
        })

        $WPF.ReturnType.Add_SelectionChanged({
            if ($WPF.ReturnType.SelectedItem.Content -eq "User") {
                $WPF.StartMenuContainer.Focusable = $false
                $WPF.StartMenuContainer.Foreground = "Gray"
                switch -Wildcard ($WPF.Target.Text) {
                    "http*" {$WPF.StartMenuContainer.Text = "`$env:APPDATA\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"}
                    Default {$WPF.StartMenuContainer.Text = "`$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$($WPF.ShortcutName.Text)"}
                }
            }
            elseif ($WPF.ReturnType.SelectedItem.Content -eq "Computer") {
                $WPF.StartMenuContainer.Focusable = $true
                $WPF.StartMenuContainer.Foreground = "Black"
                switch -Wildcard ($WPF.Target.Text) {
                    "http*" {$WPF.StartMenuContainer.Text = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"}
                    Default {$WPF.StartMenuContainer.Text = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\$($WPF.ShortcutName.Text)"}
                }
            }
        })

        $WPF.Target.Add_LostFocus({
            if ($WPF.Target.Text) {
                switch ($WPF.ReturnType.SelectedItem.Content) {
                    User {
                        switch -Wildcard ($WPF.Target.Text) {
                            "http*" {$WPF.StartMenuContainer.Text = "`$env:APPDATA\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"}
                            Default {$WPF.StartMenuContainer.Text = "`$env:APPDATA\Microsoft\Windows\Start Menu\Programs\$($WPF.ShortcutName.Text)"}
                        }
                    }
                    Computer {
                        switch -Wildcard ($WPF.Target.Text) {
                            "http*" {$WPF.StartMenuContainer.Text = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"}
                            Default {$WPF.StartMenuContainer.Text = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\$($WPF.ShortcutName.Text)"}
                        }
                    }
                }
            }
        })
        #endregion Setting textbox options

        #region Setting button actions
        $WPF.CloseButton.Add_Click({$WPF.NewWKShortcut.Close()})
        $WPF.MinimizeButton.Add_Click({$WPF.NewWKShortcut.WindowState = "Minimized"})
        $WPF.btnQuit.Add_Click({$WPF.NewWKShortcut.Close()})
        $WPF.btnOutput.Add_Click({
            $OutputDialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $OutputDialog.Description = "Select folder to store WrapperKing package"
            $Form                   = New-Object System.Windows.Forms.Form -Property @{TopMost = $true}
            $Result = $OutputDialog.ShowDialog($Form)

            if ($Result -eq [Windows.Forms.DialogResult]::OK) {
                $WPF.OutPath.Text = $OutputDialog.SelectedPath
            }
        })
        $WPF.btnIcon.Add_Click({
            $IconDialog        = New-Object System.Windows.Forms.OpenFileDialog
            $IconDialog.Title  = "Select icon to be used in the WrapperKing package"
            $IconDialog.filter = "Icon files (*.ico)|*.ico|All files (*.*)| *.*"
            $IconDialog.ShowDialog() | Out-Null

            $WPF.IconPath.Text = $IconDialog.filename
        })
        $WPF.btnReset.Add_Click({
            $WPF.ShortcutName.Text = ""
            $WPF.AppName.Text = ""
            $WPF.Vendor.Text = ""
            $WPF.Version.Text = "1"
            $WPF.Target.Text = ""
            $WPF.IconPath.Text = ""
            $WPF.LocalContent.Text = "`$env:SystemDrive\eKlient"
            $WPF.ReturnType.SelectedItem = $WPF.ReturnType.Items | Where-Object {$_.Content -eq 'Computer'}
            $WPF.OutPath.Text = ""
        })
        $WPF.btnGenerate.Add_Click({

            $PropertiesToCheck = @(
                "ShortcutName",
                "AppName",
                "Vendor",
                "Version",
                "Target",
                "StartMenuContainer",
                "IconPath",
                "LocalContent",
                "ReturnType",
                "OutPath"
            )

            $Missing = foreach ($Property in $PropertiesToCheck) {
                if (-not ($WPF.$Property.Text)) {
                    $Property
                }
            }

            if ($Missing) {
                Get-MessageBox -Type Ok -Icon Exclamation -Body "Missing information in the following text fields:`n$Missing" -Title "Missing information"
                Clear-Variable Missing
            }
            else {
                $ShortcutParams = @{
                    ShortcutName       = $WPF.ShortcutName.Text
                    AppName            = $WPF.AppName.Text
                    Vendor             = $WPF.Vendor.Text
                    Version            = $WPF.Version.Text
                    Target             = $WPF.Target.Text
                    StartMenuContainer = $WPF.StartMenuContainer.Text
                    IconPath           = $WPF.IconPath.Text
                    LocalContent       = $WPF.LocalContent.Text
                    ReturnType         = $WPF.ReturnType.SelectedItem.Content
                    OutPath            = $WPF.OutPath.Text
                }
    
                New-eKlientWKShortcut @ShortcutParams

                Get-MessageBox -Type Ok -Icon Asterisk -Body "$($ShortcutParams.OutPath)\$($ShortcutParams.AppName) has been successfully been created!" -Title Success!
            }

            
        })

        #endregion Setting button actions

        
    }
    
    end {
        $WPF.NewWKShortcut.ShowDialog()
    }
}
# End function.

function New-eKlientWKShortcut {

    [CmdletBinding()]

    param (
        # Name of the shortcut
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $ShortcutName,

        # Name of the shortcut
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $AppName,

        # Name of the manufacturer/vendor
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $Vendor,

        # Version of the wrapper package
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $Version,

        # Target for the shortcut
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $Target,

        # Name of the container in the start menu
        [Parameter()]
        [string]
        $StartMenuContainer,

        # Path to icon
        [Parameter(
            Mandatory = $false
        )]
        [string]
        $IconPath,

        # Path to local content
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $LocalContent,

        # If the installation should be applied in user context or computer context.
        [Parameter(
            Mandatory = $true
        )]
        [ValidateSet(
            "User",
            "Computer"
        )]
        [string]
        $ReturnType,

        # Where to place the reulting package
        [Parameter(
            Mandatory = $true
        )]
        [string]
        $OutPath,

        # Path to InstallKing.exe
        [Parameter(
            DontShow = $true
        )]
        [ValidateScript({
            if (Test-Path -Path $_) {
                return $true
            }
            else {
                #throw $_.Exception.Message
                Get-MessageBox -Type Ok -Icon Stop -Body "$($_.Exception.Message)" -Title "WrapperKing not installed"
            }
        })]
        [string]
        $InstallKing = $PathToInstallKing
    )
    
    begin {
        #region Declare internal functions
        function Format-XMLElement {

            [CmdletBinding()]
        
            param (
                # Orderobject
                [Parameter()]
                [System.Object]
                $InputObject,

                # Type of input
                [Parameter()]
                [ValidateSet(
                    'General',
                    'Initialize eKlient Folder',
                    'Launch URL Action',
                    'Install/Uninstall Order',
                    'Undo Create URL Action',
                    'Restore state file',
                    'Restore state file Uninstall',
                    'Rename state file'
                )]
                [string]
                $Type
            )
            
            begin {

            }
            
            process {
        
                switch ($Type) {
                    'General' {
                        
                        $XMLWriter.WriteStartElement("Appname")
                            $XMLWriter.WriteValue($AppName)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Manufacturer")
                            $XMLWriter.WriteValue($ShortcutName)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("ProductVersion")
                            $XMLWriter.WriteValue($Version)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("LogCmds")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("AppCount")
                            $XMLWriter.WriteValue("0")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("UninstallOrderModified")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("UninstallOrder")
                            # $XMLWriter.WriteValue()
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("GUID")
                            $XMLWriter.WriteValue("{$(((New-Guid).Guid).ToUpper())}")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("RepairAlternatives")
                            $XMLWriter.WriteValue("Silent")
                        $XMLWriter.WriteEndElement()
                    }
                    'Initialize eKlient Folder' {
                        $Guid = "{$(((New-Guid).Guid).ToUpper())}"
                        $XMLWriter.WriteStartElement("CommandAction")
                            $XMLWriter.WriteStartElement("uid")
                            $XMLWriter.WriteValue($Guid)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("type")
                            $XMLWriter.WriteValue("PsScript")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Cmd")
                            $XMLWriter.WriteValue($(Initialize-eKlientFolder -OutputType Encoded))
                            Initialize-eKlientFolder -OutputType ScriptFile -ScriptName $ShortcutName
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("exitcode")
                            $XMLWriter.WriteValue("0")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("IgnoreExitCode")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Comment")
                            $XMLWriter.WriteValue("Initialize eKlient Folder")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("SupportedArchitecture")
                            $XMLWriter.WriteValue("All")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Version")
                            $XMLWriter.WriteValue("1,0,0,0")
                            $XMLWriter.WriteEndElement()
                        $XMLWriter.WriteEndElement()
        
                        $TMPObject = @{                        
                            $InstCounter = [ordered]@{
                                ActionType = "CommandAction"
                                CmdType    = "PsScript"
                                index      = $InstCounter
                                uid        = $Guid
                                Text       = "Initialize eKlient Folder"
                                Type       = "Install"
                            }
                        }
        
                        $Script:OrderObject.Install += $TMPObject
        
                        $Script:InstCounter++
                    }
                    'Launch URL Action' {
                        switch ($ReturnType) {
                            User     {
                                $Script:Guid = "{$(((New-Guid).Guid).ToUpper())}"
                                $XMLWriter.WriteStartElement("WrapperApp")
                                    $XMLWriter.WriteStartElement("uid")
                                    $XMLWriter.WriteValue($Guid)
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("type")
                                    $XMLWriter.WriteValue("AppPs1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("CmdType")
                                    $XMLWriter.WriteValue("RunCmd")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("Appname")
                                    $XMLWriter.WriteValue("$ShortcutName")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UseArp")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallCommand")
                                    $XMLWriter.WriteValue("WorkFiles\Uninstall-$ShortcutName.ps1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ArgsUninstall")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SwitchUninstall")
                                    $XMLWriter.WriteValue("-ExecutionPolicy Bypass -File")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallMotor")
                                    $XMLWriter.WriteValue("powershell.exe")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("InstallCommand")
                                    $XMLWriter.WriteValue("WorkFiles\Install-$ShortcutName.ps1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SwitchInstall")
                                    $XMLWriter.WriteValue("-ExecutionPolicy Bypass -File")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ArgsInstall")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("InstallMotor")
                                    $XMLWriter.WriteValue("powershell.exe")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("MSPList")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ProductVersion")
                                    $XMLWriter.WriteValue("1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("Manufacturer")
                                    $XMLWriter.WriteValue($Vendor)
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("IgnoreExitCode")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("CustomExitCode")
                                    $XMLWriter.WriteValue("0")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SuppressDetectExitCodes")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("DetectionRules")
                                        $XMLWriter.WriteStartElement("DetectionRule")
                                            $XMLWriter.WriteStartElement("DetectionType")
                                            $XMLWriter.WriteValue("FileExists")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Path")
                                            $XMLWriter.WriteValue("%SystemDrive%\eKlient\$ShortcutName")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Name")
                                            $XMLWriter.WriteValue("$ShortcutName.state")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Value")
                                            # $XMLWriter.WriteValue("")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("VersionOperand")
                                            $XMLWriter.WriteValue("IsEqual")
                                            $XMLWriter.WriteEndElement()
                                        $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallOrder")
                                    $XMLWriter.WriteValue("0")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ReverseUsage")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("AlwaysUninstall")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SupportedArchitecture")
                                    $XMLWriter.WriteValue("All")
                                    $XMLWriter.WriteEndElement()
                                $XMLWriter.WriteEndElement()
                            }
                            Computer {
                                $Script:Guid = "{$(((New-Guid).Guid).ToUpper())}"
                                $XMLWriter.WriteStartElement("WrapperApp")
                                    $XMLWriter.WriteStartElement("uid")
                                    $XMLWriter.WriteValue($Guid)
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("type")
                                    $XMLWriter.WriteValue("AppPs1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("CmdType")
                                    $XMLWriter.WriteValue("RunCmd")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("Appname")
                                    $XMLWriter.WriteValue("$ShortcutName")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UseArp")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallCommand")
                                    $XMLWriter.WriteValue("WorkFiles\Uninstall-$ShortcutName.ps1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ArgsUninstall")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SwitchUninstall")
                                    $XMLWriter.WriteValue("-ExecutionPolicy Bypass -File")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallMotor")
                                    $XMLWriter.WriteValue("powershell.exe")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("InstallCommand")
                                    $XMLWriter.WriteValue("WorkFiles\Install-$ShortcutName.ps1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SwitchInstall")
                                    $XMLWriter.WriteValue("-ExecutionPolicy Bypass -File")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ArgsInstall")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("InstallMotor")
                                    $XMLWriter.WriteValue("powershell.exe")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("MSPList")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ProductVersion")
                                    $XMLWriter.WriteValue("1")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("Manufacturer")
                                    $XMLWriter.WriteValue($Vendor)
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("IgnoreExitCode")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("CustomExitCode")
                                    $XMLWriter.WriteValue("0")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SuppressDetectExitCodes")
                                    # $XMLWriter.WriteValue("")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("DetectionRules")
                                        $XMLWriter.WriteStartElement("DetectionRule")
                                            $XMLWriter.WriteStartElement("DetectionType")
                                            $XMLWriter.WriteValue("FileExists")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Path")
                                            $XMLWriter.WriteValue("%SystemDrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\$(Split-Path -Path $StartMenuContainer -Leaf)")
                                            <# if ($Target -like "http*") {
                                                $XMLWriter.WriteValue("%SystemDrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer")
                                            }
                                            else {
                                                $XMLWriter.WriteValue("%SystemDrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\$ShortcutName")
                                            } #>
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Name")
                                            $XMLWriter.WriteValue("$ShortcutName.lnk")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("Value")
                                            # $XMLWriter.WriteValue("")
                                            $XMLWriter.WriteEndElement()
                                            $XMLWriter.WriteStartElement("VersionOperand")
                                            $XMLWriter.WriteValue("IsEqual")
                                            $XMLWriter.WriteEndElement()
                                        $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("UninstallOrder")
                                    $XMLWriter.WriteValue("0")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("ReverseUsage")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("AlwaysUninstall")
                                    $XMLWriter.WriteValue("false")
                                    $XMLWriter.WriteEndElement()
                                    $XMLWriter.WriteStartElement("SupportedArchitecture")
                                    $XMLWriter.WriteValue("All")
                                    $XMLWriter.WriteEndElement()
                                $XMLWriter.WriteEndElement()
                            }
                        }
                        
        
                        if ($IconPath) {
                            Get-PSScriptInternetShortcut -Action Install -URL $Target -Name $ShortcutName -IconPath "Icon\Icon.ico" -OutputType ScriptFile
                        }
                        else {
                            Get-PSScriptInternetShortcut -Action Install -URL $Target -Name $ShortcutName -OutputType ScriptFile
                        }
                        
                        $TMPObject = @{
                            $InstCounter = [ordered]@{
                                ActionType = "AppPs1"
                                CmdType    = "Application"
                                index      = $InstCounter
                                uid        = $Guid
                                Text       = "Install $ShortcutName"
                                Type       = "Install"
                            }
                        }
        
                        $Script:OrderObject.Install += $TMPObject
                                            
                        $Script:InstCounter++
                    }
                    'Install/Uninstall Order' {
                        $XMLWriter.WriteStartElement("ActionRef")
                            $XMLWriter.WriteStartElement("ActionType")
                            $XMLWriter.WriteValue($InputObject.ActionType)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("CmdType")
                            $XMLWriter.WriteValue($InputObject.CmdType)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("index")
                            $XMLWriter.WriteValue($InstallOrderCounter)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("uid")
                            $XMLWriter.WriteValue($InputObject.uid)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Text")
                            $XMLWriter.WriteValue($InputObject.Text)
                            $XMLWriter.WriteEndElement()
                        $XMLWriter.WriteEndElement()
                    }
                    'Undo Create URL Action' {
                        Get-PSScriptInternetShortcut -Action Uninstall -Name $ShortcutName -OutputType ScriptFile
        
                        $TMPObject = @{
                            $UninstCounter = [ordered]@{
                                ActionType = "AppPs1"
                                CmdType    = "Application"
                                index      = $UninstCounter
                                uid        = $Guid
                                Text       = "Uninstall $ShortcutName"
                                Type       = "Uninstall"
                            }
                        }
        
                        $Script:OrderObject.Uninstall += $TMPObject
        
                        $Script:UninstCounter++
                    }
                    'Restore state file' {
                        $Guid = "{$(((New-Guid).Guid).ToUpper())}"
                        $XMLWriter.WriteStartElement("CommandAction")
                            $XMLWriter.WriteStartElement("uid")
                            $XMLWriter.WriteValue($Guid)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("type")
                            $XMLWriter.WriteValue("PsScript")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Cmd")
                            $XMLWriter.WriteValue($(Get-StateFileFix -Action Restore -Name $ShortcutName -OutputType Encoded))
                            Get-StateFileFix -Action Restore -Name $ShortcutName -OutputType ScriptFile
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("exitcode")
                            $XMLWriter.WriteValue("0")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("IgnoreExitCode")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Comment")
                            $XMLWriter.WriteValue("Restore state file")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("SupportedArchitecture")
                            $XMLWriter.WriteValue("All")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Version")
                            $XMLWriter.WriteValue("1,0,0,0")
                            $XMLWriter.WriteEndElement()
                        $XMLWriter.WriteEndElement()
        
                        $TMPObject = @{                        
                            $InstCounter = [ordered]@{
                                ActionType = "CommandAction"
                                CmdType    = "PsScript"
                                index      = $InstCounter
                                uid        = $Guid
                                Text       = "Restore state file"
                                Type       = "Install"
                            }
                        }
        
                        $Script:OrderObject.Install += $TMPObject
        
                        $Script:InstCounter++
                    }
                    'Restore state file Uninstall' {
                        $Name = $ZWXML.Bundle.Name.'#text'
                        $Guid = "{$(((New-Guid).Guid).ToUpper())}"
                        $XMLWriter.WriteStartElement("CommandAction")
                            $XMLWriter.WriteStartElement("uid")
                            $XMLWriter.WriteValue($Guid)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("type")
                            $XMLWriter.WriteValue("PsScript")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Cmd")
                            $XMLWriter.WriteValue($(Get-StateFileFix -Action Restore -Name $ShortcutName -OutputType Encoded))
                            Get-StateFileFix -Action Restore -Name $ShortcutName -OutputType ScriptFile
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("exitcode")
                            $XMLWriter.WriteValue("0")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("IgnoreExitCode")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Comment")
                            $XMLWriter.WriteValue("Restore state file")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("SupportedArchitecture")
                            $XMLWriter.WriteValue("All")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Version")
                            $XMLWriter.WriteValue("1,0,0,0")
                            $XMLWriter.WriteEndElement()
                        $XMLWriter.WriteEndElement()
        
                        $TMPObject = @{                        
                            $InstCounter = [ordered]@{
                                ActionType = "CommandAction"
                                CmdType    = "PsScript"
                                index      = $UninstCounter
                                uid        = $Guid
                                Text       = "Restore state file"
                                Type       = "Uninstall"
                            }
                        }
        
                        $Script:OrderObject.Uninstall += $TMPObject
        
                        $Script:UninstCounter++
                    }
                    'Rename state file' {
                        $Guid = "{$(((New-Guid).Guid).ToUpper())}"
                        $XMLWriter.WriteStartElement("CommandAction")
                            $XMLWriter.WriteStartElement("uid")
                            $XMLWriter.WriteValue($Guid)
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("type")
                            $XMLWriter.WriteValue("PsScript")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Cmd")
                            $XMLWriter.WriteValue($(Get-StateFileFix -Action Rename -Name $ShortcutName -OutputType Encoded))
                            Get-StateFileFix -Action Rename -Name $ShortcutName -OutputType ScriptFile
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("exitcode")
                            $XMLWriter.WriteValue("0")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("IgnoreExitCode")
                            $XMLWriter.WriteValue("false")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Comment")
                            $XMLWriter.WriteValue("Rename state file")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("SupportedArchitecture")
                            $XMLWriter.WriteValue("All")
                            $XMLWriter.WriteEndElement()
                            $XMLWriter.WriteStartElement("Version")
                            $XMLWriter.WriteValue("1,0,0,0")
                            $XMLWriter.WriteEndElement()
                        $XMLWriter.WriteEndElement()
        
                        $TMPObject = @{
                            $InstCounter = [ordered]@{
                                ActionType = "CommandAction"
                                CmdType    = "PsScript"
                                index      = $InstCounter
                                uid        = $Guid
                                Text       = "Rename state file"
                                Type       = "Install"
                            }
                        }
        
                        $Script:OrderObject.Install += $TMPObject
        
                        $Script:InstCounter++
                    }
                }
            }
            
            end {
                
            }
        }
        # End function.

        function ConvertTo-Icon {
            <#
            .Synopsis
                Converts .PNG images to icons
            .Description
                Converts a .PNG image to an icon
            .Example
                ConvertTo-Icon -Path .\Logo.png -Destination .\Favicon.ico
            #>
                [CmdletBinding()]
                param(
                # The file
                [Parameter(Mandatory=$true, Position=0,ValueFromPipelineByPropertyName=$true)]
                [Alias('Fullname','File')]
                [string]$Path,

                # If provided, will output the icon to a location
                [Parameter(Position=1, ValueFromPipelineByPropertyName=$true)]
                [Alias('OutputFile')]
                [string]$Destination
                )
                
                begin {
                    $TypeDefinition = @'
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Collections.Generic;
using System.Drawing.Drawing2D;

/// <summary>
/// Adapted from this gist: https://gist.github.com/darkfall/1656050
/// Provides helper methods for imaging
/// </summary>
public static class ImagingHelper
{
    /// <summary>
    /// Converts a PNG image to a icon (ico) with all the sizes windows likes
    /// </summary>
    /// <param name="inputBitmap">The input bitmap</param>
    /// <param name="output">The output stream</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Bitmap inputBitmap, Stream output)
    {
        if (inputBitmap == null)
            return false;

        int[] sizes = new int[] { 256, 48, 32, 16 };

        // Generate bitmaps for all the sizes and toss them in streams
        List<MemoryStream> imageStreams = new List<MemoryStream>();
        foreach (int size in sizes)
        {
            Bitmap newBitmap = ResizeImage(inputBitmap, size, size);
            if (newBitmap == null)
                return false;
            MemoryStream memoryStream = new MemoryStream();
            newBitmap.Save(memoryStream, ImageFormat.Png);
            imageStreams.Add(memoryStream);
        }

        BinaryWriter iconWriter = new BinaryWriter(output);
        if (output == null || iconWriter == null)
            return false;

        int offset = 0;

        // 0-1 reserved, 0
        iconWriter.Write((byte)0);
        iconWriter.Write((byte)0);

        // 2-3 image type, 1 = icon, 2 = cursor
        iconWriter.Write((short)1);

        // 4-5 number of images
        iconWriter.Write((short)sizes.Length);

        offset += 6 + (16 * sizes.Length);

        for (int i = 0; i < sizes.Length; i++)
        {
            // image entry 1
            // 0 image width
            iconWriter.Write((byte)sizes[i]);
            // 1 image height
            iconWriter.Write((byte)sizes[i]);

            // 2 number of colors
            iconWriter.Write((byte)0);

            // 3 reserved
            iconWriter.Write((byte)0);

            // 4-5 color planes
            iconWriter.Write((short)0);

            // 6-7 bits per pixel
            iconWriter.Write((short)32);

            // 8-11 size of image data
            iconWriter.Write((int)imageStreams[i].Length);

            // 12-15 offset of image data
            iconWriter.Write((int)offset);

            offset += (int)imageStreams[i].Length;
        }

        for (int i = 0; i < sizes.Length; i++)
        {
            // write image data
            // png data must contain the whole png data file
            iconWriter.Write(imageStreams[i].ToArray());
            imageStreams[i].Close();
        }

        iconWriter.Flush();

        return true;
    }

    /// <summary>
    /// Converts a PNG image to a icon (ico)
    /// </summary>
    /// <param name="input">The input stream</param>
    /// <param name="output">The output stream</param
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Stream input, Stream output)
    {
        Bitmap inputBitmap = (Bitmap)Bitmap.FromStream(input);
        return ConvertToIcon(inputBitmap, output);
    }

    /// <summary>
    /// Converts a PNG image to a icon (ico)
    /// </summary>
    /// <param name="inputPath">The input path</param>
    /// <param name="outputPath">The output path</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(string inputPath, string outputPath)
    {
        using (FileStream inputStream = new FileStream(inputPath, FileMode.Open))
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(inputStream, outputStream);
        }
    }



    /// <summary>
    /// Converts an image to a icon (ico)
    /// </summary>
    /// <param name="inputImage">The input image</param>
    /// <param name="outputPath">The output path</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Image inputImage, string outputPath)
    {
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(new Bitmap(inputImage), outputStream);
        }
    }


    /// <summary>
    /// Resize the image to the specified width and height.
    /// Found on stackoverflow: https://stackoverflow.com/questions/1922040/resize-an-image-c-sharp
    /// </summary>
    /// <param name="image">The image to resize.</param>
    /// <param name="width">The width to resize to.</param>
    /// <param name="height">The height to resize to.</param>
    /// <returns>The resized image.</returns>
    public static Bitmap ResizeImage(Image image, int width, int height)
    {
        var destRect = new Rectangle(0, 0, width, height);
        var destImage = new Bitmap(width, height);

        destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

        using (var graphics = Graphics.FromImage(destImage))
        {
            graphics.CompositingMode = CompositingMode.SourceCopy;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            using (var wrapMode = new ImageAttributes())
            {
                wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
            }
        }

        return destImage;
    }
}
'@
            
                    Add-Type -TypeDefinition $TypeDefinition -ReferencedAssemblies 'System.Drawing','System.IO','System.Collections'
                    if (-Not 'ImagingHelper' -as [Type]) {
                        Throw 'The custom "ImagingHelper" type is not loaded'
                    }
                }
                
                process {
                    #region Resolve Path
                    $ResolvedFile = $ExecutionContext.SessionState.Path.GetResolvedPSPathFromPSPath($Path)
                    if (-not $ResolvedFile) {
                        return
                    }
                    #endregion        
            
                    [ImagingHelper]::ConvertToIcon($ResolvedFile[0].Path,$Destination)
                }
        
                end {
            
                }
        }
        # End function.

        function Initialize-eKlientFolder {

            [CmdletBinding()]
        
            param (
                # Output type
                [Parameter()]
                [ValidateSet(
                    'Encoded',
                    'ScriptFile'
                )]
                [string]
                $OutputType,
        
                # Name of the script file
                [Parameter()]
                [string]
                $ScriptName
            )
            
            begin {
                $TheScript = @"
[CmdletBinding()]
Param([Parameter(Mandatory=`$True,Position=1)][string]`$RootFolder)

`$eKlientPath = "`$(`$env:SystemDrive)\eKlient"

if (-not (Test-Path -Path `$eKlientPath)) {
    try {
        New-Item -Path `$eKlientPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        `$eKlientFolder = Get-Item -Path `$eKlientPath -Force -ErrorAction Stop
        `$eKlientFolder.Attributes = 'Hidden'
    }
    catch {
        Write-Warning `$(`$_.Exception.Message)
    }
}
"@
            }
            
            process {
                $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($TheScript))
            }
            
            end {
                switch ($OutputType) {
                    Encoded    {return $Encoded}
                    ScriptFile {$TheScript | Set-Content ".\WorkFiles\New-eKlientFolder-$ScriptName.ps1" -NoNewline}
                }
            }
        }
        # End function.

        function Get-StateFileFix {
            [CmdletBinding()]
            param (
                # Type of action the script will preform (Install/Uninstall)
                [Parameter(Mandatory = $true)]
                [string]
                [ValidateSet(
                    'Rename',
                    'Restore'
                )]
                $Action,
        
                # Name of the URL shortcut
                [Parameter(Mandatory = $false)]
                [string]
                $Name,
        
                # Type of output
                [Parameter()]
                [ValidateSet(
                    'Encoded',
                    'ScriptFile'
                )]
                [string]
                $OutputType
            )
            
            begin {
                
            }
            
            process {
                switch ($Action) {
                    Rename   {
        
                    $TheScript = @"
[CmdletBinding()]
Param([Parameter(Mandatory=`$True,Position=1)][string]`$RootFolder)

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions external functions
. "`$(`$RootFolder)\WorkFiles\Get-LoggedOnUser.ps1"
Write-Output "External functions loaded: Get-LoggedOnUser"

# Fetch current username
`$Username = Get-LoggedOnUser -ReturnType Username

# Declare variable to current state file
`$StateFile = "`$(`$env:SystemDrive)\eKlient\$Name\$Name.state"

# If a state file exists, rename the state file to the current username
if (Test-Path -Path `$StateFile) {
    Rename-Item -Path `$StateFile -NewName "`$Username.state" -Force
    Write-Output "`$StateFile has been renamed to `$Username.state"
}
"@
                    }
                    Restore {
                        $TheScript = @"
[CmdletBinding()]
Param([Parameter(Mandatory=`$True,Position=1)][string]`$RootFolder)

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Get-LoggedOnUser.ps1"
Write-Output "External functions loaded: Get-LoggedOnUser"

# Fetch current username
`$Username = Get-LoggedOnUser -ReturnType Username

# Declare variable to current state file
`$StateFile = "`$(`$env:SystemDrive)\eKlient\$Name\`$Username.state"

# If a state file exists with the name of the current user, rename it to it's original name.
if (Test-Path -Path `$StateFile) {
    Rename-Item -Path `$StateFile -NewName "$Name.state" -Force
}
"@
                    }
                }
            }
            
            end {
                $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($TheScript))
                
                switch ($OutputType) {
                    Encoded    {return $Encoded}
                    ScriptFile {$TheScript | Set-Content ".\WorkFiles\$($Action)StateFile-$Name.ps1" -NoNewline}
                }
            }
        }
        # End function.

        function Get-PSScriptInternetShortcut {
            [CmdletBinding()]
            param (
                # Type of action the script will preform (Install/Uninstall)
                [Parameter(Mandatory = $true)]
                [string]
                [ValidateSet(
                    'Install',
                    'Uninstall'
                )]
                $Action,
        
                # URL to be opened
                [Parameter(Mandatory = $false)]
                [string]
                $URL,
        
                # Name of the URL shortcut
                [Parameter(Mandatory = $false)]
                [string]
                $Name,
        
                # Path to icon
                [Parameter(Mandatory = $false)]
                [string]
                $IconPath,
        
                # Type of output
                [Parameter()]
                [ValidateSet(
                    'Encoded',
                    'ScriptFile'
                )]
                [string]
                $OutputType
            )
            
            begin {
                switch ($ReturnType) {
                    User {
                        switch ($Action) {
                            Install   {
                                if ($IconPath) {
                            $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"
. "`$(`$RootFolder)\WorkFiles\Get-LoggedOnUser.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Install"
    Name         = "NAMEREPLACE"
    Username     = `$(Get-LoggedOnUser -ReturnType Username)
    Target       = "URLREPLACE"
    IconPath     = "ICONPATHREPLACE"
    LocalContent = "`$(`$env:SystemDrive)\eKlient\NAMEREPLACE"
    Verbose      = `$true
}

# Create the shortcut under current user profile
Set-eKlientWindowsShortcut @ShortcutParams
"@
                                }
                                else {
                            $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"
. "`$(`$RootFolder)\WorkFiles\Get-LoggedOnUser.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Uninstall"
    Name         = "NAMEREPLACE"
    Username     = `$(Get-LoggedOnUser -ReturnType Username)
    Target       = "URLREPLACE"
    LocalContent = "`$(`$env:SystemDrive)\eKlient\NAMEREPLACE"
    Verbose      = `$true
}

# Create the shortcut under current user profile
Set-eKlientWindowsShortcut @ShortcutParams
"@
                                }
                            }
                            Uninstall {
                                $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"
. "`$(`$RootFolder)\WorkFiles\Get-LoggedOnUser.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Uninstall"
    Name         = "$Name"
    Username     = `$(Get-LoggedOnUser -ReturnType Username)
    LocalContent = "`$(`$env:SystemDrive)\eKlient\$Name"
}

# Remove the shortcut
Set-eKlientWindowsShortcut @ShortcutParams
"@
                            }
                        }
                    }
                    Computer {
                        switch ($Action) {
                            Install   {
                                if ($IconPath) {
                            $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Install"
    Name         = "NAMEREPLACE"
    Target       = "URLREPLACE"
    IconPath     = "ICONPATHREPLACE"
    LocalContent = "`$(`$env:SystemDrive)\eKlient\NAMEREPLACE"
    Verbose      = `$true
}

# Create the shortcut under current user profile
Set-eKlientWindowsShortcut @ShortcutParams
"@
                                }
                                else {
                            $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Uninstall"
    Name         = "NAMEREPLACE"
    Target       = "URLREPLACE"
    LocalContent = "`$(`$env:SystemDrive)\eKlient\NAMEREPLACE"
    Verbose      = `$true
}

# Create the shortcut under current user profile
Set-eKlientWindowsShortcut @ShortcutParams
"@
                                }
                            }
                            Uninstall {
                                $TheScript = @"
# Determine root folder
`$RootFolder = `$PSScriptRoot | Split-Path

# Temporarily allow scripts to run
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# Loading functions
. "`$(`$RootFolder)\WorkFiles\Set-eKlientWindowsShortcut.ps1"

# Build parameter object for the Set-eKlientWindowsShortcut function.
`$ShortcutParams = @{
    Action       = "Uninstall"
    Name         = "$Name"
    LocalContent = "`$(`$env:SystemDrive)\eKlient\$Name"
}

# Remove the shortcut
Set-eKlientWindowsShortcut @ShortcutParams
"@
                            }
                        }
                    }
                }
            }
            
            process {
        
                switch ($Action) {
                    Install   {
                        $TheScript = $TheScript -replace "URLREPLACE", "$URL"
                        $TheScript = $TheScript -replace "NAMEREPLACE", "$Name"
                        if ($IconPath) {
                            $TheScript = $TheScript -replace "ICONPATHREPLACE", "`$(`$RootFolder)\$IconPath"
                        }
                    }
                    Uninstall {}
                }
        
                $Encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($TheScript))
            }
            
            end {
                switch ($OutputType) {
                    Encoded    {return $Encoded}
                    ScriptFile {$TheScript | Set-Content ".\WorkFiles\$Action-$Name.ps1" -NoNewline}
                }
            }
        }
        # End function.

        function New-ApplicationFolder {

            [CmdletBinding()]
        
            param (
                # Path of the application folder
                [Parameter()]
                [string]
                $Path
            )
            
            begin {
                $SubFolders = @(
                    'AcceptanceTest',
                    'FinalMedia',
                    'Icon',
                    'OriginalMedia',
                    'QA',
                    'WorkFiles'
                )
            }
            
            process {
        
                if (-not (Test-Path $Path)) {
                    try {
                        New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
        
                        foreach ($Folder in $SubFolders) {
                            New-Item -Path "$Path\$Folder" -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
        
                    }
                    catch {
                        Write-Warning -Message "New-ApplicationFolder - Error at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
                    }
                }
                else {
                    foreach ($Folder in $SubFolders) {
                        if (-not (Test-Path "$Path\$Folder")) {
                            try {
                                New-Item -Path "$Path\$Folder" -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            }
                            catch {
                                Write-Warning -Message "New-ApplicationFolder - Error at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
                            }
                        }
                    }
                }
            }
            
            end {
                if ($Script:IsWebApp) {
                    $Files  = @(
                        "$env:OneDriveCommercial\_POSH\_Functions\eKlientScriptRepo\Set-eKlientWindowsShortcut.ps1",
                        "$env:OneDriveCommercial\_POSH\_Functions\eKlientScriptRepo\Get-LoggedOnUser.ps1"
                    )
        
                    foreach ($File in $Files) {
                        Copy-Item -Path $File -Destination "$Path\WorkFiles"
                    }
                }
            }
        }
        # End function.

        function Export-ExternalFunctions {
            [CmdletBinding()]
            param (
                
            )
            
            begin {
                
            }
            
            process {
                $Get_LoggedOnUser = @"
function Get-LoggedOnUser {

    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet(
            'Username',
            'SID'
        )]
        [string]
        `$ReturnType
    )

    function Get-ActiveUser {
        `$Computer = `$env:COMPUTERNAME
        `$Users = query user /server:`$Computer 2>&1

        `$Users = `$Users | ForEach-Object {
            ((`$_.trim() -replace ">" -replace "(?m)^([A-Za-z0-9]{3,})\s+(\d{1,2}\s+\w+)", '`$1  none  `$2' -replace "\s{2,}", "," -replace "none", `$null))
        } | ConvertFrom-Csv

        `$AllUsers = foreach (`$User in `$Users) {
            [PSCustomObject]@{
                ComputerName = `$Computer
                Username = `$User.USERNAME
                SessionState = `$User.STATE.Replace("Disc", "Disconnected")
                SessionType = `$(`$User.SESSIONNAME -Replace '#', '' -Replace "[0-9]+", "")
            }
        }

        `$RunningUser = `$AllUsers | Where-Object {`$_.SessionState -eq "Active"} | Select-Object -ExpandProperty Username

        return `$RunningUser
    }
    # End function.
    
    `$ExplorerProcess = Get-WmiObject -class win32_process  | Where-Object {`$_.ProcessName -eq "explorer.exe"}
    
    if (`$null -eq `$ExplorerProcess) {
        `$LoggedOnUser = `$false
        Write-Warning "Explorer process not running."
        break
    }
    elseif (`$ExplorerProcess.getowner().user.count -gt 1) {
        `$LoggedOnUser = Get-ActiveUser
    }
    else {
        `$LoggedOnUser = `$ExplorerProcess.getowner().user
    }

    switch (`$ReturnType) {
        Username {return `$LoggedOnUser}
        SID {
            `$Domain = (Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select -ExpandProperty Domain).split('.')[0]
            `$SID = (New-Object Security.Principal.NTAccount(`$Domain, `$LoggedOnUser)).Translate([Security.Principal.SecurityIdentifier]).Value
            return `$SID
        }
    }
    return `$LoggedOnUser
}
# End function.
"@
                $Set_eKlientWindowsShortcut = switch ($ReturnType) {
                    User {
                        @"
function Set-eKlientWindowsShortcut {
    <#
    .SYNOPSIS
    Create a windows shortcut
    
    .DESCRIPTION
    Helper function to create a windows shortcut
    
    .PARAMETER Action
    Switch with possible values: Install or Uninstall
    
    .PARAMETER Name
    Name of the shortcut
    
    .PARAMETER Username
    Used in conjuntion with installation in user profile path
    
    .PARAMETER Target
    Where the shortcut is suppossed to point to. Url or file path
    
    .PARAMETER IconPath
    If an icon is to be used, provide path here
    
    .PARAMETER LocalContent
    Local path to store the icon.
    
    .EXAMPLE
    Set-eKlientWindowsShortcut -Action Install -Username test01 -Target "https://google.com" -IconPath "C:\Temp\Icon.ico" -LocalContent "`$env:SystemDrive\eKlient\AppName"
    
    .NOTES
    Author: Simon Mellergård | It-center, Värnamo kommun.
    Version: 0.0.0.1 - 2024-03-06
    #>
        [CmdletBinding()]
    
        param (
            # Install the shortcut or uninstall it
            [Parameter()]
            [ValidateSet(
                'Install',
                'Uninstall'
            )]
            [string]
            `$Action,
    
            # Name of the shortcut
            [Parameter()]
            [string]
            `$Name,
    
            # Username for whom the shortcut will be installed.
            [Parameter()]
            [string]
            `$Username,
    
            # Target of the shortcut
            [Parameter()]
            [string]
            `$Target,
    
            # Path to icon
            [Parameter()]
            [string]
            `$IconPath,
    
            # Path to store local content, typically "`$env:SystemDrive\eKlient\AppName"
            [Parameter()]
            [string]
            `$LocalContent
        )
        
        begin {
            # Declare root variable of where the shortcut will be stored.
            `$Root = "`$env:SystemDrive\Users\`$Username\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"
        }
        
        process {
            # Switch case running either installation or uninstallation steps.
            switch (`$Action) {
                Install   {
                    # If root folder doesn't exist, it will be created.
                    if (-not (Test-Path -Path `$Root)) {
                        New-Item -Path `$Root -ItemType Directory -Force | Out-Null
                    }
                    # If local content folder dosen't exist, it will be created.
                    if (-not (Test-Path -Path `$LocalContent)) {
                        New-Item -Path `$LocalContent -ItemType Directory -Force | Out-Null
                    }
    
                    # Create a state file for keeping track of the installation.
                    New-Item -Path `$LocalContent -Name "`$(`$Name).state" -ItemType File -Value "`$(`$Name) installed `$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
    
                    if (`$IconPath) {
                        # Trying to copy provided icon file to the local content path
                        try {
                            Copy-Item -Path `$IconPath -Destination `$LocalContent -Force -ErrorAction Stop
                        }
                        catch {
                            Write-Warning `$(`$_.Exception.Message)
                        }
    
                        # If shortcut doesn't exist, it will be created in the root folder based on provided information.
                        if (-not (Test-Path "`$Root\`$Name.lnk")) {
                            `$WshShell                  = New-Object -ComObject WScript.Shell
                            `$Shortcut                  = `$WshShell.CreateShortcut("`$Root\`$Name.lnk")
                            `$Shortcut.TargetPath       = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
                            `$Shortcut.Arguments        = `$Target
                            `$Shortcut.WorkingDirectory = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application"
                            `$Shortcut.IconLocation     = "`$LocalContent\Icon.ico"
                            `$Shortcut.Save()
                        
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$Shortcut) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$WshShell) | Out-Null
                            [System.GC]::Collect()
                            [System.GC]::WaitForPendingFinalizers()
                        }
                    }
                    else {
    
                        # If an icon is not provided, the shortcut will be created in the root folder without an icon connection.
                        if (-not (Test-Path "`$Root\`$Name.lnk")) {
                
                            `$WshShell                  = New-Object -ComObject WScript.Shell
                            `$Shortcut                  = `$WshShell.CreateShortcut("`$Root\`$Name.lnk")
                            `$Shortcut.TargetPath       = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
                            `$Shortcut.Arguments        = `$Target
                            `$Shortcut.WorkingDirectory = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application"
                            `$Shortcut.Save()
                        
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$Shortcut) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$WshShell) | Out-Null
                            [System.GC]::Collect()
                            [System.GC]::WaitForPendingFinalizers()
                        }
                    }
    
                    Write-Output "`$Name.lnk has been created under `$Root"
                }
                Uninstall {
                    # If a state file is present in the local content folder, it will be deleted.
                    if (Test-Path "`$(`$LocalContent)\`$(`$Name).state") {
                        Remove-Item -Path "`$(`$LocalContent)\`$(`$Name).state" -Force
                        Write-Output "State file removed."
                    }
                    # If a shortcut is present in the root folder, it will be deleted.
                    if (Test-Path "`$Root\`$Name.lnk") {
                        Remove-Item -Path "`$Root\`$Name.lnk" -Force
                        Write-Output "Link in user profile start menu removed."
                    }
                    # If a local content path is present, it will be deleted.
                    if (Test-Path `$LocalContent) {
                        if (-not (`$(Get-ChildItem -Path `$LocalContent -Filter "*.state").Count -ge 1)) {
                            Remove-Item `$LocalContent -Recurse -Force
                            Write-Output "Local content removed."
                        }
                    }
                }
            }
        }
        
        end {
            
        }
    }
    # End function.
"@
                    }
                    Computer {
                        @"
function Set-eKlientWindowsShortcut {
    <#
    .SYNOPSIS
    Create a windows shortcut
    
    .DESCRIPTION
    Helper function to create a windows shortcut
    
    .PARAMETER Action
    Switch with possible values: Install or Uninstall
    
    .PARAMETER Name
    Name of the shortcut
    
    .PARAMETER Target
    Where the shortcut is suppossed to point to. Url or file path
    
    .PARAMETER IconPath
    If an icon is to be used, provide path here
    
    .PARAMETER LocalContent
    Local path to store the icon.
    
    .EXAMPLE
    Set-eKlientWindowsShortcut -Action Install -Target "https://google.com" -IconPath "C:\Temp\Icon.ico" -LocalContent "`$env:SystemDrive\eKlient\AppName"
    
    .NOTES
    Author: Simon Mellergård | It-center, Värnamo kommun.
    Version: 0.0.0.2 - 2024-03-21
    #>
        [CmdletBinding()]
    
        param (
            # Install the shortcut or uninstall it
            [Parameter()]
            [ValidateSet(
                'Install',
                'Uninstall'
            )]
            [string]
            `$Action,
    
            # Name of the shortcut
            [Parameter()]
            [string]
            `$Name,

            # Start menu container name
            [Parameter()]
            [string]
            `$Root = "$StartMenuContainer",
    
            # Target of the shortcut
            [Parameter()]
            [string]
            `$Target,
    
            # Path to icon
            [Parameter()]
            [string]
            `$IconPath,
    
            # Path to store local content, typically "`$env:SystemDrive\eKlient\AppName"
            [Parameter()]
            [string]
            `$LocalContent
        )
        
        begin {

            # Determine wheter this will be a web application shortcut or a UNC shortcut.
            if (`$Target -like "https://*" -or `$Target -like "http://*") {
                # Declare root variable of where the shortcut will be stored.
                # `$Root = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\_Webbapplikationer"
                `$TargetPath = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
                `$WorkingDir = "`${env:ProgramFiles(x86)}\Microsoft\Edge\Application"
            }
            elseif (`$Target -like "*.ps1*") {
                # `$Root = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\`$Name"
                `$TargetPath = "`$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
                `$WorkingDir = `$(`$Target | Split-Path)
                `$Target   = "-ExecutionPolicy ByPass -WindowStyle Hidden -File ""`$(`$Target | Split-Path -Leaf)"""
            }
            else {
                # `$Root = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\`$Name"
                `$TargetPath = `$Target
                `$WorkingDir = `$(`$Target | Split-Path)
            }
        }
        
        process {
            # Switch case running either installation or uninstallation steps.
            switch (`$Action) {
                Install   {
                    # If root folder doesn't exist, it will be created.
                    if (-not (Test-Path -Path `$Root)) {
                        New-Item -Path `$Root -ItemType Directory -Force | Out-Null
                    }
                    # If local content folder dosen't exist, it will be created.
                    if (-not (Test-Path -Path `$LocalContent)) {
                        New-Item -Path `$LocalContent -ItemType Directory -Force | Out-Null
                    }

                    if (`$IconPath) {
                        # Trying to copy provided icon file to the local content path
                        try {
                            Copy-Item -Path `$IconPath -Destination `$LocalContent -Force -ErrorAction Stop
                        }
                        catch {
                            Write-Warning `$(`$_.Exception.Message)
                        }
    
                        # If shortcut doesn't exist, it will be created in the root folder based on provided information.
                        if (-not (Test-Path "`$Root\`$Name.lnk")) {
                            `$WshShell                  = New-Object -ComObject WScript.Shell
                            `$Shortcut                  = `$WshShell.CreateShortcut("`$Root\`$Name.lnk")
                            if (`$Target -like "*.exe*") {
                                `$Shortcut.TargetPath       = `$Target
                            }
                            else {
                                `$Shortcut.TargetPath       = `$TargetPath
                                `$Shortcut.Arguments        = `$Target
                            }
                            `$Shortcut.WorkingDirectory = `$WorkingDir
                            `$Shortcut.IconLocation     = "`$LocalContent\Icon.ico"
                            `$Shortcut.Save()
                        
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$Shortcut) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$WshShell) | Out-Null
                            [System.GC]::Collect()
                            [System.GC]::WaitForPendingFinalizers()
                        }
                    }
                    else {
    
                        # If an icon is not provided, the shortcut will be created in the root folder without an icon connection.
                        if (-not (Test-Path "`$Root\`$Name.lnk")) {
                
                            `$WshShell                  = New-Object -ComObject WScript.Shell
                            `$Shortcut                  = `$WshShell.CreateShortcut("`$Root\`$Name.lnk")
                            if (`$Target -like "*.exe*") {
                                `$Shortcut.TargetPath       = `$Target
                            }
                            else {
                                `$Shortcut.TargetPath       = `$TargetPath
                                `$Shortcut.Arguments        = `$Target
                            }
                            `$Shortcut.WorkingDirectory = `$WorkingDir
                            `$Shortcut.Save()
                        
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$Shortcut) | Out-Null
                            [System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$WshShell) | Out-Null
                            [System.GC]::Collect()
                            [System.GC]::WaitForPendingFinalizers()
                        }
                    }
    
                    Write-Output "`$Name.lnk has been created under `$Root"
                }
                Uninstall {
                    # If a shortcut is present in the root folder, it will be deleted.
                    if (Test-Path "`$Root\`$Name.lnk") {
                        Remove-Item -Path "`$Root\`$Name.lnk" -Force
                        Write-Output "Link in user profile start menu removed."

                        if (`$(Get-ChildItem -Path `$Root).Count -eq 0) {
                            Remove-Item -Path `$Root -Recurse -Force
                            Write-Output "Start menu container for `$Name has been removed."
                        }
                    }
                    # If a local content path is present, it will be deleted.
                    if (Test-Path `$LocalContent) {
                        Remove-Item `$LocalContent -Recurse -Force
                        Write-Output "Local content removed."
                    }
                }
            }
        }
        
        end {
            
        }
    }
    # End function.
"@
                    }
                }
            }
            
            end {
                switch ($ReturnType) {
                    User {
                        $Get_LoggedOnUser | Set-Content ".\WorkFiles\Get-LoggedOnUser.ps1" -NoNewline
                        $Set_eKlientWindowsShortcut | Set-Content ".\WorkFiles\Set-eKlientWindowsShortcut.ps1" -NoNewline
                    }
                    Computer {
                        $Set_eKlientWindowsShortcut | Set-Content ".\WorkFiles\Set-eKlientWindowsShortcut.ps1" -NoNewline
                    }
                }
            }
        }
        # End function.
        #endregion Declare internal functions

        # Get the command name
        $CommandName = $PSCmdlet.MyInvocation.InvocationName
        # Get the list of parameters for the command
        $ParameterList = (Get-Command -Name $CommandName).Parameters

        # Grab each parameter value, using Get-Variable
        $InfoTable = foreach ($Parameter in $ParameterList) {
            Get-Variable -Name $Parameter.Values.Name -ErrorAction SilentlyContinue
        }

        # Constructing resulting full path
        $FullPath = "$OutPath\$AppName"

        # Generate the application folder structure
        New-ApplicationFolder -Path $FullPath
        Set-Location -Path $FullPath

        # Copying InstallKing from WrapperKing installation directory to Wrapper package directory
        try {
            Copy-Item -Path $InstallKing -Destination $FullPath -ErrorAction Stop
        }
        catch {
            Write-Error $_.Exception.Message
            exit 1
        }

        # Identifying if there is an icon present and copy it to the newly created folder structure
        if (Test-Path -Path $IconPath) {
            $Icon = Get-ChildItem -Path $IconPath
            if ($Icon.FullName -notlike '*.ico') {
                ConvertTo-Icon -Path $Icon.FullName -Destination "$FullPath\Icon\Icon.ico"
            }
            else {
                Copy-Item -Path $IconPath -Destination "$FullPath\Icon\Icon.ico"
            }
        }

        #region Building the XML properties

        # Encoding
        # $Encoding      = [System.Text.Encoding]::GetEncoding(65001) # 65001 utf-8 Unicode (UTF-8)
        $Encoding      = [System.Text.UTF8Encoding]::new($false)
        $Builder       = New-Object -TypeName System.Text.StringBuilder
        $StringBuilder = New-Object -TypeName System.IO.StringWriter($Builder)

        # Settings
        $Settings                 = New-Object -TypeName System.Xml.XmlWriterSettings
        $Settings.Encoding        = $Encoding
        $Settings.Indent          = $true
        $Settings.CloseOutput     = $false
        $Settings.CheckCharacters = $true

        # Create the XML file
        $WrapperKingXML   = "$FullPath\WrapperKing.xml"
        $Script:XMLWriter = [System.Xml.XmlWriter]::Create($WrapperKingXML, $Settings)

        #endregion Building the XML properties

        #region Building Orderobject

        if (-not ($OrderObject)) {
            # Write-Host "Skapar räknare och OrderObject!" -ForegroundColor Green
            $Script:InstCounter   = 1
            $Script:UninstCounter = 1
            $Script:OrderObject = @{
                Install   = @{}
                Uninstall = @{}
            }
        }

        #endregion Building Orderobject
    }
    
    process {
        # Write the starting element together with schemas.
        $XMLWriter.WriteStartDocument()
        $XMLWriter.WriteStartElement("WrapperSettings")
        $XmlWriter.WriteAttributeString("xmlns","xsd", $null, "http://www.w3.org/2001/XMLSchema")
        $XmlWriter.WriteAttributeString("xmlns","xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")

        #region Writing all child elements

        #region Note to show prior to installation
        $XMLWriter.WriteStartElement("PreNoteInst")
            $XMLWriter.WriteStartElement("type")
            $XMLWriter.WriteValue("PreNotifierInstall")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PreBehavior")
            $XMLWriter.WriteValue("ForceAction")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PostBehavior")
            $XMLWriter.WriteValue("Restart")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("CountDownTime")
            $XMLWriter.WriteValue("3600")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RemindInterval")
            $XMLWriter.WriteValue("900")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("FrameColor")
            $XMLWriter.WriteValue("#2DAFE6")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("UserSoftwareCenterBranding")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessFilter")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessSettings")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ShowOnProcsOnly")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AutoKillProcsOnCountdownEnd")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Height")
            $XMLWriter.WriteValue("350")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Width")
            $XMLWriter.WriteValue("450")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AllowMove")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Note to show prior to installation

        #region Note to show prior to uninstallation
        $XMLWriter.WriteStartElement("PreNoteUninst")
            $XMLWriter.WriteStartElement("type")
            $XMLWriter.WriteValue("PreNotifierInstall")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PreBehavior")
            $XMLWriter.WriteValue("ForceAction")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PostBehavior")
            $XMLWriter.WriteValue("Restart")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("CountDownTime")
            $XMLWriter.WriteValue("3600")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RemindInterval")
            $XMLWriter.WriteValue("900")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("FrameColor")
            $XMLWriter.WriteValue("#2DAFE6")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("UserSoftwareCenterBranding")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessFilter")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessSettings")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ShowOnProcsOnly")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AutoKillProcsOnCountdownEnd")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Height")
            $XMLWriter.WriteValue("350")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Width")
            $XMLWriter.WriteValue("450")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AllowMove")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Note to show prior to uninstallation

        #region Note to show after installation
        $XMLWriter.WriteStartElement("PostNoteInst")
            $XMLWriter.WriteStartElement("type")
            $XMLWriter.WriteValue("PreNotifierInstall")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PreBehavior")
            $XMLWriter.WriteValue("ForceAction")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PostBehavior")
            $XMLWriter.WriteValue("Restart")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("CountDownTime")
            $XMLWriter.WriteValue("3600")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RemindInterval")
            $XMLWriter.WriteValue("900")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("FrameColor")
            $XMLWriter.WriteValue("#2DAFE6")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("UserSoftwareCenterBranding")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessFilter")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessSettings")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ShowOnProcsOnly")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AutoKillProcsOnCountdownEnd")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Height")
            $XMLWriter.WriteValue("350")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Width")
            $XMLWriter.WriteValue("450")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AllowMove")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Note to show after installation

        #region Note to show after uninstallation
        $XMLWriter.WriteStartElement("PostNoteUninst")
            $XMLWriter.WriteStartElement("type")
            $XMLWriter.WriteValue("PreNotifierInstall")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PreBehavior")
            $XMLWriter.WriteValue("ForceAction")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("PostBehavior")
            $XMLWriter.WriteValue("Restart")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("CountDownTime")
            $XMLWriter.WriteValue("3600")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RemindInterval")
            $XMLWriter.WriteValue("900")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("FrameColor")
            $XMLWriter.WriteValue("#2DAFE6")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("UserSoftwareCenterBranding")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessFilter")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ProcessSettings")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("ShowOnProcsOnly")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AutoKillProcsOnCountdownEnd")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Height")
            $XMLWriter.WriteValue("350")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Width")
            $XMLWriter.WriteValue("450")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AllowMove")
            $XMLWriter.WriteValue("false")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Note to show after uninstallation

        #region Generate list of commands to be included in the xml
        $XMLWriter.WriteStartElement("Commands")

        switch ($ReturnType) {
            User     {
                # Create step and powershell script that will check for state file of running user and restore it to it's former name for the detection rule to properly work.
                Format-XMLElement -Type 'Restore state file'
                
                # Create step and powershell script that will check if local content folder has been created and create it if it hasn't.
                Format-XMLElement -Type 'Initialize eKlient Folder'
                
                # Create step and powershell script that will restore the state file from the running user to it's former name in the uninstallation step.
                Format-XMLElement -Type 'Restore state file Uninstall'

                # Create step and powershell script that will rename the state file after the shortcut has been created.
                Format-XMLElement -Type 'Rename state file'
            }
            Computer {
                # Create step and powershell script that will check if local content folder has been created and create it if it hasn't.
                Format-XMLElement -Type 'Initialize eKlient Folder'
            }
        }
        
        
        $XMLWriter.WriteEndElement()
        #endregion Generate list of commands to be included in the xml

        #region Building the apps section
        $XMLWriter.WriteStartElement("Apps")
        # Create the install app step
        Format-XMLElement -Type 'Launch URL Action'
        # Create the uninstall step
        Format-XMLElement -Type 'Undo Create URL Action'
        $XMLWriter.WriteEndElement()
        #endregion Building the apps section

        #region Building the General section
        $XMLWriter.WriteStartElement("General")
        Format-XMLElement -Type General
        $XMLWriter.WriteEndElement()
        #endregion Building the General section

        #region Building the InstallOrder section
        $XMLWriter.WriteStartElement("InstallOrder")
            $Script:InstallOrderCounter = 1
            switch ($ReturnType) {
                User {
                    if ($OrderObject.Install.Values | Where-Object {($_.CmdType -eq "PsScript") -or ($_.ActionType -eq "AppPs1") -and ($_.Text -ne 'Rename state file')}) {

                        $FinalOrder = $OrderObject.Install.Values | Where-Object {($_.CmdType -eq "PsScript") -or ($_.ActionType -eq "AppPs1") -and ($_.Text -ne 'Rename state file')}
                        $Numbers = $FinalOrder.index | Sort-Object

                        foreach ($Num in $Numbers) {
                            Format-XMLElement -InputObject $($FinalOrder | Where-Object {$_.index -eq $Num}) -Type 'Install/Uninstall Order'
                            $Script:InstallOrderCounter++
                        }

                        $RealFinal = $OrderObject.Install.Values | Where-Object {$_.Text -eq "Rename state file"}
                        Format-XMLElement -InputObject $RealFinal -Type 'Install/Uninstall Order'
                        $Script:InstallOrderCounter++
                    }
                }
                Computer {
                    Format-XMLElement -InputObject $($OrderObject.Values.Values | Where-Object {$_.CmdType -eq "PsScript"}) -Type 'Install/Uninstall Order'
                    $Script:InstallOrderCounter++
                    Format-XMLElement -InputObject $($OrderObject.Values.Values | Where-Object {($_.ActionType -eq "AppPs1") -and ($_.Type -eq "Install")}) -Type 'Install/Uninstall Order'
                }
            }
        $XMLWriter.WriteEndElement()
        #endregion Building the InstallOrder section

        #region Building the UninstallOrder section
        $XMLWriter.WriteStartElement("UninstallOrder")
            $Script:InstallOrderCounter = 1
            switch ($ReturnType) {
                User {
                    foreach ($UninstallEntry in $($OrderObject.Uninstall.Values | Sort-Object -Property Name)) {
                        Format-XMLElement -InputObject $UninstallEntry -Type 'Install/Uninstall Order'
                        $Script:InstallOrderCounter++
                    }
                }
                Computer {
                    Format-XMLElement -InputObject $OrderObject.Uninstall.Values -Type 'Install/Uninstall Order'
                }
            }
        $XMLWriter.WriteEndElement()
        #endregion Building the UninstallOrder section
        
        #region Building the InstallProgSettings section
        $XMLWriter.WriteStartElement("InstallProgSettings")
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Title")
            $XMLWriter.WriteValue("Installing $ShortcutName")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Position")
            $XMLWriter.WriteValue("LowerRight")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Size")
            $XMLWriter.WriteValue("Normal")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Background")
            $XMLWriter.WriteValue("#A3BD6A")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Foreground")
            $XMLWriter.WriteValue("#000000")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RunTimeSeconds")
            $XMLWriter.WriteValue("0")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("LogoFile")
            $XMLWriter.WriteValue("Icon/Icon.ico")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AlwaysOnTop")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Building the InstallProgSettings section

        #region Building the UninstallProgSettings section
        $XMLWriter.WriteStartElement("UninstallProgSettings")
            $XMLWriter.WriteStartElement("Active")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Title")
            $XMLWriter.WriteValue("Uninstalling $ShortcutName")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Position")
            $XMLWriter.WriteValue("LowerRight")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Size")
            $XMLWriter.WriteValue("Normal")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Background")
            $XMLWriter.WriteValue("#A3BD6A")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("Foreground")
            $XMLWriter.WriteValue("#000000")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("RunTimeSeconds")
            $XMLWriter.WriteValue("0")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("LogoFile")
            $XMLWriter.WriteValue("Icon/Icon.ico")
            $XMLWriter.WriteEndElement()
            $XMLWriter.WriteStartElement("AlwaysOnTop")
            $XMLWriter.WriteValue("true")
            $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndElement()
        #endregion Building the UninstallProgSettings section

        # Closing starting element
        $XMLWriter.WriteEndElement()
        $XMLWriter.WriteEndDocument()
        $XMLWriter.Flush()
        $XMLWriter.Close()
        #endregion Writing all child elements
    }
    
    end {

        Export-ExternalFunctions
        Remove-Variable OrderObject -Scope Script
        return "Wrapper package successfully created in $FullPath"
    }
}
# End function.

Set-Location -Path $PSScriptRoot

$PathToInstallKing = "$(${env:ProgramFiles(x86)})\eKlient\WrapperKing 3\InstallKing.exe"

if (-not (Test-Path -Path $PathToInstallKing)) {
    Get-MessageBox -Type Ok -Icon Stop -Body "Failed to locate InstallKing.exe.`nPlease make sure WrapperKing is installed before running this application." -Title "WrapperKing not installed"
}
else {
    Initialize-WPF
    Show-NewWKShortcut
}