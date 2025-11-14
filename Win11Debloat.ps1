#Requires -RunAsAdministrator

[CmdletBinding(SupportsShouldProcess)]
param (
    [switch]$Silent,
    [switch]$Sysprep,
    [string]$LogPath,
    [string]$User,
    [switch]$CreateRestorePoint,
    [switch]$RunAppsListGenerator, [switch]$RunAppConfigurator,
    [switch]$RunDefaults,
    [switch]$RunDefaultsLite,
    [switch]$RunSavedSettings,
    [switch]$RemoveApps, 
    [switch]$RemoveAppsCustom,
    [switch]$RemoveGamingApps,
    [switch]$RemoveCommApps,
    [switch]$RemoveHPApps,
    [switch]$RemoveW11Outlook,
    [switch]$ForceRemoveEdge,
    [switch]$DisableDVR,
    [switch]$DisableTelemetry,
    [switch]$DisableFastStartup,
    [switch]$DisableModernStandbyNetworking,
    [switch]$DisableBingSearches, [switch]$DisableBing,
    [switch]$DisableDesktopSpotlight,
    [switch]$DisableLockscrTips, [switch]$DisableLockscreenTips,
    [switch]$DisableWindowsSuggestions, [switch]$DisableSuggestions,
    [switch]$DisableEdgeAds,
    [switch]$DisableSettings365Ads,
    [switch]$DisableSettingsHome,
    [switch]$ShowHiddenFolders,
    [switch]$ShowKnownFileExt,
    [switch]$HideDupliDrive,
    [switch]$EnableDarkMode,
    [switch]$DisableTransparency,
    [switch]$DisableAnimations,
    [switch]$TaskbarAlignLeft,
    [switch]$CombineTaskbarAlways, [switch]$CombineTaskbarWhenFull, [switch]$CombineTaskbarNever,
    [switch]$CombineMMTaskbarAlways, [switch]$CombineMMTaskbarWhenFull, [switch]$CombineMMTaskbarNever,
    [switch]$MMTaskbarModeAll, [switch]$MMTaskbarModeMainActive, [switch]$MMTaskbarModeActive,
    [switch]$HideSearchTb, [switch]$ShowSearchIconTb, [switch]$ShowSearchLabelTb, [switch]$ShowSearchBoxTb,
    [switch]$HideTaskview,
    [switch]$DisableStartRecommended,
    [switch]$DisableStartPhoneLink,
    [switch]$DisableCopilot,
    [switch]$DisableRecall,
    [switch]$DisableClickToDo,
    [switch]$DisablePaintAI,
    [switch]$DisableNotepadAI,
    [switch]$DisableEdgeAI,
    [switch]$DisableWidgets, [switch]$HideWidgets,
    [switch]$DisableChat, [switch]$HideChat,
    [switch]$EnableEndTask,
    [switch]$EnableLastActiveClick,
    [switch]$ClearStart,
    [string]$ReplaceStart,
    [switch]$ClearStartAllUsers,
    [string]$ReplaceStartAllUsers,
    [switch]$RevertContextMenu,
    [switch]$DisableMouseAcceleration,
    [switch]$DisableStickyKeys,
    [switch]$HideHome,
    [switch]$HideGallery,
    [switch]$ExplorerToHome,
    [switch]$ExplorerToThisPC,
    [switch]$ExplorerToDownloads,
    [switch]$ExplorerToOneDrive,
    [switch]$NoRestartExplorer,
    [switch]$DisableOnedrive, [switch]$HideOnedrive,
    [switch]$Disable3dObjects, [switch]$Hide3dObjects,
    [switch]$DisableMusic, [switch]$HideMusic,
    [switch]$DisableIncludeInLibrary, [switch]$HideIncludeInLibrary,
    [switch]$DisableGiveAccessTo, [switch]$HideGiveAccessTo,
    [switch]$DisableShare, [switch]$HideShare
)


# Show error if current powershell environment is limited by security policies
if ($ExecutionContext.SessionState.LanguageMode -ne "FullLanguage") {
    Write-Host "错误：Win11Debloat 无法在您的系统上运行，因为安全策略限制了 PowerShell 的执行。" -ForegroundColor Red
    AwaitKeyToExit
}

# Log script output to 'Win11Debloat.log' at the specified path
if ($LogPath -and (Test-Path $LogPath)) {
    Start-Transcript -Path "$LogPath/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}
else {
    Start-Transcript -Path "$PSScriptRoot/Win11Debloat.log" -Append -IncludeInvocationHeader -Force | Out-Null
}

# Shows application selection form that allows the user to select what apps they want to remove or keep
function ShowAppSelectionForm {
    [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
    [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

    # Initialise form objects
    $form = New-Object System.Windows.Forms.Form
    $label = New-Object System.Windows.Forms.Label
    $button1 = New-Object System.Windows.Forms.Button
    $button2 = New-Object System.Windows.Forms.Button
    $selectionBox = New-Object System.Windows.Forms.CheckedListBox 
    $loadingLabel = New-Object System.Windows.Forms.Label
    $onlyInstalledCheckBox = New-Object System.Windows.Forms.CheckBox
    $checkUncheckCheckBox = New-Object System.Windows.Forms.CheckBox
    $initialFormWindowState = New-Object System.Windows.Forms.FormWindowState

    $script:selectionBoxIndex = -1

    # saveButton eventHandler
    $handler_saveButton_Click= 
    {
        if ($selectionBox.CheckedItems -contains "Microsoft.WindowsStore" -and -not $Silent) {
            $warningSelection = [System.Windows.Forms.Messagebox]::Show('您确定要卸载“微软商店”吗？该应用程序无法轻松重新安装。', 'Are you sure?', 'YesNo', 'Warning')
        
            if ($warningSelection -eq 'No') {
                return
            }
        }

        $script:SelectedApps = $selectionBox.CheckedItems

        # Create file that stores selected apps if it doesn't exist
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            $null = New-Item "$PSScriptRoot/CustomAppsList"
        } 

        Set-Content -Path "$PSScriptRoot/CustomAppsList" -Value $script:SelectedApps

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    }

    # cancelButton eventHandler
    $handler_cancelButton_Click= 
    {
        $form.Close()
    }

    $selectionBox_SelectedIndexChanged= 
    {
        $script:selectionBoxIndex = $selectionBox.SelectedIndex
    }

    $selectionBox_MouseDown=
    {
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            if ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift) {
                if ($script:selectionBoxIndex -ne -1) {
                    $topIndex = $script:selectionBoxIndex

                    if ($selectionBox.SelectedIndex -gt $topIndex) {
                        for (($i = ($topIndex)); $i -le $selectionBox.SelectedIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                    elseif ($topIndex -gt $selectionBox.SelectedIndex) {
                        for (($i = ($selectionBox.SelectedIndex)); $i -le $topIndex; $i++) {
                            $selectionBox.SetItemChecked($i, $selectionBox.GetItemChecked($topIndex))
                        }
                    }
                }
            }
            elseif ($script:selectionBoxIndex -ne $selectionBox.SelectedIndex) {
                $selectionBox.SetItemChecked($selectionBox.SelectedIndex, -not $selectionBox.GetItemChecked($selectionBox.SelectedIndex))
            }
        }
    }

    $check_All=
    {
        for (($i = 0); $i -lt $selectionBox.Items.Count; $i++) {
            $selectionBox.SetItemChecked($i, $checkUncheckCheckBox.Checked)
        }
    }

    $load_Apps=
    {
        # Correct the initial state of the form to prevent the .Net maximized form issue
        $form.WindowState = $initialFormWindowState

        # Reset state to default before loading appslist again
        $script:selectionBoxIndex = -1
        $checkUncheckCheckBox.Checked = $False

        # Show loading indicator
        $loadingLabel.Visible = $true
        $form.Refresh()

        # Clear selectionBox before adding any new items
        $selectionBox.Items.Clear()

        # Set filePath where Appslist can be found
        $appsFile = "$PSScriptRoot/Appslist.txt"
        $listOfApps = ""

        if ($onlyInstalledCheckBox.Checked -and ($script:wingetInstalled -eq $true)) {
            # Attempt to get a list of installed apps via winget, times out after 10 seconds
            $job = Start-Job { return winget list --accept-source-agreements --disable-interactivity }
            $jobDone = $job | Wait-Job -TimeOut 10

            if (-not $jobDone) {
                # Show error that the script was unable to get list of apps from winget
                [System.Windows.MessageBox]::Show('由于无法通过 winget 加载已安装应用程序的列表，因此某些应用程序可能不会出现在列表中。', 'Error', 'Ok', 'Error')
            }
            else {
                # Add output of job (list of apps) to $listOfApps
                $listOfApps = Receive-Job -Job $job
            }
        }

        # Go through appslist and add items one by one to the selectionBox
        Foreach ($app in (Get-Content -Path $appsFile | Where-Object { $_ -notmatch '^\s*$' -and $_ -notmatch '^#  .*' -and $_ -notmatch '^# -* #' } )) { 
            $appChecked = $true

            # Remove first # if it exists and set appChecked to false
            if ($app.StartsWith('#')) {
                $app = $app.TrimStart("#")
                $appChecked = $false
            }

            # Remove any comments from the Appname
            if (-not ($app.IndexOf('#') -eq -1)) {
                $app = $app.Substring(0, $app.IndexOf('#'))
            }
            
            # Remove leading and trailing spaces and `*` characters from Appname
            $app = $app.Trim()
            $appString = $app.Trim('*')

            # Make sure appString is not empty
            if ($appString.length -gt 0) {
                if ($onlyInstalledCheckBox.Checked) {
                    # onlyInstalledCheckBox is checked, check if app is installed before adding it to selectionBox
                    if (-not ($listOfApps -like ("*$appString*")) -and -not (Get-AppxPackage -Name $app)) {
                        # App is not installed, continue with next item
                        continue
                    }
                    if (($appString -eq "Microsoft.Edge") -and -not ($listOfApps -like "* Microsoft.Edge *")) {
                        # App is not installed, continue with next item
                        continue
                    }
                }

                # Add the app to the selectionBox and set it's checked status
                $selectionBox.Items.Add($appString, $appChecked) | Out-Null
            }
        }
        
        # Hide loading indicator
        $loadingLabel.Visible = $False

        # Sort selectionBox alphabetically
        $selectionBox.Sorted = $True
    }

    $form.Text = "Win11Debloat Application Selection"
    $form.Name = "appSelectionForm"
    $form.DataBindings.DefaultDataSourceUpdateMode = 0
    $form.ClientSize = New-Object System.Drawing.Size(400,502)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $False

    $button1.TabIndex = 4
    $button1.Name = "saveButton"
    $button1.UseVisualStyleBackColor = $True
    $button1.Text = "Confirm"
    $button1.Location = New-Object System.Drawing.Point(27,472)
    $button1.Size = New-Object System.Drawing.Size(75,23)
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0
    $button1.add_Click($handler_saveButton_Click)

    $form.Controls.Add($button1)

    $button2.TabIndex = 5
    $button2.Name = "cancelButton"
    $button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $button2.UseVisualStyleBackColor = $True
    $button2.Text = "Cancel"
    $button2.Location = New-Object System.Drawing.Point(129,472)
    $button2.Size = New-Object System.Drawing.Size(75,23)
    $button2.DataBindings.DefaultDataSourceUpdateMode = 0
    $button2.add_Click($handler_cancelButton_Click)

    $form.Controls.Add($button2)

    $label.Location = New-Object System.Drawing.Point(13,5)
    $label.Size = New-Object System.Drawing.Size(400,14)
    $Label.Font = 'Microsoft Sans Serif,8'
    $label.Text = '勾选您想要删除的应用程序，取消勾选您想要保留的应用程序。'

    $form.Controls.Add($label)

    $loadingLabel.Location = New-Object System.Drawing.Point(16,46)
    $loadingLabel.Size = New-Object System.Drawing.Size(300,418)
    $loadingLabel.Text = '正在加载APP列表...'
    $loadingLabel.BackColor = "White"
    $loadingLabel.Visible = $false

    $form.Controls.Add($loadingLabel)

    $onlyInstalledCheckBox.TabIndex = 6
    $onlyInstalledCheckBox.Location = New-Object System.Drawing.Point(230,474)
    $onlyInstalledCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $onlyInstalledCheckBox.Text = '仅显示已安装的应用程序'
    $onlyInstalledCheckBox.add_CheckedChanged($load_Apps)

    $form.Controls.Add($onlyInstalledCheckBox)

    $checkUncheckCheckBox.TabIndex = 7
    $checkUncheckCheckBox.Location = New-Object System.Drawing.Point(16,22)
    $checkUncheckCheckBox.Size = New-Object System.Drawing.Size(150,20)
    $checkUncheckCheckBox.Text = '全选/取消全选'
    $checkUncheckCheckBox.add_CheckedChanged($check_All)

    $form.Controls.Add($checkUncheckCheckBox)

    $selectionBox.FormattingEnabled = $True
    $selectionBox.DataBindings.DefaultDataSourceUpdateMode = 0
    $selectionBox.Name = "selectionBox"
    $selectionBox.Location = New-Object System.Drawing.Point(13,43)
    $selectionBox.Size = New-Object System.Drawing.Size(374,424)
    $selectionBox.TabIndex = 3
    $selectionBox.add_SelectedIndexChanged($selectionBox_SelectedIndexChanged)
    $selectionBox.add_Click($selectionBox_MouseDown)

    $form.Controls.Add($selectionBox)

    # Save the initial state of the form
    $initialFormWindowState = $form.WindowState

    # Load apps into selectionBox
    $form.add_Load($load_Apps)

    # Focus selectionBox when form opens
    $form.Add_Shown({$form.Activate(); $selectionBox.Focus()})

    # Show the Form
    return $form.ShowDialog()
}


# Returns list of apps from the specified file, it trims the app names and removes any comments
function ReadAppslistFromFile {
    param (
        $appsFilePath
    )

    $appsList = @()

    # Get list of apps from file at the path provided, and remove them one by one
    Foreach ($app in (Get-Content -Path $appsFilePath | Where-Object { $_ -notmatch '^#.*' -and $_ -notmatch '^\s*$' } )) { 
        # Remove any comments from the Appname
        if (-not ($app.IndexOf('#') -eq -1)) {
            $app = $app.Substring(0, $app.IndexOf('#'))
        }

        # Remove any spaces before and after the Appname
        $app = $app.Trim()
        
        $appString = $app.Trim('*')
        $appsList += $appString
    }

    return $appsList
}


# Removes apps specified during function call from all user accounts and from the OS image.
function RemoveApps {
    param (
        $appslist
    )

    Foreach ($app in $appsList) { 
        Write-Output "试图移除 $app..."

        if (($app -eq "Microsoft.OneDrive") -or ($app -eq "Microsoft.Edge")) {
            # Use winget to remove OneDrive and Edge
            if ($script:wingetInstalled -eq $false) {
                Write-Host "错误：WinGet 未安装或已过时, $app 无法移除" -ForegroundColor Red
            }
            else {
                # Uninstall app via winget
                Strip-Progress -ScriptBlock { winget uninstall --accept-source-agreements --disable-interactivity --id $app } | Tee-Object -Variable wingetOutput 

                If (($app -eq "Microsoft.Edge") -and (Select-String -InputObject $wingetOutput -Pattern "卸载操作失败，返回代码为")) {
                    Write-Host "无法通过 Winget 卸载 Microsoft Edge" -ForegroundColor Red
                    Write-Output ""

                    if ($( Read-Host -Prompt "您是否想要强行卸载 Edge 浏览器？不建议这样操作！ (y/n)" ) -eq 'y') {
                        Write-Output ""
                        ForceRemoveEdge
                    }
                }
            }
        }
        else {
            # Use Remove-AppxPackage to remove all other apps
            $app = '*' + $app + '*'

            # Remove installed app for all existing users
            if ($WinVersion -ge 22000) {
                # Windows 11 build 22000 or later
                try {
                    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction Continue

                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "为所有用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
            }
            else {
                # Windows 10
                try {
                    Get-AppxPackage -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
                    
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "为当前用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为当前用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
                
                try {
                    Get-AppxPackage -Name $app -PackageTypeFilter Main, Bundle, Resource -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                    
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "为所有用户移除 $app" -ForegroundColor DarkGray
                    }
                }
                catch {
                    if ($DebugPreference -ne "SilentlyContinue") {
                        Write-Host "无法为所有用户移除 $app" -ForegroundColor Yellow
                        Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
                    }
                }
            }

            # Remove provisioned app from OS image, so the app won't be installed for any new users
            try {
                Get-AppxProvisionedPackage -Online | Where-Object { $_.PackageName -like $app } | ForEach-Object { Remove-ProvisionedAppxPackage -Online -AllUsers -PackageName $_.PackageName }
            }
            catch {
                Write-Host "无法从Windows镜像中移除 $app" -ForegroundColor Yellow
                Write-Host $psitem.Exception.StackTrace -ForegroundColor Gray
            }
        }
    }
            
    Write-Output ""
}


# Forcefully removes Microsoft Edge using it's uninstaller
function ForceRemoveEdge {
    # Based on work from loadstring1 & ave9858
    Write-Output "> 强制卸载微软 Edge..."

    $regView = [Microsoft.Win32.RegistryView]::Registry32
    $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
    $hklm.CreateSubKey('SOFTWARE\Microsoft\EdgeUpdateDev').SetValue('AllowUninstall', '')

    # Create stub (Creating this somehow allows uninstalling Edge)
    $edgeStub = "$env:SystemRoot\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe"
    New-Item $edgeStub -ItemType Directory | Out-Null
    New-Item "$edgeStub\MicrosoftEdge.exe" | Out-Null

    # Remove edge
    $uninstallRegKey = $hklm.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Microsoft Edge')
    if ($null -ne $uninstallRegKey) {
        Write-Output "正在运行卸载程序..."
        $uninstallString = $uninstallRegKey.GetValue('UninstallString') + ' --force-uninstall'
        Start-Process cmd.exe "/c $uninstallString" -WindowStyle Hidden -Wait

        Write-Output "正在删除剩余文件..."

        $edgePaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Edge.lnk",
            "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Tombstones\Microsoft Edge.lnk",
            "$env:PUBLIC\Desktop\Microsoft Edge.lnk",
            "$env:USERPROFILE\Desktop\Microsoft Edge.lnk",
            "$edgeStub"
        )

        foreach ($path in $edgePaths) {
            if (Test-Path -Path $path) {
                Remove-Item -Path $path -Force -Recurse -ErrorAction SilentlyContinue
                Write-Host "  Removed $path" -ForegroundColor DarkGray
            }
        }

        Write-Output "正在清理注册表..."

        # Remove MS Edge from autostart
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run" /v "Microsoft Edge Update" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "MicrosoftEdgeAutoLaunch_A9F6DCE4ABADF4F51CF45CD7129E3C6C" /f *>$null
        reg delete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" /v "Microsoft Edge Update" /f *>$null

        Write-Output "微软 Edge 已被卸载。"
    }
    else {
        Write-Output ""
        Write-Host "错误：无法强制卸载 Microsoft Edge，未能找到卸载程序" -ForegroundColor Red
    }
    
    Write-Output ""
}


# Execute provided command and strips progress spinners/bars from console output
function Strip-Progress {
    param(
        [ScriptBlock]$ScriptBlock
    )

    # Regex pattern to match spinner characters and progress bar patterns
    $progressPattern = 'Γ?[?ê]|^\s+[-\\|/]\s+$'

    # Corrected regex pattern for size formatting, ensuring proper capture groups are utilized
    $sizePattern = '(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB) /\s+(\d+(\.\d{1,2})?)\s+(B|KB|MB|GB|TB|PB)'

    & $ScriptBlock 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            "ERROR: $($_.Exception.Message)"
        } else {
            $line = $_ -replace $progressPattern, '' -replace $sizePattern, ''
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and -not ($line.StartsWith('  '))) {
                $line
            }
        }
    }
}


# Check if this machine supports S0 Modern Standby power state. Returns true if S0 Modern Standby is supported, false otherwise.
function CheckModernStandbySupport {
    $count = 0

    try {
        switch -Regex (powercfg /a) {
            ':' {
                $count += 1
            }

            '(.*S0.{1,}\))' {
                if ($count -eq 1) {
                    return $true
                }
            }
        }
    }
    catch {
        Write-Host "错误：无法检查 S0 现代待机支持情况，执行 powercfg 命令失败。" -ForegroundColor Red
        Write-Host ""
        Write-Host "请按任意键继续..."
        $null = [System.Console]::ReadKey()
        return $true
    }

    return $false
}


# Returns the directory path of the specified user, exits script if user path can't be found
function GetUserDirectory {
    param (
        $userName,
        $fileName = "",
        $exitIfPathNotFound = $true
    )

    try {
        $userDirectoryExists = Test-Path "$env:SystemDrive\Users\$userName"
        $userPath = "$env:SystemDrive\Users\$userName\$fileName"
    
        if ((Test-Path $userPath) -or ($userDirectoryExists -and (-not $exitIfPathNotFound))) {
            return $userPath
        }
    
        $userDirectoryExists = Test-Path ($env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$userName")
        $userPath = $env:USERPROFILE -Replace ('\\' + $env:USERNAME + '$'), "\$userName\$fileName"
    
        if ((Test-Path $userPath) -or ($userDirectoryExists -and (-not $exitIfPathNotFound))) {
            return $userPath
        }
    } catch {
        Write-Host "错误：在尝试查找用户 $userName 的用户目录路径时出现错误。请确认该用户已在本系统中存在。" -ForegroundColor Red
        AwaitKeyToExit
    }

    Write-Host "错误：无法找到 $userName 用户的目录路径。 " -ForegroundColor Red
    AwaitKeyToExit
}


# Import & execute regfile
function RegImport {
    param (
        $message,
        $path
    )

    Write-Output $message

    if ($script:Params.ContainsKey("Sysprep")) {
        $defaultUserPath = GetUserDirectory -userName "Default" -fileName "NTUSER.DAT"
        
        reg load "HKU\Default" $defaultUserPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
    }
    elseif ($script:Params.ContainsKey("User")) {
        $userPath = GetUserDirectory -userName $script:Params.Item("User") -fileName "NTUSER.DAT"
        
        reg load "HKU\Default" $userPath | Out-Null
        reg import "$PSScriptRoot\Regfiles\Sysprep\$path"
        reg unload "HKU\Default" | Out-Null
        
    }
    else {
        reg import "$PSScriptRoot\Regfiles\$path"  
    }

    Write-Output ""
}


# Restart the Windows Explorer process
function RestartExplorer {
    if ($script:Params.ContainsKey("Sysprep") -or $script:Params.ContainsKey("User") -or $script:Params.ContainsKey("NoRestartExplorer")) {
        return
    }

    Write-Output "> 正在重新启动 Windows 资源管理器进程以应用所有更改……（这可能会导致一些闪烁现象）"

    if ($script:Params.ContainsKey("DisableMouseAcceleration")) {
        Write-Host "警告：提高指针精度 设置的更改仅在重新启动后才会生效。" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableStickyKeys")) {
        Write-Host "警告：粘贴键 设置的更改需在重新启动后才会生效。" -ForegroundColor Yellow
    }

    if ($script:Params.ContainsKey("DisableAnimations")) {
        Write-Host "警告：动画功能 设置的更改需在重新启动后才会生效。" -ForegroundColor Yellow
    }

    # Only restart if the powershell process matches the OS architecture.
    # Restarting explorer from a 32bit PowerShell window will fail on a 64bit OS
    if ([Environment]::Is64BitProcess -eq [Environment]::Is64BitOperatingSystem) {
        Stop-Process -processName: Explorer -Force
    }
    else {
        Write-Warning "由于一些问题，我们无法重新启动 Windows 资源管理器进程，请手动重启您的电脑以应用所有更改。"
    }
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenuForAllUsers {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin"
    )

    Write-Output "> 正在为所有用户从开始菜单中移除所有已锁定的应用程序……"

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法清理开始菜单，脚本文件夹中缺少 start2.bin 文件" -ForegroundColor Red
        Write-Output ""
        return
    }

    # Get path to start menu file for all users
    $userPathString = GetUserDirectory -userName "*" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState"
    $usersStartMenuPaths = get-childitem -path $userPathString

    # Go through all users and replace the start menu file
    ForEach ($startMenuPath in $usersStartMenuPaths) {
        ReplaceStartMenu $startMenuTemplate "$($startMenuPath.Fullname)\start2.bin"
    }

    # Also replace the start menu file for the default user profile
    $defaultStartMenuPath = GetUserDirectory -userName "Default" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState" -exitIfPathNotFound $false

    # Create folder if it doesn't exist
    if (-not (Test-Path $defaultStartMenuPath)) {
        new-item $defaultStartMenuPath -ItemType Directory -Force | Out-Null
        Write-Output "为默认用户配置文件创建了 LocalState 文件夹"
    }

    # Copy template to default profile
    Copy-Item -Path $startMenuTemplate -Destination $defaultStartMenuPath -Force
    Write-Output "为默认用户账户替换了开始菜单"
    Write-Output ""
}


# Replace the startmenu for all users, when using the default startmenuTemplate this clears all pinned apps
# Credit: https://lazyadmin.nl/win-11/customize-windows-11-start-menu-layout/
function ReplaceStartMenu {
    param (
        $startMenuTemplate = "$PSScriptRoot/Assets/Start/start2.bin",
        $startMenuBinFile = "$env:LOCALAPPDATA\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin"
    )

    # Change path to correct user if a user was specified
    if ($script:Params.ContainsKey("User")) {
        $startMenuBinFile = GetUserDirectory -userName "$(GetUserName)" -fileName "AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\start2.bin" -exitIfPathNotFound $false
    }

    # Check if template bin file exists, return early if it doesn't
    if (-not (Test-Path $startMenuTemplate)) {
        Write-Host "错误：无法替换开始菜单，模板文件未找到" -ForegroundColor Red
        return
    }

    if ([IO.Path]::GetExtension($startMenuTemplate) -ne ".bin" ) {
        Write-Host "错误：无法替换开始菜单，模板文件不是有效的.bin 文件" -ForegroundColor Red
        return
    }

    $userName = [regex]::Match($startMenuBinFile, '(?:Users\\)([^\\]+)(?:\\AppData)').Groups[1].Value

    $backupBinFile = $startMenuBinFile + ".bak"

    if (Test-Path $startMenuBinFile) {
        # Backup current start menu file
        Move-Item -Path $startMenuBinFile -Destination $backupBinFile -Force
    } else {
        Write-Host "警告：无法找到用户 $userName  所需的原始 start2.bin 文件。 因此该用户的备份未被创建！" -ForegroundColor Yellow
        New-Item -ItemType File -Path $startMenuBinFile -Force
    }

    # Copy template file
    Copy-Item -Path $startMenuTemplate -Destination $startMenuBinFile -Force

    Write-Output "为用户 $userName 替换了开始菜单"
}


# Add parameter to script and write to file
function AddParameter {
    param (
        $parameterName,
        $message,
        $addToFile = $true
    )

    # Add key if it doesn't already exist
    if (-not $script:Params.ContainsKey($parameterName)) {
        $script:Params.Add($parameterName, $true)
    }

    if (-not $addToFile) {
        Write-Output "- $message"
        return
    }

    # Create or clear file that stores last used settings
    if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
        $null = New-Item "$PSScriptRoot/SavedSettings"
    } 
    elseif ($script:FirstSelection) {
        $null = Clear-Content "$PSScriptRoot/SavedSettings"
    }
    
    $script:FirstSelection = $false

    # Create entry and add it to the file
    $entry = "$parameterName#- $message"
    Add-Content -Path "$PSScriptRoot/SavedSettings" -Value $entry
}


function PrintHeader {
    param (
        $title
    )

    $fullTitle = " Win11Debloat-zh_CN Script - $title"

    if ($script:Params.ContainsKey("Sysprep")) {
        $fullTitle = "$fullTitle (Sysprep mode)"
    }
    else {
        $fullTitle = "$fullTitle (User: $(GetUserName))"
    }

    Clear-Host
    Write-Output "-------------------------------------------------------------------------------------------"
    Write-Output $fullTitle
    Write-Output "-------------------------------------------------------------------------------------------"
}


function PrintFromFile {
    param (
        $path,
        $title,
        $printHeader = $true
    )

    if ($printHeader) {
        Clear-Host

        PrintHeader $title
    }

    # Get & print script menu from file
    Foreach ($line in (Get-Content -Path $path )) {   
        Write-Output $line
    }
}


function PrintAppsList {
    param (
        $path,
        $printCount = $false
    )

    if (-not (Test-Path $path)) {
        return
    }
    
    $appsList = ReadAppslistFromFile $path

    if ($printCount) {
        Write-Output "- Remove $($appsList.Count) apps:"
    }

    Write-Host $appsList -ForegroundColor DarkGray
}


function AwaitKeyToExit {
    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output "按任意按键退出..."
        $null = [System.Console]::ReadKey()
    }

    Stop-Transcript
    Exit
}


function GetUserName {
    if ($script:Params.ContainsKey("User")) { 
        return $script:Params.Item("User") 
    }
    
    return $env:USERNAME
}


function CreateSystemRestorePoint {
    Write-Output "> 正在尝试创建系统还原点......"
    
    $SysRestore = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "RPSessionInterval"

    if ($SysRestore.RPSessionInterval -eq 0) {
        if ($Silent -or $( Read-Host -Prompt "系统还原已禁用，您是否要启用它并创建还原点？ (y/n)") -eq 'y') {
            $enableSystemRestoreJob = Start-Job { 
                try {
                    Enable-ComputerRestore -Drive "$env:SystemDrive"
                } catch {
                    Write-Host "错误：系统还原功能无法启用: $_" -ForegroundColor Red
                    return
                }
            }
    
            $enableSystemRestoreJobDone = $enableSystemRestoreJob | Wait-Job -TimeOut 20

            if (-not $enableSystemRestoreJobDone) {
                Write-Host "错误：无法启用系统还原功能并创建还原点，操作超时。" -ForegroundColor Red
                return
            } else {
                Receive-Job $enableSystemRestoreJob
            }
        } else {
            Write-Output ""
            return
        }
    }

    $createRestorePointJob = Start-Job { 
        # Find existing restore points that are less than 24 hours old
        try {
            $recentRestorePoints = Get-ComputerRestorePoint | Where-Object { (Get-Date) - [System.Management.ManagementDateTimeConverter]::ToDateTime($_.CreationTime) -le (New-TimeSpan -Hours 24) }
        } catch {
            Write-Host "错误：无法获取已有的还原点： $_" -ForegroundColor Red
            return
        }
    
        if ($recentRestorePoints.Count -eq 0) {
            try {
                Checkpoint-Computer -Description "Restore point created by Win11Debloat" -RestorePointType "MODIFY_SETTINGS"
                Write-Output "系统还原点创建成功!"
            } catch {
                Write-Host "错误：无法创建还原点: $_" -ForegroundColor Red
            }
        } else {
            Write-Host "当前已存在一个还原点，因此未创建新的还原点。" -ForegroundColor Yellow
        }
    }
    
    $createRestorePointJobDone = $createRestorePointJob | Wait-Job -TimeOut 20

    if (-not $createRestorePointJobDone) {
        Write-Host "错误：创建系统还原点失败，操作超时。" -ForegroundColor Red
    } else {
        Receive-Job $createRestorePointJob
    }

    Write-Output ""
}


function DisplayCustomModeOptions {
    # Get current Windows build version to compare against features
    $WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild
            
    PrintHeader '自定义模式'

    AddParameter 'CreateRestorePoint' '创建一个系统还原点'

    # Show options for removing apps, only continue on valid input
    Do {
        Write-Host "选项:" -ForegroundColor Yellow
        Write-Host " (n) 不要移除任何应用程序" -ForegroundColor Yellow
        Write-Host " (1) 仅移除脚本默认选择的程序" -ForegroundColor Yellow
        Write-Host " (2) 移除脚本默认选择的程序, 并且包括邮件和日历应用程序以及与游戏相关的应用程序。"  -ForegroundColor Yellow
        Write-Host " (3) 手动选择需要移除的应用程序" -ForegroundColor Yellow
        $RemoveAppsInput = Read-Host "您是否需要删除某些应用程序？这些应用程序将对所有用户进行删除操作 (n/1/2/3)"

        # Show app selection form if user entered option 3
        if ($RemoveAppsInput -eq '3') {
            $result = ShowAppSelectionForm

            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                Write-Output ""
                Write-Host "应用选择已取消, 请重试" -ForegroundColor Red

                $RemoveAppsInput = 'c'
            }
            
            Write-Output ""
        }
    }
    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2' -and $RemoveAppsInput -ne '3') 

    # Select correct option based on user input
    switch ($RemoveAppsInput) {
        '1' {
            AddParameter 'RemoveApps' '移除所有脚本默认选择的应用程序'
        }
        '2' {
            AddParameter 'RemoveApps' '移除所有脚本默认选择的应用程序'
            AddParameter 'RemoveCommApps' '移除“邮件”、“日历”和“通讯录”应用程序'
            AddParameter 'RemoveW11Outlook' '移除新的 Windows 版 Outlook 应用程序'
            AddParameter 'RemoveGamingApps' '移除 Xbox 应用程序和 Xbox 游戏栏'
            AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制功能'
        }
        '3' {
            Write-Output "您已选定 $($script:SelectedApps.Count) 个应用程序进行删除操作。"

            AddParameter 'RemoveAppsCustom' "移除 $($script:SelectedApps.Count) 应用:"

            Write-Output ""

            if ($( Read-Host -Prompt "是否禁用 Xbox 游戏/屏幕录制功能？这还会阻止游戏覆盖界面的弹出窗口。 (y/n)" ) -eq 'y') {
                AddParameter 'DisableDVR' '禁用 Xbox 游戏/屏幕录制功能'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用遥测功能、诊断数据、活动记录、应用程序启动追踪以及定向广告？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableTelemetry' '禁用遥测功能、诊断数据、活动记录、应用程序启动追踪以及定向广告。'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否在开始、设置、通知、资源管理器、锁屏以及Edge中禁用提示、技巧、建议以及广告？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableSuggestions' '在开始、设置、通知以及文件资源管理器中禁用提示、技巧、建议以及广告。'
        AddParameter 'DisableEdgeAds' '在微软 Edge 浏览器中禁用广告、建议以及 MSN 新闻推送功能。'
        AddParameter 'DisableSettings365Ads' '在“设置”主界面中禁用 Microsoft 365 广告'
        AddParameter 'DisableLockscreenTips' '在锁屏界面禁用提示与技巧功能'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用并移除 Windows 搜索中的必应网页搜索、必应人工智能以及小娜？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableBing' '禁用并移除 Windows 搜索中的必应网页搜索、必应人工智能以及小娜功能。'
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        # Show options for disabling/removing AI features, only continue on valid input
        Do {
            Write-Host "选项:" -ForegroundColor Yellow
            Write-Host " (n) 不要禁用任何人工智能功能" -ForegroundColor Yellow
            Write-Host " (1) 禁用微软智能助手、Windows Recall功能和Click to Do功能" -ForegroundColor Yellow
            Write-Host " (2) 禁用 Microsoft Edge、画图和记事本中的 Microsoft Copilot、Windows Recall功能、Click to Do功能以及人工智能相关特性。"  -ForegroundColor Yellow
            $DisableAIInput = Read-Host "您是否想要禁用任何人工智能功能？这适用于所有用户 (n/1/2)"
        }
        while ($DisableAIInput -ne 'n' -and $DisableAIInput -ne '0' -and $DisableAIInput -ne '1' -and $DisableAIInput -ne '2') 

        # Select correct option based on user input
        switch ($DisableAIInput) {
            '1' {
                AddParameter 'DisableCopilot' '禁用并移除微软“Copilot”功能'
                AddParameter 'DisableRecall' '禁用 Windows Recall 功能'
                AddParameter 'DisableClickToDo' '禁用 Click to Do 功能(AI 文本与图像分析)'
            }
            '2' {
                AddParameter 'DisableCopilot' '禁用并移除微软“Copilot”功能'
                AddParameter 'DisableRecall' '禁用 Windows Recall 功能'
                AddParameter 'DisableClickToDo' '禁用 Click to Do 功能(AI 文本与图像分析)'
                AddParameter 'DisableEdgeAI' '禁用 Edge 浏览器中的 AI 功能'
                AddParameter 'DisablePaintAI' '禁用画图程序中的 AI 功能'
                AddParameter 'DisableNotepadAI' '禁用记事本中的 AI 功能'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否关闭桌面上的 Windows 聚焦？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableDesktopSpotlight' '禁用 Windows 聚焦 桌面背景选项'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否为系统和应用程序启用暗黑模式？ (y/n)" ) -eq 'y') {
        AddParameter 'EnableDarkMode' '启用系统及应用程序的暗黑模式'
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用透明度、动画和视觉效果？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableTransparency' '禁用透明效果'
        AddParameter 'DisableAnimations' '禁用动画和视觉效果'
    }

    # Only show this option for Windows 11 users running build 22000 or later
    if ($WinVersion -ge 22000) {
        Write-Output ""

        if ($( Read-Host -Prompt "是否恢复旧版 Windows 10 的上下文菜单吗？ (y/n)" ) -eq 'y') {
            AddParameter 'RevertContextMenu' '恢复旧版 Windows 10 的上下文菜单'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否关闭 增强指针精度 功能，也就是所谓的 鼠标加速 功能？ (y/n)" ) -eq 'y') {
        AddParameter 'DisableMouseAcceleration' '关闭 增强指针精度 功能 (鼠标加速)'
    }

    # Only show this option for Windows 11 users running build 26100 or later
    if ($WinVersion -ge 26100) {
        Write-Output ""

        if ($( Read-Host -Prompt "是否禁用 锁定键 快捷键？ (y/n)" ) -eq 'y') {
            AddParameter 'DisableStickyKeys' '禁用 锁定键 快捷键'
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "是否禁用快速启动？这适用于所有用户 (y/n)" ) -eq 'y') {
        AddParameter 'DisableFastStartup' '禁用快速启动'
    }

    # Only show this option for Windows 11 users running build 22000 or later, and if the machine has at least one battery
    if (($WinVersion -ge 22000) -and $script:ModernStandbySupported) {
        Write-Output ""

        if ($( Read-Host -Prompt "是否在“现代待机”模式下禁用网络连接？这适用于所有用户。 (y/n)" ) -eq 'y') {
            AddParameter 'DisableModernStandbyNetworking' '在“现代待机”模式下禁用网络连接'
        }
    }

    # Only show option for disabling context menu items for Windows 10 users or if the user opted to restore the Windows 10 context menu
    if ((get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") -or $script:Params.ContainsKey('RevertContextMenu')) {
        Write-Output ""

        if ($( Read-Host -Prompt "是否禁用任何上右键菜单选项？ (y/n)" ) -eq 'y') {
            Write-Output ""

            if ($( Read-Host -Prompt "   是否要隐藏 包含在库中 这一选项在右键菜单中？ (y/n)" ) -eq 'y') {
                AddParameter 'HideIncludeInLibrary' "隐藏 包含在库中 这一选项在右键菜单中"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否要隐藏 授予访问权限 这一选项在右键菜单中？ (y/n)" ) -eq 'y') {
                AddParameter 'HideGiveAccessTo' "隐藏 授予访问权限 这一选项在右键菜单中"
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否要隐藏 分享 这一选项在右键菜单中？ (y/n)" ) -eq 'y') {
                AddParameter 'HideShare' "隐藏 分享 这一选项在右键菜单中"
            }
        }
    }

    # Only show this option for Windows 11 users running build 22621 or later
    if ($WinVersion -ge 22621) {
        Write-Output ""

        if ($( Read-Host -Prompt "您是否想要对 开始 菜单进行任何更改？ (y/n)" ) -eq 'y') {
            Write-Output ""

            if ($script:Params.ContainsKey("Sysprep")) {
                if ($( Read-Host -Prompt "是否要移除所有用户（包括现有用户和新用户）的 开始 菜单中的所有固定应用？ (y/n)" ) -eq 'y') {
                    AddParameter 'ClearStartAllUsers' '为现有用户和新用户移除 开始 菜单中的所有固定应用程序。'
                }
            }
            else {
                Do {
                    Write-Host "   选项:" -ForegroundColor Yellow
                    Write-Host "    (n) 不要从开始菜单中移除任何已设置为固定的应用程序。" -ForegroundColor Yellow
                    Write-Host "    (1) 仅移除 ($(GetUserName)) 开始菜单中任何已设置为固定的应用程序。" -ForegroundColor Yellow
                    Write-Host "    (2) 移除所有用户（包括现有用户和新用户）的 开始 菜单中的所有固定应用"  -ForegroundColor Yellow
                    $ClearStartInput = Read-Host "   将所有已固定在开始菜单中的应用程序移除？ (n/1/2)" 
                }
                while ($ClearStartInput -ne 'n' -and $ClearStartInput -ne '0' -and $ClearStartInput -ne '1' -and $ClearStartInput -ne '2') 

                # Select correct option based on user input
                switch ($ClearStartInput) {
                    '1' {
                        AddParameter 'ClearStart' "仅移除当前用户在开始菜单中任何已设置为固定的应用程序。"
                    }
                    '2' {
                        AddParameter 'ClearStartAllUsers' "移除所有用户（包括现有用户和新用户）的 开始 菜单中的所有固定应用"
                    }
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否要禁用 开始 菜单中的推荐功能？此操作适用于所有用户。 (y/n)" ) -eq 'y') {
                AddParameter 'DisableStartRecommended' '禁用 开始 菜单中的推荐功能，此操作适用于所有用户。'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   是否在 开始 菜单中禁用 手机连接 移动设备集成功能？ (y/n)" ) -eq 'y') {
                AddParameter 'DisableStartPhoneLink' '在 开始 菜单中禁用 手机连接 移动设备集成功能'
            }
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "您是否想要对任务栏及相关服务进行任何更改？ (y/n)" ) -eq 'y') {
        # Only show these specific options for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   将任务栏按钮调整至左侧？ (y/n)" ) -eq 'y') {
                AddParameter 'TaskbarAlignLeft' '将任务栏按钮调整至左侧'
            }

            # Show options for combine icon on taskbar, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 不更改" -ForegroundColor Yellow
                Write-Host "    (1) 总是" -ForegroundColor Yellow
                Write-Host "    (2) 当任务栏已满时" -ForegroundColor Yellow
                Write-Host "    (3) 从不" -ForegroundColor Yellow
                $TbCombineTaskbar = Read-Host "   将任务栏按钮合并并隐藏标签? (n/1/2/3)" 
            }
            while ($TbCombineTaskbar -ne 'n' -and $TbCombineTaskbar -ne '0' -and $TbCombineTaskbar -ne '1' -and $TbCombineTaskbar -ne '2' -and $TbCombineTaskbar -ne '3') 

            # Select correct taskbar goup option based on user input
            switch ($TbCombineTaskbar) {
                '1' {
                    AddParameter 'CombineTaskbarAlways' '始终将任务栏按钮组合在一起，并隐藏主显示屏上的标签。'
                    AddParameter 'CombineMMTaskbarAlways' '始终将任务栏按钮进行组合，并隐藏次要显示器上的标签。'
                }
                '2' {
                    AddParameter 'CombineTaskbarWhenFull' '在主显示屏上，当任务栏已满时，将任务栏按钮合并并隐藏标签。'
                    AddParameter 'CombineMMTaskbarWhenFull' '在副显示屏上，当任务栏已满时，将任务栏按钮合并并隐藏标签。'
                }
                '3' {
                    AddParameter 'CombineTaskbarNever' '从不将任务栏按钮与主显示屏上的显示标签进行组合设置'
                    AddParameter 'CombineMMTaskbarNever' '从不将任务栏按钮与副显示屏上的显示标签进行组合设置'
                }
            }

            # Show options for changing on what taskbar(s) app icons are shown, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   选项:" -ForegroundColor Yellow
                Write-Host "    (n) 不更改" -ForegroundColor Yellow
                Write-Host "    (1) 在所有任务栏上显示应用程序图标" -ForegroundColor Yellow
                Write-Host "    (2) 在主任务栏以及窗口正在打开时所在的任务栏上显示应用程序图标" -ForegroundColor Yellow
                Write-Host "    (3) 仅在窗口处于打开状态时，在任务栏上显示应用程序图标" -ForegroundColor Yellow
                $TbCombineTaskbar = Read-Host "   更改在使用多显示器时在任务栏上显示应用程序图标的方式？ (n/1/2/3)" 
            }
            while ($TbCombineTaskbar -ne 'n' -and $TbCombineTaskbar -ne '0' -and $TbCombineTaskbar -ne '1' -and $TbCombineTaskbar -ne '2' -and $TbCombineTaskbar -ne '3') 

            # Select correct taskbar goup option based on user input
            switch ($TbCombineTaskbar) {
                '1' {
                    AddParameter 'MMTaskbarModeAll' 'Show app icons on all taskbars'
                }
                '2' {
                    AddParameter 'MMTaskbarModeMainActive' 'Show app icons on main taskbar and on taskbar where the windows is open'
                }
                '3' {
                    AddParameter 'MMTaskbarModeActive' 'Show app icons only on taskbar where the window is open'
                }
            }

            # Show options for search icon on taskbar, only continue on valid input
            Do {
                Write-Output ""
                Write-Host "   Options:" -ForegroundColor Yellow
                Write-Host "    (n) No change" -ForegroundColor Yellow
                Write-Host "    (1) Hide search icon from the taskbar" -ForegroundColor Yellow
                Write-Host "    (2) Show search icon on the taskbar" -ForegroundColor Yellow
                Write-Host "    (3) Show search icon with label on the taskbar" -ForegroundColor Yellow
                Write-Host "    (4) Show search box on the taskbar" -ForegroundColor Yellow
                $TbSearchInput = Read-Host "   Hide or change the search icon on the taskbar? (n/1/2/3/4)" 
            }
            while ($TbSearchInput -ne 'n' -and $TbSearchInput -ne '0' -and $TbSearchInput -ne '1' -and $TbSearchInput -ne '2' -and $TbSearchInput -ne '3' -and $TbSearchInput -ne '4') 

            # Select correct taskbar search option based on user input
            switch ($TbSearchInput) {
                '1' {
                    AddParameter 'HideSearchTb' 'Hide search icon from the taskbar'
                }
                '2' {
                    AddParameter 'ShowSearchIconTb' 'Show search icon on the taskbar'
                }
                '3' {
                    AddParameter 'ShowSearchLabelTb' 'Show search icon with label on the taskbar'
                }
                '4' {
                    AddParameter 'ShowSearchBoxTb' 'Show search box on the taskbar'
                }
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   Hide the taskview button from the taskbar? (y/n)" ) -eq 'y') {
                AddParameter 'HideTaskview' 'Hide the taskview button from the taskbar'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   Disable the widgets service to remove widgets on the taskbar & lockscreen? (y/n)" ) -eq 'y') {
            AddParameter 'DisableWidgets' 'Disable widgets on the taskbar & lockscreen'
        }

        # Only show this options for Windows users running build 22621 or earlier
        if ($WinVersion -le 22621) {
            Write-Output ""

            if ($( Read-Host -Prompt "   Hide the chat (meet now) icon from the taskbar? (y/n)" ) -eq 'y') {
                AddParameter 'HideChat' 'Hide the chat (meet now) icon from the taskbar'
            }
        }
        
        # Only show this options for Windows users running build 22631 or later
        if ($WinVersion -ge 22631) {
            Write-Output ""

            if ($( Read-Host -Prompt "   Enable the 'End Task' option in the taskbar right click menu? (y/n)" ) -eq 'y') {
                AddParameter 'EnableEndTask' "Enable the 'End Task' option in the taskbar right click menu"
            }
        }
        
        Write-Output ""
        if ($( Read-Host -Prompt "   Enable the 'Last Active Click' behavior in the taskbar app area? (y/n)" ) -eq 'y') {
            AddParameter 'EnableLastActiveClick' "Enable the 'Last Active Click' behavior in the taskbar app area"
        }
    }

    Write-Output ""

    if ($( Read-Host -Prompt "Do you want to make any changes to File Explorer? (y/n)" ) -eq 'y') {
        # Show options for changing the File Explorer default location
        Do {
            Write-Output ""
            Write-Host "   Options:" -ForegroundColor Yellow
            Write-Host "    (n) No change" -ForegroundColor Yellow
            Write-Host "    (1) Open File Explorer to 'Home'" -ForegroundColor Yellow
            Write-Host "    (2) Open File Explorer to 'This PC'" -ForegroundColor Yellow
            Write-Host "    (3) Open File Explorer to 'Downloads'" -ForegroundColor Yellow
            Write-Host "    (4) Open File Explorer to 'OneDrive'" -ForegroundColor Yellow
            $ExplSearchInput = Read-Host "   Change the default location that File Explorer opens to? (n/1/2/3/4)" 
        }
        while ($ExplSearchInput -ne 'n' -and $ExplSearchInput -ne '0' -and $ExplSearchInput -ne '1' -and $ExplSearchInput -ne '2' -and $ExplSearchInput -ne '3' -and $ExplSearchInput -ne '4') 

        # Select correct taskbar search option based on user input
        switch ($ExplSearchInput) {
            '1' {
                AddParameter 'ExplorerToHome' "Change the default location that File Explorer opens to 'Home'"
            }
            '2' {
                AddParameter 'ExplorerToThisPC' "Change the default location that File Explorer opens to 'This PC'"
            }
            '3' {
                AddParameter 'ExplorerToDownloads' "Change the default location that File Explorer opens to 'Downloads'"
            }
            '4' {
                AddParameter 'ExplorerToOneDrive' "Change the default location that File Explorer opens to 'OneDrive'"
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   Show hidden files, folders and drives? (y/n)" ) -eq 'y') {
            AddParameter 'ShowHiddenFolders' 'Show hidden files, folders and drives'
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   Show file extensions for known file types? (y/n)" ) -eq 'y') {
            AddParameter 'ShowKnownFileExt' 'Show file extensions for known file types'
        }

        # Only show this option for Windows 11 users running build 22000 or later
        if ($WinVersion -ge 22000) {
            Write-Output ""

            if ($( Read-Host -Prompt "   Hide the Home section from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                AddParameter 'HideHome' 'Hide the Home section from the File Explorer sidepanel'
            }

            Write-Output ""

            if ($( Read-Host -Prompt "   Hide the Gallery section from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                AddParameter 'HideGallery' 'Hide the Gallery section from the File Explorer sidepanel'
            }
        }

        Write-Output ""

        if ($( Read-Host -Prompt "   Hide duplicate removable drive entries from the File Explorer sidepanel so they only show under This PC? (y/n)" ) -eq 'y') {
            AddParameter 'HideDupliDrive' 'Hide duplicate removable drive entries from the File Explorer sidepanel'
        }

        # Only show option for disabling these specific folders for Windows 10 users
        if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") {
            Write-Output ""

            if ($( Read-Host -Prompt "Do you want to hide any folders from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                Write-Output ""

                if ($( Read-Host -Prompt "   Hide the OneDrive folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                    AddParameter 'HideOnedrive' 'Hide the OneDrive folder in the File Explorer sidepanel'
                }

                Write-Output ""
                
                if ($( Read-Host -Prompt "   Hide the 3D objects folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                    AddParameter 'Hide3dObjects' "Hide the 3D objects folder under 'This pc' in File Explorer" 
                }
                
                Write-Output ""

                if ($( Read-Host -Prompt "   Hide the music folder from the File Explorer sidepanel? (y/n)" ) -eq 'y') {
                    AddParameter 'HideMusic' "Hide the music folder under 'This pc' in File Explorer"
                }
            }
        }
    }

    # Suppress prompt if Silent parameter was passed
    if (-not $Silent) {
        Write-Output ""
        Write-Output ""
        Write-Output ""
        Write-Output "Press enter to confirm your choices and execute the script or press CTRL+C to quit..."
        Read-Host | Out-Null
    }

    PrintHeader 'Custom Mode'
}



##################################################################################################################
#                                                                                                                #
#                                                  SCRIPT START                                                  #
#                                                                                                                #
##################################################################################################################



# Check if winget is installed & if it is, check if the version is at least v1.4
if ((Get-AppxPackage -Name "*Microsoft.DesktopAppInstaller*") -and ([int](((winget -v) -replace 'v','').split('.')[0..1] -join '') -gt 14)) {
    $script:wingetInstalled = $true
}
else {
    $script:wingetInstalled = $false

    # Show warning that requires user confirmation, Suppress confirmation if Silent parameter was passed
    if (-not $Silent) {
        Write-Warning "Winget is not installed or outdated. This may prevent Win11Debloat from removing certain apps."
        Write-Output ""
        Write-Output "Press any key to continue anyway..."
        $null = [System.Console]::ReadKey()
    }
}

# Get current Windows build version to compare against features
$WinVersion = Get-ItemPropertyValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' CurrentBuild

# Check if the machine supports Modern Standby, this is used to determine if the DisableModernStandbyNetworking option can be used
$script:ModernStandbySupported = CheckModernStandbySupport

$script:Params = $PSBoundParameters
$script:FirstSelection = $true
$SPParams = 'WhatIf', 'Confirm', 'Verbose', 'Silent', 'Sysprep', 'Debug', 'User', 'CreateRestorePoint', 'LogPath'
$SPParamCount = 0

# Count how many SPParams exist within Params
# This is later used to check if any options were selected
foreach ($Param in $SPParams) {
    if ($script:Params.ContainsKey($Param)) {
        $SPParamCount++
    }
}

# Hide progress bars for app removal, as they block Win11Debloat's output
if (-not ($script:Params.ContainsKey("Verbose"))) {
    $ProgressPreference = 'SilentlyContinue'
}
else {
    Write-Host "Verbose mode is enabled"
    Write-Output ""
    Write-Output "Press any key to continue..."
    $null = [System.Console]::ReadKey()

    $ProgressPreference = 'Continue'
}

if ($script:Params.ContainsKey("Sysprep")) {
    $defaultUserPath = GetUserDirectory -userName "Default"

    # Exit script if run in Sysprep mode on Windows 10
    if ($WinVersion -lt 22000) {
        Write-Host "Error: Win11Debloat Sysprep mode is not supported on Windows 10" -ForegroundColor Red
        AwaitKeyToExit
    }
}

# Make sure all requirements for User mode are met, if User is specified
if ($script:Params.ContainsKey("User")) {
    $userPath = GetUserDirectory -userName $script:Params.Item("User")
}

# Remove SavedSettings file if it exists and is empty
if ((Test-Path "$PSScriptRoot/SavedSettings") -and ([String]::IsNullOrWhiteSpace((Get-content "$PSScriptRoot/SavedSettings")))) {
    Remove-Item -Path "$PSScriptRoot/SavedSettings" -recurse
}

# Only run the app selection form if the 'RunAppsListGenerator' parameter was passed to the script
if ($RunAppConfigurator -or $RunAppsListGenerator) {
    PrintHeader "Custom Apps List Generator"

    $result = ShowAppSelectionForm

    # Show different message based on whether the app selection was saved or cancelled
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "Application selection window was closed without saving." -ForegroundColor Red
    }
    else {
        Write-Output "Your app selection was saved to the 'CustomAppsList' file, found at:"
        Write-Host "$PSScriptRoot" -ForegroundColor Yellow
    }

    AwaitKeyToExit
}

# Change script execution based on provided parameters or user input
if ((-not $script:Params.Count) -or $RunDefaults -or $RunDefaultsLite -or $RunSavedSettings -or ($SPParamCount -eq $script:Params.Count)) {
    if ($RunDefaults -or $RunDefaultsLite) {
        $Mode = '1'
    }
    elseif ($RunSavedSettings) {
        if (-not (Test-Path "$PSScriptRoot/SavedSettings")) {
            PrintHeader 'Custom Mode'
            Write-Host "Error: No saved settings found, no changes were made" -ForegroundColor Red
            AwaitKeyToExit
        }

        $Mode = '4'
    }
    else {
        # Show menu and wait for user input, loops until valid input is provided
        Do { 
            $ModeSelectionMessage = "请选择一个选项 (1/2/3/0)" 

            PrintHeader 'Menu'

            Write-Output "(1) 默认模式：快速应用推荐的更改内容"
            Write-Output "(2) 自定义模式：手动选择要进行的更改内容"
            Write-Output "(3) 应用卸载模式：选择并卸载应用，但不进行其他设置更改操作"

            # Only show this option if SavedSettings file exists
            if (Test-Path "$PSScriptRoot/SavedSettings") {
                Write-Output "(4) 应用上次保存的自定义设置"
                
                $ModeSelectionMessage = "请选择一个选项 (1/2/3/4/0)" 
            }

            Write-Output ""
            Write-Output "(0) 展示更多信息"
            Write-Output ""
            Write-Output ""

            $Mode = Read-Host $ModeSelectionMessage

            if ($Mode -eq '0') {
                # Print information screen from file
                PrintFromFile "$PSScriptRoot/Assets/Menus/Info" "Information"

                Write-Output "Press any key to go back..."
                $null = [System.Console]::ReadKey()
            }
            elseif (($Mode -eq '4') -and -not (Test-Path "$PSScriptRoot/SavedSettings")) {
                $Mode = $null
            }
        }
        while ($Mode -ne '1' -and $Mode -ne '2' -and $Mode -ne '3' -and $Mode -ne '4') 
    }

    # Add execution parameters based on the mode
    switch ($Mode) {
        # Default mode, loads defaults after confirmation
        '1' { 
            AddParameter 'CreateRestorePoint' 'Create a system restore point' $false

            # Show the default settings with confirmation, unless Silent parameter was passed
            if (-not $Silent) {
                # Show options for app removal
                if ((-not $RunDefaults) -and (-not $RunDefaultsLite)) {
                    PrintHeader 'Default Mode'

                    Write-Host "Please note: The default selection of apps includes Microsoft Teams, Spotify, Sticky Notes and more. Select option 2 to verify and change what apps are removed by the script." -ForegroundColor DarkGray
                    Write-Output ""

                    Do {
                        Write-Host "Options:" -ForegroundColor Yellow
                        Write-Host " (n) Don't remove any apps" -ForegroundColor Yellow
                        Write-Host " (1) Only remove the default selection of apps" -ForegroundColor Yellow
                        Write-Host " (2) Manually select which apps to remove" -ForegroundColor Yellow
                        $RemoveAppsInput = Read-Host "Do you want to remove any apps? Apps will be removed for all users (n/1/2)"
                
                        # Show app selection form if user entered option 3
                        if ($RemoveAppsInput -eq '2') {
                            $result = ShowAppSelectionForm
                
                            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                                # User cancelled or closed app selection, show error and change RemoveAppsInput so the menu will be shown again
                                Write-Output ""
                                Write-Host "Cancelled application selection, please try again" -ForegroundColor Red
                
                                $RemoveAppsInput = 'c'
                            }
                            
                            Write-Output ""
                        }
                    }
                    while ($RemoveAppsInput -ne 'n' -and $RemoveAppsInput -ne '0' -and $RemoveAppsInput -ne '1' -and $RemoveAppsInput -ne '2') 
                } elseif ($RunDefaultsLite) {
                    $RemoveAppsInput = '0'                
                } else {
                    $RemoveAppsInput = '1'
                }

                PrintHeader 'Default Mode'

                Write-Output "Win11Debloat will make the following changes:"
    
                # Select correct option based on user input
                switch ($RemoveAppsInput) {
                    '1' {
                        AddParameter 'RemoveApps' 'Remove the default selection of apps:' $false
                        PrintAppsList "$PSScriptRoot/Appslist.txt"
                    }
                    '2' {
                        AddParameter 'RemoveAppsCustom' "Remove $($script:SelectedApps.Count) apps:" $false
                        PrintAppsList "$PSScriptRoot/CustomAppsList"
                    }
                }

                 # Only add this option for Windows 10 users
                if (get-ciminstance -query "select caption from win32_operatingsystem where caption like '%Windows 10%'") {
                    AddParameter 'Hide3dObjects' "Hide the 3D objects folder under 'This pc' in File Explorer" $false
                    AddParameter 'HideChat' 'Hide the chat (meet now) icon from the taskbar' $false
                }

                # Only add these options for Windows 11 users (build 22000+)
                if ($WinVersion -ge 22000) {
                    if ($script:ModernStandbySupported) {
                        AddParameter 'DisableModernStandbyNetworking' 'Disable network connectivity during Modern Standby' $false
                    }

                    AddParameter 'DisableRecall' 'Disable Windows Recall' $false
                    AddParameter 'DisableClickToDo' 'Disable Click to Do (AI text & image analysis)' $false
                } 

                PrintFromFile "$PSScriptRoot/Assets/Menus/DefaultSettings" "Default Mode" $false
        
                Write-Output "Press enter to execute the script or press CTRL+C to quit..."
                Read-Host | Out-Null
            }

            $DefaultParameterNames = 'DisableCopilot','DisableTelemetry','DisableSuggestions','DisableEdgeAds','DisableLockscreenTips','DisableBing','ShowKnownFileExt','DisableWidgets','DisableFastStartup'

            PrintHeader 'Default Mode'

            # Add default parameters, if they don't already exist
            foreach ($ParameterName in $DefaultParameterNames) {
                if (-not $script:Params.ContainsKey($ParameterName)) {
                    $script:Params.Add($ParameterName, $true)
                }
            }
        }

        # Custom mode, show & add options based on user input
        '2' { 
            DisplayCustomModeOptions
        }

        # App removal, remove apps based on user selection
        '3' {
            PrintHeader "App Removal"

            $result = ShowAppSelectionForm

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Output "You have selected $($script:SelectedApps.Count) apps for removal"
                AddParameter 'RemoveAppsCustom' "Remove $($script:SelectedApps.Count) apps:"

                # Suppress prompt if Silent parameter was passed
                if (-not $Silent) {
                    Write-Output ""
                    Write-Output ""
                    Write-Output "Press enter to remove the selected apps or press CTRL+C to quit..."
                    Read-Host | Out-Null
                    PrintHeader "App Removal"
                }
            }
            else {
                Write-Host "Selection was cancelled, no apps have been removed" -ForegroundColor Red
                Write-Output ""
            }
        }

        # Load custom options from the "SavedSettings" file
        '4' {
            PrintHeader 'Custom Mode'
            Write-Output "Win11Debloat will make the following changes:"

            # Print the saved settings info from file
            Foreach ($line in (Get-Content -Path "$PSScriptRoot/SavedSettings" )) { 
                # Remove any spaces before and after the line
                $line = $line.Trim()
            
                # Check if the line contains a comment
                if (-not ($line.IndexOf('#') -eq -1)) {
                    $parameterName = $line.Substring(0, $line.IndexOf('#'))

                    # Print parameter description and add parameter to Params list
                    switch ($parameterName) {
                        'RemoveApps' {
                            PrintAppsList "$PSScriptRoot/Appslist.txt" $true
                        }
                        'RemoveAppsCustom' {
                            PrintAppsList "$PSScriptRoot/CustomAppsList" $true
                        }
                        default {
                            Write-Output $line.Substring(($line.IndexOf('#') + 1), ($line.Length - $line.IndexOf('#') - 1))
                        }
                    }

                    if (-not $script:Params.ContainsKey($parameterName)) {
                        $script:Params.Add($parameterName, $true)
                    }
                }
            }

            if (-not $Silent) {
                Write-Output ""
                Write-Output ""
                Write-Output "Press enter to execute the script or press CTRL+C to quit..."
                Read-Host | Out-Null
            }

            PrintHeader 'Custom Mode'
        }
    }
}
else {
    PrintHeader 'Custom Mode'
}

# If the number of keys in SPParams equals the number of keys in Params then no modifications/changes were selected
#  or added by the user, and the script can exit without making any changes.
if ($SPParamCount -eq $script:Params.Keys.Count) {
    Write-Output "The script completed without making any changes."

    AwaitKeyToExit
}

# Execute all selected/provided parameters
switch ($script:Params.Keys) {
    'CreateRestorePoint' {
        CreateSystemRestorePoint
        continue
    }
    'RemoveApps' {
        $appsList = ReadAppslistFromFile "$PSScriptRoot/Appslist.txt" 
        Write-Output "> Removing default selection of $($appsList.Count) apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveAppsCustom' {
        if (-not (Test-Path "$PSScriptRoot/CustomAppsList")) {
            Write-Host "> Error: Could not load custom apps list from file, no apps were removed" -ForegroundColor Red
            Write-Output ""
            continue
        }
        
        $appsList = ReadAppslistFromFile "$PSScriptRoot/CustomAppsList"
        Write-Output "> Removing $($appsList.Count) apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveCommApps' {
        $appsList = 'Microsoft.windowscommunicationsapps', 'Microsoft.People'
        Write-Output "> Removing Mail, Calendar and People apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveW11Outlook' {
        $appsList = 'Microsoft.OutlookForWindows'
        Write-Output "> Removing new Outlook for Windows app..."
        RemoveApps $appsList
        continue
    }
    'RemoveGamingApps' {
        $appsList = 'Microsoft.GamingApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay'
        Write-Output "> Removing gaming related apps..."
        RemoveApps $appsList
        continue
    }
    'RemoveHPApps' {
        $appsList = 'AD2F1837.HPAIExperienceCenter', 'AD2F1837.HPJumpStarts', 'AD2F1837.HPPCHardwareDiagnosticsWindows', 'AD2F1837.HPPowerManager', 'AD2F1837.HPPrivacySettings', 'AD2F1837.HPSupportAssistant', 'AD2F1837.HPSureShieldAI', 'AD2F1837.HPSystemInformation', 'AD2F1837.HPQuickDrop', 'AD2F1837.HPWorkWell', 'AD2F1837.myHP', 'AD2F1837.HPDesktopSupportUtilities', 'AD2F1837.HPQuickTouch', 'AD2F1837.HPEasyClean', 'AD2F1837.HPConnectedMusic', 'AD2F1837.HPFileViewer', 'AD2F1837.HPRegistration', 'AD2F1837.HPWelcome', 'AD2F1837.HPConnectedPhotopoweredbySnapfish', 'AD2F1837.HPPrinterControl'
        Write-Output "> Removing HP apps..."
        RemoveApps $appsList
        continue
    }
    "ForceRemoveEdge" {
        ForceRemoveEdge
        continue
    }
    'DisableDVR' {
        RegImport "> Disabling Xbox game/screen recording..." "Disable_DVR.reg"
        continue
    }
    'DisableTelemetry' {
        RegImport "> Disabling telemetry, diagnostic data, activity history, app-launch tracking and targeted ads..." "Disable_Telemetry.reg"
        continue
    }
    {$_ -in "DisableSuggestions", "DisableWindowsSuggestions"} {
        RegImport "> Disabling tips, tricks, suggestions and ads across Windows..." "Disable_Windows_Suggestions.reg"
        continue
    }
    'DisableEdgeAds' {
        RegImport "> Disabling ads, suggestions and the MSN news feed in Microsoft Edge..." "Disable_Edge_Ads_And_Suggestions.reg"
        continue
    }
    {$_ -in "DisableLockscrTips", "DisableLockscreenTips"} {
        RegImport "> Disabling tips & tricks on the lockscreen..." "Disable_Lockscreen_Tips.reg"
        continue
    }
    'DisableDesktopSpotlight' {
        RegImport "> Disabling the 'Windows Spotlight' desktop background option..." "Disable_Desktop_Spotlight.reg"
        continue
    }
    'DisableSettings365Ads' {
        RegImport "> Disabling Microsoft 365 ads in Settings Home..." "Disable_Settings_365_Ads.reg"
        continue
    }
    'DisableSettingsHome' {
        RegImport "> Disabling the Settings Home page..." "Disable_Settings_Home.reg"
        continue
    }
    {$_ -in "DisableBingSearches", "DisableBing"} {
        RegImport "> Disabling Bing web search, Bing AI and Cortana from Windows search..." "Disable_Bing_Cortana_In_Search.reg"
        
        # Also remove the app package for Bing search
        $appsList = 'Microsoft.BingSearch'
        RemoveApps $appsList
        continue
    }
    'DisableCopilot' {
        RegImport "> Disabling Microsoft Copilot..." "Disable_Copilot.reg"

        # Also remove the app package for Copilot
        $appsList = 'Microsoft.Copilot'
        RemoveApps $appsList
        continue
    }
    'DisableRecall' {
        RegImport "> Disabling Windows Recall..." "Disable_AI_Recall.reg"
        continue
    }
    'DisableClickToDo' {
        RegImport "> Disabling Click to Do..." "Disable_Click_to_Do.reg"
        continue
    }
    'DisableEdgeAI' {
        RegImport "> Disabling AI features in Microsoft Edge..." "Disable_Edge_AI_Features.reg"
        continue
    }
    'DisablePaintAI' {
        RegImport "> Disabling AI features in Paint..." "Disable_Paint_AI_Features.reg"
        continue
    }
    'DisableNotepadAI' {
        RegImport "> Disabling AI features in Notepad..." "Disable_Notepad_AI_Features.reg"
        continue
    }
    'RevertContextMenu' {
        RegImport "> Restoring the old Windows 10 style context menu..." "Disable_Show_More_Options_Context_Menu.reg"
        continue
    }
    'DisableMouseAcceleration' {
        RegImport "> Turning off Enhanced Pointer Precision..." "Disable_Enhance_Pointer_Precision.reg"
        continue
    }
    'DisableStickyKeys' {
        RegImport "> Disabling the Sticky Keys keyboard shortcut..." "Disable_Sticky_Keys_Shortcut.reg"
        continue
    }
    'DisableFastStartup' {
        RegImport "> Disabling Fast Start-up..." "Disable_Fast_Startup.reg"
        continue
    }
    'DisableModernStandbyNetworking' {
        RegImport "> Disabling network connectivity during Modern Standby..." "Disable_Modern_Standby_Networking.reg"
        continue
    }
    'ClearStart' {
        Write-Output "> Removing all pinned apps from the start menu for user $(GetUserName)..."
        ReplaceStartMenu
        Write-Output ""
        continue
    }
    'ReplaceStart' {
        Write-Output "> Replacing the start menu for user $(GetUserName)..."
        ReplaceStartMenu $script:Params.Item("ReplaceStart")
        Write-Output ""
        continue
    }
    'ClearStartAllUsers' {
        ReplaceStartMenuForAllUsers
        continue
    }
    'ReplaceStartAllUsers' {
        ReplaceStartMenuForAllUsers $script:Params.Item("ReplaceStartAllUsers")
        continue
    }
    'DisableStartRecommended' {
        RegImport "> Disabling the start menu recommended section..." "Disable_Start_Recommended.reg"
        continue
    }
    'DisableStartPhoneLink' {
        RegImport "> Disabling the Phone Link mobile devices integration in the start menu..." "Disable_Phone_Link_In_Start.reg"
        continue
    }
    'EnableDarkMode' {
        RegImport "> Enabling dark mode for system and apps..." "Enable_Dark_Mode.reg"
        continue
    }
    'DisableTransparency' {
        RegImport "> Disabling transparency effects..." "Disable_Transparency.reg"
        continue
    }
    'DisableAnimations' {
        RegImport "> Disabling animations and visual effects..." "Disable_Animations.reg"
        continue
    }
    'TaskbarAlignLeft' {
        RegImport "> Aligning taskbar buttons to the left..." "Align_Taskbar_Left.reg"
        continue
    }
    'CombineTaskbarAlways' {
        RegImport "> Setting the taskbar on the main display to always combine buttons and hide labels..." "Combine_Taskbar_Always.reg"
        continue
    }
    'CombineTaskbarWhenFull' {
        RegImport "> Setting the taskbar on the main display to only combine buttons and hide labels when the taskbar is full..." "Combine_Taskbar_When_Full.reg"
        continue
    }
    'CombineTaskbarNever' {
        RegImport "> Setting the taskbar on the main display to never combine buttons or hide labels..." "Combine_Taskbar_Never.reg"
        continue
    }
    'CombineMMTaskbarAlways' {
        RegImport "> Setting the taskbar on secondary displays to always combine buttons and hide labels..." "Combine_MMTaskbar_Always.reg"
        continue
    }
    'CombineMMTaskbarWhenFull' {
        RegImport "> Setting the taskbar on secondary displays to only combine buttons and hide labels when the taskbar is full..." "Combine_MMTaskbar_When_Full.reg"
        continue
    }
    'CombineMMTaskbarNever' {
        RegImport "> Setting the taskbar on secondary displays to never combine buttons or hide labels..." "Combine_MMTaskbar_Never.reg"
        continue
    }
    'MMTaskbarModeAll' {
        RegImport "> Setting the taskbar to only show app icons on main taskbar..." "MMTaskbarMode_All.reg"
        continue
    }
    'MMTaskbarModeMainActive' {
        RegImport "> Setting the taskbar to show app icons on all taskbars..." "MMTaskbarMode_Main_Active.reg"
        continue
    }
    'MMTaskbarModeActive' {
        RegImport "> Setting the taskbar to only show app icons on the taskbar where the window is open..." "MMTaskbarMode_Active.reg"
        continue
    }
    'HideSearchTb' {
        RegImport "> Hiding the search icon from the taskbar..." "Hide_Search_Taskbar.reg"
        continue
    }
    'ShowSearchIconTb' {
        RegImport "> Changing taskbar search to icon only..." "Show_Search_Icon.reg"
        continue
    }
    'ShowSearchLabelTb' {
        RegImport "> Changing taskbar search to icon with label..." "Show_Search_Icon_And_Label.reg"
        continue
    }
    'ShowSearchBoxTb' {
        RegImport "> Changing taskbar search to search box..." "Show_Search_Box.reg"
        continue
    }
    'HideTaskview' {
        RegImport "> Hiding the taskview button from the taskbar..." "Hide_Taskview_Taskbar.reg"
        continue
    }
    {$_ -in "HideWidgets", "DisableWidgets"} {
        RegImport "> Disabling widgets on the taskbar & lockscreen..." "Disable_Widgets_Service.reg"

        # Also remove the app package for Widgets
        $appsList = 'Microsoft.StartExperiencesApp'
        RemoveApps $appsList
        continue
    }
    {$_ -in "HideChat", "DisableChat"} {
        RegImport "> Hiding the chat icon from the taskbar..." "Disable_Chat_Taskbar.reg"
        continue
    }
    'EnableEndTask' {
        RegImport "> Enabling the 'End Task' option in the taskbar right click menu..." "Enable_End_Task.reg"
        continue
    }
    'EnableLastActiveClick' {
        RegImport "> Enabling the 'Last Active Click' behavior in the taskbar app area..." "Enable_Last_Active_Click.reg"
        continue
    }
    'ExplorerToHome' {
        RegImport "> Changing the default location that File Explorer opens to `Home`..." "Launch_File_Explorer_To_Home.reg"
        continue
    }
    'ExplorerToThisPC' {
        RegImport "> Changing the default location that File Explorer opens to `This PC`..." "Launch_File_Explorer_To_This_PC.reg"
        continue
    }
    'ExplorerToDownloads' {
        RegImport "> Changing the default location that File Explorer opens to `Downloads`..." "Launch_File_Explorer_To_Downloads.reg"
        continue
    }
    'ExplorerToOneDrive' {
        RegImport "> Changing the default location that File Explorer opens to `OneDrive`..." "Launch_File_Explorer_To_OneDrive.reg"
        continue
    }
    'ShowHiddenFolders' {
        RegImport "> Unhiding hidden files, folders and drives..." "Show_Hidden_Folders.reg"
        continue
    }
    'ShowKnownFileExt' {
        RegImport "> Enabling file extensions for known file types..." "Show_Extensions_For_Known_File_Types.reg"
        continue
    }
    'HideHome' {
        RegImport "> Hiding the home section from the File Explorer navigation pane..." "Hide_Home_from_Explorer.reg"
        continue
    }
    'HideGallery' {
        RegImport "> Hiding the gallery section from the File Explorer navigation pane..." "Hide_Gallery_from_Explorer.reg"
        continue
    }
    'HideDupliDrive' {
        RegImport "> Hiding duplicate removable drive entries from the File Explorer navigation pane..." "Hide_duplicate_removable_drives_from_navigation_pane_of_File_Explorer.reg"
        continue
    }
    {$_ -in "HideOnedrive", "DisableOnedrive"} {
        RegImport "> Hiding the OneDrive folder from the File Explorer navigation pane..." "Hide_Onedrive_Folder.reg"
        continue
    }
    {$_ -in "Hide3dObjects", "Disable3dObjects"} {
        RegImport "> Hiding the 3D objects folder from the File Explorer navigation pane..." "Hide_3D_Objects_Folder.reg"
        continue
    }
    {$_ -in "HideMusic", "DisableMusic"} {
        RegImport "> Hiding the music folder from the File Explorer navigation pane..." "Hide_Music_folder.reg"
        continue
    }
    {$_ -in "HideIncludeInLibrary", "DisableIncludeInLibrary"} {
        RegImport "> Hiding 'Include in library' in the context menu..." "Disable_Include_in_library_from_context_menu.reg"
        continue
    }
    {$_ -in "HideGiveAccessTo", "DisableGiveAccessTo"} {
        RegImport "> Hiding 'Give access to' in the context menu..." "Disable_Give_access_to_context_menu.reg"
        continue
    }
    {$_ -in "HideShare", "DisableShare"} {
        RegImport "> Hiding 'Share' in the context menu..." "Disable_Share_from_context_menu.reg"
        continue
    }
}

RestartExplorer

Write-Output ""
Write-Output ""
Write-Output ""
Write-Output "Script completed! Please check above for any errors."

AwaitKeyToExit
