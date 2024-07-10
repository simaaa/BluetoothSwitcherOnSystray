# Default variables
$TITLE = "Bluetooth switcher"
$SYSTRAY_TOOLTIP = "Bluetooth switcher"
$ICON_FILE_ON = ".\BluetoothStatusOn.ico"
$ICON_FILE_OFF = ".\BluetoothStatusOff.ico"
$HIDE_CONSOLE = $false
$HIDE_CONSOLE = $true

#Load necessary assembilies
Add-Type -AssemblyName 'System.Windows.Forms'
Add-Type -Name Window -Namespace Console -MemberDefinition '
  [DllImport("Kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

function Console-Show {
    $PSConsole = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($PSConsole, 5)
}

function Console-Hide {
    $PSConsole = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($PSConsole, 0)
}

function Show-Balloon {
    Param(
		[string] $Title = $TITLE, $Text
    )
	$SystrayLauncher.BalloonTipTitle = $Title
	$SystrayLauncher.BalloonTipText = $Text
	$SystrayLauncher.ShowBalloonTip(1000) #The time period, in milliseconds, the balloon tip should display
}

## BLUETOOTH
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
Function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
function Set-Bluetooth-Status($BluetoothStatus) {
	[Windows.Devices.Radios.Radio,Windows.System.Devices,ContentType=WindowsRuntime] | Out-Null
	#[Windows.Devices.Radios.RadioAccessStatus,Windows.System.Devices,ContentType=WindowsRuntime] | Out-Null
	Await ([Windows.Devices.Radios.Radio]::RequestAccessAsync()) ([Windows.Devices.Radios.RadioAccessStatus]) | Out-Null
	$radios = Await ([Windows.Devices.Radios.Radio]::GetRadiosAsync()) ([System.Collections.Generic.IReadOnlyList[Windows.Devices.Radios.Radio]])
	$bluetooth = $radios | ? { $_.Kind -eq 'Bluetooth' }
	if ($BluetoothStatus -ne $null) {
		Write-Host -NoNewline "[Set-Bluetooth-Status] Bluetooth status before: "
		Write-Host -NoNewline -ForegroundColor Yellow ($bluetooth.State | Out-String)
		Write-Host "[Set-Bluetooth-Status] Switch Bluetooth status..."
		[Windows.Devices.Radios.RadioState,Windows.System.Devices,ContentType=WindowsRuntime] | Out-Null
		Await ($bluetooth.SetStateAsync($BluetoothStatus)) ([Windows.Devices.Radios.RadioAccessStatus]) | Out-Null
	}
	Write-Host -NoNewline "[Set-Bluetooth-Status] Bluetooth status after: "
	Write-Host -NoNewline -ForegroundColor Yellow ($bluetooth.State | Out-String)
	return $bluetooth.State
}

function Set-Systray-Icon($SystrayLauncher, $BTStatus) {
	Write-Host "[Set-Systray-Icon] BTStatus=$BTStatus"
	if ($SystrayLauncher -eq $null) {
		Write-Host "[Set-Systray-Icon] SystrayLauncher IS NULL"
	}
	if ($BTStatus -eq $null) {
		$BTStatus = Set-Bluetooth-Status
		Write-Host -NoNewline "[Set-Systray-Icon] Bluetooth status: "
		Write-Host -ForegroundColor Yellow $BTStatus
	}
	Write-Host -NoNewline "[Set-Systray-Icon] Set Systray icon to "
	if ($BTStatus -eq 'On') {
		Write-Host -ForegroundColor Yellow "ON"
		$SystrayLauncher.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ICON_FILE_ON)
	} else {
		Write-Host -ForegroundColor Yellow "OFF"
		$SystrayLauncher.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ICON_FILE_OFF)
	}
}

function New-MenuItem{
    Param(
        [string] $Text, $Program, $Arguments,
        [switch] $ExitOnly = $false
    )
    
    #Initialization
    $MenuItem = New-Object System.Windows.Forms.MenuItem
	$MenuItem.Text = $Text
	
    #Apply click event logic
    if($Program -and !$ExitOnly){
		$MenuItem | Add-Member -Name Program -Value $Program -MemberType NoteProperty
		$MenuItem | Add-Member -Name Arguments -Value $Arguments -MemberType NoteProperty
        $MenuItem.Add_Click({
            try{
                $Text = $This.Text
                $Program = $This.Program
				$Arguments = $This.Arguments
                Write-Host ""
				Write-Host "[MenuItem_Click] ----> Processing: ""$Text"""
                Write-Host "[MenuItem_Click]   Program=$Program"
				Write-Host "[MenuItem_Click]   Arguments=$Arguments"
                #if(Test-Path $Program) {
				#	Show-Balloon -Text "Launching '$Program' with '$Arguments' arguments..."
                #    if ([IO.Path]::GetExtension($Program) -eq '.ps1') {
				#		Start-Process -FilePath "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-NoProfile -NoLogo -ExecutionPolicy Bypass -File `"$Program`" $Arguments" -ErrorAction Stop
				#	} else {
				#		if($Arguments) {
				#			Start-Process -FilePath "$Program" -ArgumentList " $Arguments" -ErrorAction Stop
				#		} else {
				#			Start-Process -FilePath "$Program" -ErrorAction Stop
				#		}
				#	}
                #} else {
                #    throw "Could not find program: '$Program'"
                #}
				Write-Host "[MenuItem_Click] Set Bluetooth status to: ""$Arguments"""
				$bluetooth_status = Set-Bluetooth-Status $Arguments
				Write-Host "[MenuItem_Click] Set SystrayLauncher icon to: ""$bluetooth_status"""
				Set-Systray-Icon $SystrayLauncher $bluetooth_status
            } catch {
                $Text = $This.Text
                [System.Windows.Forms.MessageBox]::Show("Failed to launch $Text`n`n$_") > $null
            }
        })
    }
    
    #Provide a way to exit the launcher
    if($ExitOnly -and !$Program){
        $MenuItem.Add_Click({
            $Form.Close()
            #Handle any hung processes
            Stop-Process $PID
        })
    }
    
    #Return our new MenuItem
    $MenuItem
}

#Create Form to serve as a container for our components
$Form = New-Object System.Windows.Forms.Form

#Configure our form to be hidden
$Form.BackColor = "Magenta" #Pick a color you won't use again and match it to the TransparencyKey property
$Form.TransparencyKey = "Magenta"
$Form.ShowInTaskbar = $false
$Form.FormBorderStyle = "None"

#Initialize/configure necessary components
$SystrayLauncher = New-Object System.Windows.Forms.NotifyIcon
$SystrayLauncher.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ICON_FILE_ON)
$SystrayLauncher.Text = $SYSTRAY_TOOLTIP
$SystrayLauncher.Visible = $true

#Create context menu items
$Program1= New-MenuItem -Text "Bluetooth ON" -Program ".\BluetoothSwitch.ps1" -Arguments "ON" 
$Program2= New-MenuItem -Text "Bluetooth OFF" -Program ".\BluetoothSwitch.ps1" -Arguments "OFF" 
$ExitLauncher = New-MenuItem -Text "Exit" -ExitOnly

#Add menu items to context menu
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$ContextMenu.MenuItems.AddRange($Program1)
$ContextMenu.MenuItems.AddRange($Program2)
$ContextMenu.MenuItems.AddRange($ExitLauncher)

#Add components to our form
$SystrayLauncher.ContextMenu = $ContextMenu

Set-Systray-Icon $SystrayLauncher

#Launch
Show-Balloon -Title $TITLE -Text "Right click to show menu"
if($HIDE_CONSOLE) { Console-Hide }
$Form.ShowDialog() > $null
Console-Show
