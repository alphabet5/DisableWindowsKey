$code = @'
    [DllImport("user32.dll")]
     public static extern IntPtr GetForegroundWindow();
'@
Add-Type $code -Name Utils -Namespace Win32
$settings = Import-Csv -Path .\DisableWindowsKeySettings.txt
$start_count = 0
$stop_count = 0
$registry_path = "Registry::HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
if ((Get-Item -Path $registry_path).GetValue("NoWinKeys") -eq $null){
    New-ItemProperty -Path $registry_path -Name NoWinKeys -Value 0 -PropertyType DWORD -Force | Out-Null
}

while(1){
    $hwnd = [Win32.Utils]::GetForegroundWindow()
    $process = Get-Process | Where-Object { $_.mainWindowHandle -eq $hwnd }
    $guac_not_running = $true
    ForEach ($row in $settings) {
        if ($process.ProcessName -eq $row.ProcessName -and $process.MainWindowTitle -like $row.MainWindowTitle){
            $start_count += 1
            if ($start_count -gt 20){
                if ((Get-ItemPropertyValue $registry_path -Name NoWinKeys) -ne 1){
                    New-ItemProperty -Path $registry_path -Name NoWinKeys -Value 1 -PropertyType DWORD -Force | Out-Null
                    Stop-Process -ProcessName explorer
                    sleep 2
                }
            }
            $guac_not_running = $false
            $stop_count = 0
        }
    }
    if ($guac_not_running){
    $start_count = 0
    $stop_count += 1
    }
    if ($stop_count -gt 20){
        if ((Get-ItemPropertyValue $registry_path -Name NoWinKeys) -ne 0){
            New-ItemProperty -Path $registry_path -Name NoWinKeys -Value 0 -PropertyType DWORD -Force | Out-Null
            Stop-Process -ProcessName explorer
        }
    }
    sleep -Milliseconds 100
}