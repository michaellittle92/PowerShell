# Ensure AnyDesk processes are stopped
Get-Process | Where-Object { $_.ProcessName -like "*AnyDesk*" } | Stop-Process -Force -ErrorAction SilentlyContinue

# Stop AnyDesk Service if running
$serviceName = 'AnyDesk' # Adjust the service name accordingly
if (Get-Service -Name $serviceName -ErrorAction SilentlyContinue) {
    Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
    Write-Host "AnyDesk service stopped."
}

 Set-Location "C:\Program Files (x86)\AnyDesk"

# Execute the uninstaller silently
& .\AnyDesk.exe --silent --remove -wait

# Specify the full path to the folder you want to delete
$folderPath = "C:\Program Files (x86)\AnyDesk"

# Attempt to delete the folder
if (Test-Path $folderPath) {
    Remove-Item -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $folderPath) {
        Write-Host "Failed to delete the folder. It might still be in use or require higher permissions."
    } else {
        Write-Host "Folder deleted successfully."
    }
} else {
    Write-Host "Folder does not exist."
}
