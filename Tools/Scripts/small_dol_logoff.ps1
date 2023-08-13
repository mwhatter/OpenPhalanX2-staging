$selectedUser = if ($comboboxUsername.SelectedItem) {
    $comboboxUsername.SelectedItem.ToString()
} else {
    $comboboxUsername.Text
}
$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

# Function to remotely execute commands and return the result
function Execute-RemoteCommand {
    param (
        [string]$ComputerName,
        [string]$Command
    )
    try {
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            param($Command)
            
            $processInfo = New-Object System.Diagnostics.ProcessStartInfo
            $processInfo.FileName = "powershell.exe"
            $processInfo.Arguments = "-NoProfile -ExecutionPolicy Bypass -Command $Command"
            $processInfo.RedirectStandardOutput = $true
            $processInfo.RedirectStandardError = $true
            $processInfo.UseShellExecute = $false
            $processInfo.CreateNoWindow = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $processInfo
            $process.Start() | Out-Null

            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()

            $process.WaitForExit()

            if ($stderr) {
                return "Error: $stderr"
            } elseif ($stdout) {
                return $stdout
            }

            return $null
        } -ArgumentList $Command

        return $result
    }
    catch {
        return "Error executing command: $_"
    }
}

# Fetch the session ID(s) of the specified user on the remote machine
$sessionIdCommandOutput = Execute-RemoteCommand -ComputerName $computerName -Command "query user"

$sessionIds = $sessionIdCommandOutput | Where-Object { $_ -match "\s+(\d+)\s+($selectedUser)" } | ForEach-Object { if ($matches[1]) { $matches[1] } }

foreach ($sessionId in $sessionIds) {
    # Logoff the user using the retrieved session ID on the remote machine
    $logoffResult = Execute-RemoteCommand -ComputerName $computerName -Command "logoff $sessionId"
    # You can use $logoffResult for any additional logging or processing as needed
}

$gotlogoffusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "$Username logged $selectedUser off from $computerName at $gotlogoffusr" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "$Username logged $selectedUser off from $computerName"
$textboxResults.AppendText("$Username logged $selectedUser off from $computerName at $gotlogoffusr`r`n")
