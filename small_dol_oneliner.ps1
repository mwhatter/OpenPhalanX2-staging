#Aligned and ready for testing
param (
    [string]$ComputerNameFromMain,
    [string]$CommandFromMain
)

# Initial output
$outputMessage = ""

# Check for computer name
if (-not $ComputerNameFromMain) {
    return "Please enter a remote computer name."
}

# Check for command
if (-not $CommandFromMain) {
    return "Please enter a command to execute."
}

try {
    $result = Invoke-Command -ComputerName $ComputerNameFromMain -ScriptBlock {
        param($Command)
        $output = $null

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
    } -ArgumentList $CommandFromMain

    if ($result -and $result.Trim()) {
        $outputMessage = "$CommandFromMain executed. $ComputerNameFromMain responded with:`n$result"
    } else {
        $outputMessage = "$CommandFromMain executed on $ComputerNameFromMain but there's no response to display."
    }
}
catch {
    $ErrorMessage = $_.Exception.Message
    $outputMessage = "Error executing command: $ErrorMessage"
}

# Return the output message to be used in main.ps1
return $outputMessage
