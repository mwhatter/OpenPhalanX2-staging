$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

# Trim trailing backslash if it exists
$filePath = $textboxremoteFilePath.Text.TrimEnd('\')

$filecopy = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Copied $filePath from $computerName at $filecopy" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "Copied $filePath from $computerName"

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($filePath)) {
    # Convert drive letter path to UNC format
    $driveLetter = $filePath.Substring(0, 1)
    $uncPath = "\\$computerName\$driveLetter$" + $filePath.Substring(2)
    try {
        if ((Test-Path -Path $uncPath) -and (Get-Item -Path $uncPath).PSIsContainer) {
            # It's a directory
            $copyDecision = [System.Windows.Forms.MessageBox]::Show("Do you want to copy the directory and all its child directories? Select No to only copy files from specified directory", "Copy Directory Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($copyDecision -eq "Yes") {
                # Copy directory and child directories
                Copy-Item -Path $uncPath -Destination $CopiedFilesDir -Force -Recurse
            } elseif ($copyDecision -eq "No") {
                # Copy only the files in the directory, not any subdirectories
                Get-ChildItem -Path $uncPath -File | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $CopiedFilesDir -Force
                }
            } else {
                # Cancel operation
                return
            }
        } else {
            # It's a file
            Copy-Item -Path $uncPath -Destination $CopiedFilesDir -Force
        }

        [System.Windows.Forms.MessageBox]::Show("File '$filePath' copied to local directory.", "Copy File Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $textboxResults.AppendText("File '$filePath' copied to local directory.`r`n")
        $filecopy = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Copied $filePath from $computerName at $filecopy" -ForegroundColor Cyan 
        Log_Message -logfile $logfile -Message "Copied $filePath from $computerName"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error copying file '$filePath' to local directory: $_", "Copy File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $textboxResults.AppendText("Error copying file '$filePath' to local directory: $_`r`n")
    }
}

