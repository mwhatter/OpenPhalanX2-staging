$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}
$filePath = $textboxremoteFilePath.Text

if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($filePath)) {
    # Extract filename from path
    $filename = [System.IO.Path]::GetFileName($filePath)
    try {
        $result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete '$filePath'?", "Delete File or Directory", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq "Yes") {
            Invoke-Command -ComputerName $computerName -ScriptBlock { param($path) Remove-Item -Path $path -Recurse -Force -ErrorAction Stop } -ArgumentList $filePath
            [System.Windows.Forms.MessageBox]::Show("'$filename' deleted from remote host.", "Delete Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $textboxResults.AppendText("'$filename' deleted from remote host.`r`n")
        }
        $filedelete = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Deleted $filePath from $computerName at $filedelete" -ForegroundColor Cyan 
        Log_Message -logfile $logfile -Message "Deleted $filePath from $computerName"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error deleting '$filename' from remote host: $_", "Delete Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $textboxResults.AppendText("Error deleting '$filename' from remote host: $_`r`n")
    }
}