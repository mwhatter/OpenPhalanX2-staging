$CopiedFilesPath = ".\CopiedFiles"
    
if (![string]::IsNullOrEmpty($CopiedFilesPath)) {
    try {
        # Use -Recurse to get all nested files
        $files = Get-ChildItem -Path $CopiedFilesPath -File -Recurse | Select-Object -ExpandProperty FullName

        $textboxResults.AppendText(($files -join "`r`n") + "`r`n")
        $comboboxlocalFilePath.Items.Clear()
        foreach ($file in $files) {
            $comboboxlocalFilePath.Items.Add($file)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to list files in the CopiedFiles directory: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Please enter a valid CopiedFiles path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
