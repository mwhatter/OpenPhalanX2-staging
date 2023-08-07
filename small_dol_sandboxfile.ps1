$sampleFile = if ($comboboxlocalFilePath.SelectedItem) {
    $comboboxlocalFilePath.SelectedItem.ToString()
} else {
    $comboboxlocalFilePath.Text
}

$pythonScript = "submit_file_anomali.py"
$pythonExe = "python"
$apiKey = "redacted"

$output = & $pythonExe $pythonScript $sampleFile $apiKey
$textboxResults.AppendText("Python script output: $output `r`n")

$outputObject = $output | ConvertFrom-Json
$script:Note = $outputObject.Note
$textboxResults.AppendText("Tags added: $script:Note `r`n")

if ($script:Note -ne $null) {
    $buttonRetrieveReportfile.Enabled = $true
}