$fileName = Split-Path $comboboxlocalFilePath.Text -Leaf
    $Response = Invoke-RestMethod -Uri "https://api.threatstream.com/api/v1/submit/search/?q=$script:Note" -Headers @{"Authorization" = "apikey redacted"} -Method Get 
    if ($Response.meta.total_count -gt 0) {
        $textboxResults.AppendText("Total Count: " + $Response.meta.total_count + [Environment]::NewLine)
        foreach ($report in $Response.objects) {
            $textboxResults.AppendText(($report | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
        }
    } else {
        $textboxResults.AppendText("No reports found for file: $fileName. Please wait and try again.")
    }