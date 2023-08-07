$TextBoxUrlt = $TextBoxUrl.Text
    $sha256 = New-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    $hash = [System.BitConverter]::ToString($sha256.ComputeHash([Text.Encoding]::UTF8.GetBytes($TextBoxUrl)))
    $hash = $hash -replace '-', '' # Remove dashes for a cleaner hash
    $time = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $script:note = "Submitted on $time, URL: $TextBoxUrlt, SHA256: $hash"
    $script:virusTotalId = ""
    $script:pulsediveQid = ""
    # Anomali API
    $queryAnomali = [System.Windows.Forms.MessageBox]::Show("Do you want to submit URL to Anomali API?", "Anomali API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($queryAnomali -eq 'Yes') {
        $response = Invoke-RestMethod -Uri "https://api.threatstream.com/api/v1/submit/new/" -Headers @{"Authorization" = "apikey redacted"} -Method Post -Body @{ "use_vmray_sandbox" = "true"; "vmray_max_jobs" = "3"; "report_radio-classification" = "private"; "report_radio-url" = "$TextBoxUrlt"; "report_radio-notes" = "$note" } -ContentType "application/x-www-form-urlencoded"
        $script:anomaliId = $response.reports.AUTOMATIC.id
        $textboxResults.AppendText("URL submitted to Anomali. Response: " + [Environment]::NewLine + ($response | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
        $buttonRetrieveReport.Enabled = $true
        $script:includeanomalireport = $true
    } elseif ($queryAnomali -eq 'Cancel') {
        return
    }

    # VirusTotal API
    $queryVirusTotal = [System.Windows.Forms.MessageBox]::Show("Do you want to submit URL to VirusTotal API?", "VirusTotal API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($queryVirusTotal -eq 'Yes') {
        $headers=@{}
        $headers.Add("accept", "application/json")
        $headers.Add("x-apikey", "redacted")
        $headers.Add("content-type", "application/x-www-form-urlencoded")
        $response = Invoke-RestMethod -Uri 'https://www.virustotal.com/api/v3/urls' -Method POST -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body "url=$TextBoxUrlt"
        $script:virusTotalId = $response.data.id
        #$script:virusTotalId = $response
        $textboxResults.AppendText("URL submitted to VirusTotal. Response: " + [Environment]::NewLine + ($response | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
        $buttonRetrieveReport.Enabled = $true
        $script:includevtreport = $true
    } elseif ($queryVirusTotal -eq 'Cancel') {
        return
    }

    # Pulsedive API
    $queryPulsedive = [System.Windows.Forms.MessageBox]::Show("Do you want to submit URL to Pulsedive API?", "Pulsedive API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($queryPulsedive -eq 'Yes') {
        $response = Invoke-RestMethod -Uri 'https://pulsedive.com/api/analyze.php' -Method POST -Body "value=$TextBoxUrlt&probe=1&pretty=1&key=redacted" -ContentType 'application/x-www-form-urlencoded'
        $script:pulsediveQid = $response.qid
        $textboxResults.AppendText("URL submitted to Pulsedive. Response: " + [Environment]::NewLine + ($response | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
        $buttonRetrieveReport.Enabled = $true
        $script:includepulsereport = $true
    } elseif ($queryPulsedive -eq 'Cancel') {
        return
    }


    