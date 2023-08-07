$indicator = $TextBoxUrl.Text

    #Anomali API
    $queryThreatStream = [System.Windows.Forms.MessageBox]::Show("Do you want to query ThreatStream API?", "ThreatStream API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($queryThreatStream -eq 'Yes') {
        $anomaliresponse = Invoke-RestMethod -Uri "https://api.threatstream.com/api/v2/intelligence/?value__contains=$indicator&limit=0" -Headers @{'Authorization' = 'apikey redacted'} 
        $selectedFields = $anomaliresponse.objects | ForEach-Object {
            New-Object PSObject -Property @{
                value = $_.value
                threatscore = $_.threatscore
                confidence = $_.confidence
                threat_type = $_.threat_type
                itype = $_.itype
                org = $_.org
                country = $_.country
                asn = $_.asn
                source = $_.source
                #tags = $_.tags
            }
        }
        # Convert each object to a HTML row
$htmlRows = $selectedFields | ForEach-Object {
    "<div class='grid-item'>
        <p><strong>Value:</strong> $($_.value)</p>
        <p><strong>Threatscore:</strong> $($_.threatscore)</p>
        <p><strong>Confidence:</strong> $($_.confidence)</p>
        <p><strong>Threat Type:</strong> $($_.threat_type)</p>
        <p><strong>Org:</strong> $($_.org)</p>
        <p><strong>Country:</strong> $($_.country)</p>
        <p><strong>ASN:</strong> $($_.asn)</p>
        <p><strong>Source:</strong> $($_.source)</p>
    </div>"
}

        # Join all the rows together
        $htmlGrid = $htmlRows -join "`n"
        $threatScoreStats = $selectedFields.threatscore | Measure-Object -Sum -Average -Maximum -Minimum
        $confidenceStats = $selectedFields.confidence | Measure-Object -Sum -Average -Maximum -Minimum

        $totalEntries = $selectedFields.Count

        $anomaliReport = @"
Total Entries: $totalEntries
Max ThreatScore: $($threatScoreStats.Maximum)
Min ThreatScore: $($threatScoreStats.Minimum)
Average ThreatScore: $($threatScoreStats.Average)
Max Confidence: $($confidenceStats.Maximum)
Min Confidence: $($confidenceStats.Minimum)
Average Confidence: $($confidenceStats.Average)
"@

        $textboxResults.AppendText($anomaliReport)
        $textboxResults.AppendText(($selectedFields | ConvertTo-Json -Depth 100))
    } elseif ($queryThreatStream -eq 'Cancel') {
        return
    }

    # Pulsedive API
    $queryPulsedive = [System.Windows.Forms.MessageBox]::Show("Do you want to query Pulsedive API?", "Pulsedive API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
    if ($queryPulsedive -eq 'Yes') {
        $pulseresponse = Invoke-RestMethod -Uri "https://pulsedive.com/api/explore.php?q=ioc%3D$indicator&limit=0&pretty=1&key=redacted"
        $textboxResults.AppendText("Pulsedive Response: " + [Environment]::NewLine + ($pulseresponse | ConvertTo-Json -Depth 100))
        $pdhtmlTableRows = $pulseresponse.results | ForEach-Object {
            $pdresult = $_
            $pdcolor = switch ($pdresult.risk) {
                'high' { 'red' }
                'medium' { 'orange' }
                'low' { 'yellow' }
                'none' { 'green' }
                default { '#66ccff' }
            }
        
            $properties = $pdresult.summary.properties
        
            @"
            <tr style='color: $pdcolor'>
                <td>Indicator</td>
                <td>$($pdresult.indicator)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Risk</td>
                <td>$($pdresult.risk)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Type</td>
                <td>$($pdresult.type)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Added</td>
                <td>$($pdresult.stamp_added)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Status Code</td>
                <td>$($properties.http.'++code')</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Content-Type</td>
                <td>$($properties.http.'++content-type')</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Region</td>
                <td>$($properties.geo.region)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Country Code</td>
                <td>$($properties.geo.countrycode)</td>
            </tr>
            <tr style='color: $pdcolor'>
                <td>Country</td>
                <td>$($properties.geo.country)</td>
            </tr>
"@
        }
        
        # Join all the table rows together
        $pdhtmlTable = $pdhtmlTableRows -join "`n"
    } elseif ($queryPulsedive -eq 'Cancel') {
        return
    }

    # VirusTotal API
$queryVirusTotal = [System.Windows.Forms.MessageBox]::Show("Do you want to query VirusTotal API?", "VirusTotal API", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
if ($queryVirusTotal -eq 'Yes') {
    $headers=@{}
    $headers.Add("accept", "application/json")
    $headers.Add("x-apikey", "redacted")
    $vtresponse = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/search?query=$indicator" -Method GET -Headers $headers

    $newTableFields = $vtresponse.data | ForEach-Object {
        New-Object PSObject -Property @{
            url = $_.id
            title = $_.attributes.title
            categories = $_.attributes.categories
            last_analysis_date = $_.attributes.last_analysis_date
            last_analysis_stats = $_.attributes.last_analysis_stats
            last_final_url = $_.attributes.last_final_url
            redirection_chain = $_.attributes.redirection_chain
            targeted_brand = $_.attributes.targeted_brand
            total_votes_harmless = $_.attributes.total_votes.harmless
            total_votes_malicious = $_.attributes.total_votes.malicious
        }
    }
    
    $vtselectedFields = $vtresponse.data | ForEach-Object {
        New-Object PSObject -Property @{
            id = $_.id
            lastAnalysisDate = $_.attributes.last_analysis_date
            lastAnalysisResults = $_.attributes.last_analysis_results
            categories = $_.attributes.categories
            totalVotes = $_.attributes.total_votes
        }
    }

$textboxResults.AppendText("VirusTotal Response: " + [Environment]::NewLine + ($vtselectedFields | ConvertTo-Json -Depth 100))
$epochDatevtint = $newTableFields.last_analysis_date
        $datevt = [DateTimeOffset]::FromUnixTimeSeconds($epochDatevtint).DateTime

$newHtmlTableRows = $newTableFields | ForEach-Object {
    "<tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>URL:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.url)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Categories:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.categories | ConvertTo-Json)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>First Submission Date:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.first_submission_date)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Last Analysis Date:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$datevt</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Last Analysis Stats:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.last_analysis_stats | ConvertTo-Json)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Last Final URL:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.last_final_url)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Redirection Chain:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.redirection_chain | ConvertTo-Json)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Targeted Brand:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.targeted_brand | ConvertTo-Json)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Total Community Votes Harmless:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.total_votes_harmless)</td></tr>
    <tr><td style='color:#66ccff; text-align:right; font-weight:bold; text-decoration:underline;'>Total Community Votes Malicious:</td><td style='color:#66ccff; text-align:center; word-wrap:break-word; max-width:30vw;'>$($_.total_votes_malicious)</td></tr>"
}

$newHtmlTable = "<table style='width:100%;'>" + ($newHtmlTableRows -join "`n") + "</table>"

# Convert last_analysis_results to HTML table rows with conditional formatting
$htmlTableRows = $vtselectedFields | ForEach-Object {
    $_.lastAnalysisResults.PSObject.Properties | ForEach-Object {
        $scannerResult = $_.Value
        $color = switch ($scannerResult.category) {
            'harmless' { 'green' }
            'undetected' { 'grey' }
            'suspicious' { 'orange' }
            'malicious' { 'red' }
            default { 'gray' }
        }
        "<tr style='color: $color'>
            <td>$($_.Name)</td><br>
            <td>$($scannerResult.method)</td><br>
            <td>$($scannerResult.category)</td><br>
            <td>$($scannerResult.result)</td><br>
        </tr>"
    }
}

# Join all the table rows together
$htmlTable = $htmlTableRows -join "`n"
} elseif ($queryVirusTotal -eq 'Cancel') {
        return
    }
    
    # Convert the results to HTML and save to a file
    $htmlContent = @"
    <html>
    <head>
    <style>
    /* Your CSS styles go here */
    body {
        font-family: Arial, sans-serif;
        font-size: 16px;
        background-color: #181818;
        color: #c0c0c0;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 0; /* Remove default padding */
        margin: 0; /* Remove default margin */
    }
    pre {
        color: #66ccff;
        font-size: 18px;
        white-space: pre-wrap; /* Wrap text onto the next line */
        word-wrap: break-word; /* Break long words onto the next line */
        overflow-wrap: break-word; /* Same as 'word-wrap', but has better support */
        max-width: 50vw; /* Limit the width to 90% of the viewport width */
    }
    
    h1 {
        color: #ffffff;
        font-size: 50px;
    }
    .grid-container {
        display: grid;
        grid-template-columns: repeat(6, minmax(200px, 1fr));
        gap: 10px;
        padding: 10px;
        margin: 0 auto;   /* Center the grid if it's less than screen width */
    }
    .grid-item {
		word-wrap: break-word;
        padding: 10px;
        border-radius: 5px;
		color: #66ccff;
    }
    table {
        border-collapse: separate;
        border-spacing: 10px;   /* add space between cells */
        max-width: 30%;
    }
    </style>
    </head>
    <body>
        <h1>Anomali Report</h1>
        <pre>$anomaliReport</pre>
        <div class='grid-container'>
        $htmlGrid
        </div>
        <h1>VirusTotal Analysis Results</h1>
        <table>
    $newHtmlTable
    </table>
    <table>
    <tr>
        <th>Scanner</th><br>
        <th>Method</th><br>
        <th>Category</th><br>
        <th>Result</th><br>
    </tr>
    $htmlTable
    </table>
    <h1>Pulsedive Report</h1>
    <table>
    $pdhtmlTable
    </table>
    </body>
    </html>
"@
    
    # Check if TextBoxUrlt starts with http:// or https://
if ($indicator -notmatch '^(http|https)://') {
    $indicator = "http://$indicator"
}

# Create a new Uri object
try {
    $uri = New-Object System.Uri($indicator)
    # Get the base domain
    $baseDomainint = $uri.Host
} catch {
    # Use the TextBoxUrlt itself if the URI cannot be determined
    $baseDomainint = $indicator
}

    # Output the base domain
    Write-Output $baseDomainint
    New-Item -ItemType Directory -Path ".\Logs\Reports\Indicators\$baseDomainint" -Force | Out-Null
    $htmlContent | Out-File -FilePath .\Logs\Reports\Indicators\$baseDomainint\intelreport.html

    # Open the HTML file in the default web browser
    Start-Process -FilePath .\Logs\Reports\Indicators\$baseDomainint\intelreport.html