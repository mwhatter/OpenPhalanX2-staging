$TextBoxUrlt = $TextBoxUrl.Text
    if ($script:includeanomalireport -eq $true) {
    # Anomali API
    $ResponseAnomali = Invoke-RestMethod -Uri "https://api.threatstream.com/api/v1/submit/search/?q=$script:note" -Headers @{"Authorization" = "apikey redacted"} -Method Get
    $anomaliObjects = $ResponseAnomali.objects | Select-Object confidence, verdict, url, file, date_added, notes, sandbox_vendor, status

    # Sort Anomali objects by confidence in descending order
    $anomaliObjects = $anomaliObjects | Sort-Object date_added -Descending

    # Format Anomali objects as a CSS Grid
    $anomaliGrid = $anomaliObjects | ForEach-Object {
        if ($_.status -eq 'processing') {
            $_.verdict = 'processing'
        }
    $verdictClass = "verdict-$($_.verdict.ToLower())"
    @"
    <div class="grid-item">
        <table>
            <tr><th>Verdict</th><td class="$verdictClass">$($_.verdict)</td></tr>
            <tr><th>URL</th><td>$($_.url)</td></tr>
            <tr><th>Date Added</th><td>$($_.date_added)</td></tr>
            <tr><th>Notes</th><td>$($_.notes)</td></tr>
            <tr><th>Sandbox Vendor</th><td>$($_.sandbox_vendor)</td></tr>
            <tr><th>Status</th><td>$($_.status)</td></tr>
        </table>
    </div>
"@
}
    }
if ($script:includepulsereport -eq $true) {
    # Pulsedive API
    $ResponsePulsedive = Invoke-RestMethod -Uri "https://pulsedive.com/api/analyze.php?qid=$script:pulsediveQid&pretty=1&key=redacted" -Method GET 
    $pulsediveData = $ResponsePulsedive.data | Select-Object indicator, type, risk, risk_recommended, manualrisk, stamp_added, stamp_retired, recent, submissions, umbrella_rank, umbrella_domain, riskfactors, redirects, threats, feeds, comments, attributes
    # Extract and join riskfactors descriptions
    $riskFactors = ($pulsediveData.riskfactors | ForEach-Object { $_.description }) -join ', '
    # Define a hashtable to map risk to CSS classes
$riskClasses = @{
    "unknown" = "risk-unknown"
    "none" = "risk-none"
    "low" = "risk-low"
    "medium" = "risk-medium"
    "high" = "risk-high"
    "critical" = "risk-critical"
}
    # Format Pulsedive data as a table
    if ($pulsediveData.risk) {
$pulsediveTable = @"
<table>
    <tr>
        <th>Indicator</th>
        <td>$($pulsediveData.indicator)</td>
    </tr>
    <tr>
        <th>Type</th>
        <td>$($pulsediveData.type)</td>
    </tr>
    <tr>
        <th>Risk</th>
        <td class="$($riskClasses[$pulsediveData.risk])">$($pulsediveData.risk)</td>
    </tr>
    <tr>
        <th>Recommended Risk</th>
        <td class="$($riskClasses[$pulsediveData.risk_recommended])">$($pulsediveData.risk_recommended)</td>
    </tr>
    <tr>
        <th>Added Date</th>
        <td>$($pulsediveData.stamp_added)</td>
    </tr>
    <tr>
        <th>Recent</th>
        <td>$($pulsediveData.recent)</td>
    </tr>
    <tr>
        <th>Submissions</th>
        <td>$($pulsediveData.submissions)</td>
    </tr>
    <tr>
        <th>Umbrella Rank</th>
        <td>$($pulsediveData.umbrella_rank)</td>
    </tr>
    <tr>
        <th>Umbrella Domain</th>
        <td>$($pulsediveData.umbrella_domain)</td>
    </tr>
    <tr>
        <th>Risk Factors</th>
        <td>$riskfactors</td>
    </tr>
</table>
"@

# Check if redirects.from and redirects.to are arrays and join the elements
# Extract and join redirects
if ($pulsediveData.redirects) {
    $pulsediveRedirectsFrom = ($pulsediveData.redirects.from | ForEach-Object { $_.indicator }) -join ', '
    $pulsediveRedirectsTo = ($pulsediveData.redirects.to | ForEach-Object { $_.indicator }) -join ', '
}

$pulsediveTable += @"
<table>
    <tr>
        <th>Redirects From</th>
        <td>$pulsediveRedirectsFrom</td>
    </tr>
    <tr>
        <th>Redirects To</th>
        <td>$pulsediveRedirectsTo</td>
    </tr>
</table>
"@


# Include attributes in a separate table if they exist
if ($pulsediveData.attributes) {
$pulsediveAttributes = $pulsediveData.attributes.PSObject.Properties | ForEach-Object {
    @"
    <tr>
        <th>$($_.Name)</th>
        <td>$($_.Value -join ', ')</td>
    </tr>
"@
}

$pulsediveTable += @"
<table>
    <tr>
        <th>Attributes</th>
    </tr>
    $pulsediveAttributes
</table>
"@
}
}
}
if ($script:includevtreport -eq $true) {
    # VirusTotal API
    $headers=@{}
    $headers.Add("accept", "application/json")
    $headers.Add("x-apikey", "redacted")
    $ResponseVirusTotal = Invoke-RestMethod -Uri "https://www.virustotal.com/api/v3/analyses/$script:virusTotalId" -Method GET -Headers $headers 


    # Extracting VirusTotal metadata and creating an HTML table
    if ($ResponseVirusTotal.meta.url_info) {
$vtMeta = $ResponseVirusTotal.meta.url_info | ForEach-Object {
    @"
        <table>
            <tr>
                <th>URL</th>
                <td>$($_.url)</td>
            </tr>
            <tr>
                <th>ID</th>
                <td>$($_.id)</td>
            </tr>
        </table>
"@
    }
}
    # Selecting necessary attributes from VirusTotal data
    $dataAttributes = $ResponseVirusTotal.data.attributes | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | Where-Object {$_.Name -ne "results"} | Select-Object -ExpandProperty Name | out-null
    
    # Creating a table with selected attributes
    $vtData = $ResponseVirusTotal.data.attributes | Select-Object $dataAttributes | ForEach-Object {
        $epochDate = $_.date
        $date = [DateTimeOffset]::FromUnixTimeSeconds($epochDate).DateTime
    
        @"
            <table>
                <tr>
                    <th>Date</th>
                    <td>$date</td>
                </tr>
                <tr>
                    <th>Harmless</th>
                    <td>$($_.stats.harmless)</td>
                </tr>
                <tr>
                    <th>Malicious</th>
                    <td>$($_.stats.malicious)</td>
                </tr>
                <tr>
                    <th>Suspicious</th>
                    <td>$($_.stats.suspicious)</td>
                </tr>
                <tr>
                    <th>Undetected</th>
                    <td>$($_.stats.undetected)</td>
                </tr>
                <tr>
                    <th>Timeout</th>
                    <td>$($_.stats.timeout)</td>
                </tr>
                <tr>
                    <th>Status</th>
                    <td>$($_.status)</td>
                </tr>
            </table>
"@
}


# Define a hashtable to map results to CSS classes
$resultClasses = @{
    "Clean" = "result-clean"
    "Unrated" = "result-unrated"
    "Malware" = "result-malware"
    "Phishing" = "result-phishing"
    "Malicious" = "result-malicious"
    "Suspicious" = "result-suspicious"
    "Spam" = "result-spam"
}

# Define a hashtable to map categories to CSS classes
$categoryClasses = @{
    "confirmed-timeout" = "category-confirmed-timeout"
    "failure" = "category-failure"
    "harmless" = "category-harmless"
    "undetected" = "category-undetected"
    "suspicious" = "category-suspicious"
    "malicious" = "category-malicious"
    "type-unsupported" = "category-type-unsupported"
}

# Format VirusTotal scanner results as a table
$vtResults = $ResponseVirusTotal.data.attributes.results.PSObject.Properties | ForEach-Object {
    $scannerName = $_.Name
    $result = $_.Value
    if ($null -ne $result.result -and $null -ne $result.category) {
        $resultClass = $resultClasses[$result.result] 
        $categoryClass = $categoryClasses[$result.category] 
    }
    @"
    <tr>
        <td>$scannerName</td>
        <td>$($result.method)</td>
        <td class="$categoryClass">$($result.category)</td>
        <td class="$resultClass">$($result.result)</td>
    </tr>
"@
} 

# Join the array of strings into a single string
$vtResults = $vtResults -join ""

# Wrap the VirusTotal results in a table
$vtResults = @"
<table>
    <tr>
        <th>Scanner Name</th>
        <th>Method</th>
        <th>Category</th>
        <th>Result</th>
    </tr>
    $vtResults
</table>
"@
}
    
    # Convert the results to HTML and save to a file
    $htmlContent = @"
    <html>
    <head>
    <style>
    body {
        font-family: Arial, sans-serif;
        font-size: 12px;     /* reduced from 16px */
        background-color: #181818;
        color: #c0c0c0;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 0;
        margin: 0;
    }
    pre {
        color: #66ccff;
        font-size: 14px;     /* reduced from 18px */
        white-space: pre-wrap;
        word-wrap: break-word;
        overflow-wrap: break-word;
        max-width: 50vw;
    }
    h1 {
        color: #ffffff;
        font-size: 40px;     /* reduced from 50px */
    }
    .grid-container {
		word-wrap: break-word;
		width: 50vw;
        max-width: 100%;  /* Ensures the grid doesn't exceed the screen width */
        margin: 0 auto;   /* Center the grid if it's less than screen width */
        box-sizing: border-box; /* Include padding and border in element's total width and height */
    }
    .vtgrid-container {
        max-width: 100%;  /* Ensures the grid doesn't exceed the screen width */
        margin: 0 auto;   /* Center the grid if it's less than screen width */
        box-sizing: border-box; /* Include padding and border in element's total width and height */
    }
    
    .grid-item {
        box-sizing: border-box; /* Include padding and border in element's total width and height */
        padding: 10px;  /* Inner spacing. Adjust as needed */
        width: 100%; /* Takes the full width of the grid column */
    }
    
    table {
        border-collapse: separate;
        border-spacing: 10px;   /* add space between cells */
        width: 100%;
    }
    th, td {
        border: 1px solid #ddd;
        padding: 8px;
        color: #66ccff;        /* color of text in table cells */
    }
    th {
        padding-top: 12px;
        padding-bottom: 12px;
        text-align: left;
        background-color: #242424;
        color: white;           /* color of text in table headers */
    }
    td {
		word-wrap: anywhere;
	}
    .result-clean { color: green; }
    .result-suspicious, .result-spam { color: yellow; }
    .result-unrated, .category-undetected, .category-type-unsupported { color: gray; }
    .result-malware, .result-phishing, .result-malicious { color: red; }
    .category-confirmed-timeout, .category-failure, .category-suspicious { color: yellow; }
    .category-harmless { color: green; }
    .category-malicious { color: red; }
    .verdict-benign { color: green; }
    .verdict-suspicious { color: yellow; }
    .verdict-malicious { color: red; }
    .risk-unknown, .risk-recommended-unknown { background-color: gray; }
    .risk-none, .risk-recommended-none { background-color: green; }
    .risk-low, .risk-recommended-low { background-color: yellow; }
    .risk-medium, .risk-recommended-medium { background-color: orange; }
    .risk-high, .risk-recommended-high { background-color: red; }
    .risk-critical, .risk-recommended-critical { background-color: brightred; }
    </style>
    </head>
    <body>
    <h1>Anomali Report</h1>
    <div class="grid-container">
    $anomaliGrid
    </div>
    <h1>VirusTotal Report</h1>
    <h2>Meta</h2>
    <pre>$vtMeta</pre>
    <h2>Data</h2>
    <pre>$vtData</pre>
    <h2>Results</h2>
    <div class="vtgrid-container">
    $vtResults
    </div>
    <h1>Pulsedive Report</h1>
    <pre>$pulsediveTable</pre>
    </body>
    </html>
"@
    # Check if TextBoxUrlt starts with http:// or https://
if ($TextBoxUrlt -notmatch '^(http|https)://') {
    $TextBoxUrlt = "http://$TextBoxUrlt"
}

# Create a new Uri object
try {
    $uri = New-Object System.Uri($TextBoxUrlt)
    # Get the base domain
    $baseDomain = $uri.Host
} catch {
    # Use the TextBoxUrlt itself if the URI cannot be determined
    $baseDomain = $TextBoxUrlt
}

# Output the base domain
Write-Output $baseDomain
New-Item -ItemType Directory -Path ".\Logs\Reports\Indicators\$baseDomain" -Force | Out-Null
$htmlContent | Out-File -FilePath .\Logs\Reports\Indicators\$baseDomain\urlscanreport.html

# Open the HTML file in the default web browser
Start-Process -FilePath .\Logs\Reports\Indicators\$baseDomain\urlscanreport.html
