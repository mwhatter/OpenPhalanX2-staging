# Function to calculate SHA256 hash
function Get-FileHashSHA256 {
    param (
        [String]$FilePath
    )
    
    $hasher = [System.Security.Cryptography.HashAlgorithm]::Create("SHA256")
    $fileStream = New-Object -TypeName System.IO.FileStream -ArgumentList $FilePath, 'Open'
    $hash = $hasher.ComputeHash($fileStream)
    $fileStream.Close()
    return ([BitConverter]::ToString($hash)).Replace("-", "")
}

function Get-File($path) {
    $file = Get-Item $path
    $fileType = New-Object -TypeName PSObject -Property @{
        IsText = $file.Extension -in @(".txt", ".csv", ".log", ".evtx", ".xlsx", ".xml", ".json", ".html", ".htm", ".md", ".ps1", ".bat", ".css", ".js")  # Add any other text file extensions you want to support
    }
    return $fileType
}

# Create an ArrayList for output
$output = New-Object System.Collections.ArrayList
$patterns = @{
    'HTTP/S' = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    'FTP' = 'ftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?'
    'SFTP' = 'sftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?'
    'SCP' = 'scp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?'
    'DATA' = 'data://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'
    'SSH' = 'ssh://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?'
    'LDAP' = 'ldap://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?'
    'RFC 1918 IP Address' = '\b(10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3})\b'
    'Non-RFC 1918 IP Address' = '\b((?!10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|127\.\d{1,3}\.\d{1,3}\.\d{1,3})\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b'
    'Email' = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
}

# Prompt user to include File Path pattern
$includeFilePath = [System.Windows.Forms.MessageBox]::Show("Do you want to include File Paths?", "Include File Paths", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($includeFilePath -eq 'Yes') {
    $patterns['File Path'] = '(file:///)?(?![sSpP]:)[a-zA-Z]:[\\\\/].+?\.[a-zA-Z0-9]{2,5}(?=[\s,;]|$)'


}    

# Rest of your code continues here
$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

Remove-Item .\Logs\Reports\$computerName\indicators.html -ErrorAction SilentlyContinue 
New-Item -ItemType Directory -Path ".\Logs\Reports\$computerName\" -Force | Out-Null
$Intelligazerstart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Intelligazer started at $Intelligazerstart" -ForegroundColor Cyan
Log_Message -logfile $logfile -Message "Intelligazer started at "
$textboxResults.AppendText("Intelligazer started at $Intelligazerstart `r`n")
$logDirs = ".\Logs\EVTX\$computerName", ".\Logs\Reports\$computerName\RapidTriage", ".\Logs\Reports\$computerName\ADRecon"


function processMatches($content, $file) {
$matchList = New-Object System.Collections.ArrayList
foreach ($type in $patterns.Keys) {
    $pattern = $patterns[$type]
    $matchResults = [regex]::Matches($content, $pattern) | ForEach-Object {$_.Value}
    foreach ($match in $matchResults) {
        $newObject = New-Object PSObject -Property @{
            'Source File' = $file
            'Data' = $match
            'Type' = $type
        }
        if ($null -ne $newObject) {
            [void]$matchList.Add($newObject)
        }

        # If the type is URL, extract the parent domain and add it as a separate indicator
        if ($type -eq 'HTTP/S' -and $match -match '(?i)(?:http[s]?://)?(?:www.)?([^/]+)') {
            $parentDomain = $matches[1]
            $domainObject = New-Object PSObject -Property @{
                'Source File' = $file
                'Data' = $parentDomain
                'Type' = 'Domain'
            }
            if ($null -ne $domainObject) {
                [void]$matchList.Add($domainObject)
            }
        }
    }
}
return $matchList
}

foreach ($dir in $logDirs) {
    $files = Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName 
    foreach ($file in $files) {
        switch -regex ($file) {
            '\.sqlite$' {
                $tableNames = Invoke-SqliteQuery -DataSource $file -Query "SELECT name FROM sqlite_master WHERE type='table';"
                foreach ($tableName in $tableNames.name) {
                    try {
                        $query = "SELECT * FROM [$tableName]"
                        $data = Invoke-SqliteQuery -DataSource $file -Query $query -ErrorAction Stop
                        $content = $data | Out-String
                        $matchess = processMatches $content $file
                        if ($matchess -ne $null) {
                            $output.AddRange(@($matchess))
                        } 
                    } catch {
                        #Add error message if desired
                    }
                }
            }
            '\.(csv|txt|json|evtx|html)$' {
                if ($file -match "\.evtx$") {
                    try {
                        $content = Get-WinEvent -Path $file -ErrorAction Stop | Format-List | Out-String
                    } catch {
                        Write-Host "No events found in $file" -ForegroundColor Magenta
                        continue
                    }
                } else {
                    $content = Get-Content $file
                }
                $matchess = processMatches $content $file
                if ($matchess -ne $null) {
                    $output.AddRange(@($matchess))
                }
            }
            '\.xlsx?$' {
                $excel = $null
                $workbook = $null
            
                try {
                    $excel = New-Object -ComObject Excel.Application
                    $excel.Visible = $false # Ensure Excel runs in the background
                    
                    $workbook = $excel.Workbooks.Open($file)
                    
                    foreach ($sheet in $workbook.Worksheets) {
                        try {
                            $range = $sheet.UsedRange
                            $content = $range.Value2 | Out-String
                            $matches = processMatches $content $file
                            if ($matches -ne $null) {
                                $output.AddRange($matches)
                            } 
                        }
                        catch {
                            Write-Error "Error processing sheet: $_"
                        }
                    }
                }
                catch {
                    Write-Error "An error occurred: $_"
                }
                finally {
                    if ($workbook) {
                        $workbook.Close($false)
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    }
            
                    if ($excel) {
                        $excel.Quit()
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    }
            
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                }
            }
            
            default {
                if ((Get-File $file).IsText) {
                    $content = Get-Content $file
                    $matchess = processMatches $content $file
                    if ($matchess -ne $null) {
                        $output.AddRange($matchess)
                    }
                }
            }
        }
    }
}     


# Process CopiedFiles for SHA256
$copiedFiles = Get-ChildItem $copiedFilesDir -Recurse -File | Select-Object -ExpandProperty FullName
foreach ($file in $copiedFiles) {
$sha256 = Get-FileHashSHA256 -FilePath $file

# Add new object to the ArrayList
[void]$output.Add((New-Object PSObject -Property @{
    'Source File' = $file
    'Data' = $sha256
    'Type' = 'SHA256'
}))
}

# Deduplicate the output
$output = $output | Sort-Object 'Data' -Unique

$Indicatordone = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$count = $output.Count
Write-Host "Indicators extracted at $Indicatordone. Total count: $count" -ForegroundColor Cyan
Log_Message -logfile $logfile -Message "Indicators extracted "
$textboxResults.AppendText("Indicators extracted at $Indicatordone. Total count: $count `r`n")
# Create an HTML file
$htmlFile = New-Item -Path ".\Logs\Reports\$computerName\indicators.html" -ItemType File -Force

# Write the start of the HTML file
$html = @"
<!DOCTYPE html>
<html>
<head>
<title>Indicator List Explorer</title>
<style>
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
#controls {
    position: sticky;
    top: 0;
    z-index: 1;
    width: calc(100% - 20px);
    padding: 10px 0;
    display: flex;
    justify-content: space-between;
    align-items: center;
    height: 60px; 
    box-sizing: border-box;
    flex-direction: row;
}
#filter-type {
    height: 310px;
    width: auto;
    font-size: 16px;
    margin-right: 10px;
    z-index: 2; /* set z-index higher than the other controls */
}
    .indicator {
        margin: 10px 0;
        text-align: left;
        width: 90%;
        color: #4dc5b5;
    }
    .indicator:hover {
        cursor: pointer;
    }
    .indicator-info {
        display: none;
        margin-top: 20px;
        text-align: center;
    }
    
    #controls h1 {
        margin: 0;
        color: #ffffff;
        font-size: 50px;
        margin-right: auto; 
    }
    #controls div {
        display: flex;
        gap: 10px;
        align-items: center;
        order: 1;  /* change order */
        align-self: center; 
        z-index: 1; /* set z-index lower than the select */
    }
    #indicators {
        margin-top: 60px;  /* add this line */
        width: 60vw; /* Set the width to 75% of the viewport width */
    }
    #controls-select {
        position: sticky;
        top: 60px;
        z-index: 1;
        width: 100%;
        display: flex;
        justify-content: flex-end;
        align-items: center;
        padding: 10px 0;
        box-sizing: border-box;
    }
    #filter-keyword {
        height: 35px;
        font-size: 20px;
    }
    #filter-type {
        order: 2;  /* change order */
        height: 325px;  /* adjust as needed */
        width: auto;
        font-size: 16px;
        margin-top: 10px;  /* add margin at the top */
        position: absolute; /* takes the element out of the normal document flow */
        right: 0; /* aligns it to the right of the #controls container */
        top: 100%; /* pushes it just below the #controls container */
        z-index: 2; /* set z-index higher than the other controls */
    }
    #filter-button {
        height: 35px;
        font-size: 20px; /* Adjust the size as you wish */
    }
    #reset-button {
        height: 35px;
        font-size: 20px; /* Adjust the size as you wish */
    }
    #scroll-button {
        height: 35px;
        font-size: 20px; /* Adjust the size as you wish */
    }
    .indicator-data {
        color: #66ccff;  /* adjust color as needed */
        font-size: 18px;  /* adjust font size as needed */
        word-wrap: break-word;  /* Wrap long words onto the next line */
    }
    .top-button {
        display: none;
        position: fixed;
        bottom: 20px;
        right: 30px;
        z-index: 99;
        border: none;
        outline: none;
        background-color: #555;
        color: white;
        cursor: pointer;
        border-radius: 4px;
    }
    .top-button:hover {
        background-color: #444;
    }
</style>
</head>
<body>
<div id='controls'>
<h1>Indicator List</h1>
<div>
    <input type='text' id='filter-keyword' placeholder='Enter keyword'>
    <button id='filter-button' onclick='filter()'>Filter List</button>
    <button id='reset-button' onclick='resetFilters()'>Reset filters</button>
</div>
</div>
<div id='controls-select'>
<select id='filter-type' onchange='filter()' multiple>
<option value=''>All (ctrl+click for multi-select)</option>
<option value='HTTP/S'>HTTP/S</option>
<option value='FTP'>FTP</option>
<option value='SFTP'>SFTP</option>
<option value='SCP'>SCP</option>
<option value='SSH'>SSH</option>
<option value='LDAP'>LDAP</option>
<option value='DATA'>DATA</option>
<option value='RFC 1918 IP Address'>RFC 1918 IP Address</option>
<option value='Non-RFC 1918 IP Address'>Non-RFC 1918 IP Address</option>
<option value='Email'>Email</option>
<option value='Domain'>Domain</option>
<option value='SHA256'>SHA256</option>
<option value='File Path'>File Path</option>
</select>
<button id='scroll-button' class='top-button' onclick='scrollToTop()'>Return to Top</button>

</select>
</div>
<div id='indicators'>
"@

Add-Content -Path $htmlFile.FullName -Value $html

$index = 0
foreach ($indicator in $output) {
$sourceFile = $indicator.'Source File'
$data = $indicator.Data
$type = $indicator.Type

# Write the indicator as a div that shows the source file and type
$indicatorHtml = @"
<div class='indicator' data-type='$type'>
<strong>Data:</strong> <span class='indicator-data'>$data</span><br>
<strong>Type:</strong> $type <br>
<strong>Source:</strong> $sourceFile
</div>
"@
Add-Content -Path $htmlFile.FullName -Value $indicatorHtml
$index++
}

# Write the end of the HTML file
$html = @"
</div>
<script>
function filter() {
    var filterKeyword = document.getElementById('filter-keyword').value.toLowerCase();
    var filterTypes = Array.from(document.getElementById('filter-type').selectedOptions).map(option => option.value.toLowerCase());
    var indicators = document.getElementsByClassName('indicator');
    for (var i = 0; i < indicators.length; i++) {
        var matchesKeyword = filterKeyword === '' || indicators[i].textContent.toLowerCase().includes(filterKeyword);
        var matchesType = filterTypes.includes(indicators[i].getAttribute('data-type').toLowerCase()) || filterTypes.includes('');
        if (matchesKeyword && matchesType) {
            indicators[i].style.display = 'block';
        } else {
            indicators[i].style.display = 'none';
        }
    }
}

function toggleInfo(index) {
    var indicatorInfo = document.getElementById('indicator-info-' + index);
    if (indicatorInfo.style.display === 'none') {
        indicatorInfo.style.display = 'block';
    } else {
        indicatorInfo.style.display = 'none';
    }
}

function resetFilters() {
    document.getElementById('filter-keyword').value = '';
    document.getElementById('filter-type').value = '';

    var indicators = document.getElementsByClassName('indicator');
    for (var i = 0; i < indicators.length; i++) {
        indicators[i].style.display = 'block';
    }
}

function expandAll() {
    var infos = document.getElementsByClassName('indicator-info');
    for (var i = 0; i < infos.length; i++) {
        infos[i].style.display = 'block';
    }
}

function collapseAll() {
    var infos = document.getElementsByClassName('indicator-info');
    for (var i = 0; i < infos.length; i++) {
        infos[i].style.display = 'none';
    }
}

function scrollToTop() {
    document.body.scrollTop = 0; // For Safari
    document.documentElement.scrollTop = 0; // For Chrome, Firefox, IE and Opera
  }
  
  window.onscroll = function() {scrollFunction()};
  
  function scrollFunction() {
    if (document.body.scrollTop > 20 || document.documentElement.scrollTop > 20) {
      document.getElementById("scroll-button").style.display = "block";
    } else {
      document.getElementById("scroll-button").style.display = "none";
    }
  }
</script>
</body>
</html>
"@

Add-Content -Path $htmlFile.FullName -Value $html

Invoke-Item .\Logs\Reports\$computerName\indicators.html

$Intelligazerdone = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Intelligazer completed at $Intelligazerdone" -ForegroundColor Cyan
Log_Message -logfile $logfile -Message "Intelligazer completed"
$textboxResults.AppendText("Intelligazer completed at $Intelligazerdone `r`n")
