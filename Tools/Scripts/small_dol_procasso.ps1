$computerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

New-Item -ItemType Directory -Path ".\Logs\Reports\$computerName\ProcessAssociations" -Force | Out-Null
$colprocassStart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Mapping Process Associations from $computerName at $colprocassStart" -ForegroundColor Cyan 
Log_Message -logfile $logfile -Message "Mapped Process Associations from $computerName"
$textboxResults.AppendText("Mapping Process Associations from $computerName at $colprocassStart `r`n")

$scriptBlock = {
    function Get-ProcessAssociations {
        param(
            [Parameter(Mandatory=$false)]
            [int]$parentId = 0,
            [Parameter(Mandatory=$false)]
            [int]$depth = 0,
            [Parameter(Mandatory=$false)]
            [array]$allProcesses,
            [Parameter(Mandatory=$false)]
            [array]$allWin32Processes,
            [Parameter(Mandatory=$false)]
            [array]$allServices,
            [Parameter(Mandatory=$false)]
            [array]$allNetTCPConnections,
            [Parameter(Mandatory=$false)]
            [bool]$isChild = $false
        )
        
        if ($depth -gt 10) { return }

        if ($depth -eq 0) {
            $allProcesses = Get-Process | Where-Object { $_.Id -ne 0 }
            $allWin32Processes = Get-CimInstance -ClassName Win32_Process | Where-Object { $_.ProcessId -ne 0 } | Sort-Object CreationDate
            $allServices = Get-CimInstance -ClassName Win32_Service
            $allNetTCPConnections = Get-NetTCPConnection
        }
        
        $processes = if ($parentId -eq 0) {
            $allProcesses
        } else {
            $allWin32Processes | Where-Object { $_.ParentProcessId -eq $parentId } | ForEach-Object { $processId = $_.ProcessId; $allProcesses | Where-Object { $_.Id -eq $processId } }
        }
        
        $combinedData = @()
        
        $processes | ForEach-Object {
            $process = $_

            $processInfo = $allWin32Processes | Where-Object { $_.ProcessId -eq $process.Id } | Select-Object -Property CreationDate,CSName,ProcessName,CommandLine,Path,ParentProcessId,ProcessId
            
            $childProcesses = Get-ProcessAssociations -parentId $process.Id -depth ($depth + 1) -allProcesses $allProcesses -allWin32Processes $allWin32Processes -allServices $allServices -allNetTCPConnections $allNetTCPConnections -isChild $true
            
            $combinedObj = New-Object System.Collections.Specialized.OrderedDictionary
                $combinedObj["ProcessInfo"] = $processInfo
                $combinedObj["ChildProcesses"] = $childProcesses

            if (!$isChild) {
                $associatedServices = $allServices | Where-Object { $_.ProcessId -eq $process.Id } | select-object -Property Caption,Description,Name,StartMode,PathName,ProcessId,ServiceType,StartName,State
                $associatedModules = $process.Modules | Select-Object @{Name = "ProcessId"; Expression = {$process.Id}}, ModuleName, FileName
                $associatedThreads = $process.Threads | Select-Object @{Name = "ProcessId"; Expression = {$process.Id}}, Id, TotalProcessorTime, ThreadState, WaitReason
                $NetTCPConnections = $allNetTCPConnections | Where-Object { $_.OwningProcess -eq $process.Id } | select-object -Property CreationTime,State,LocalAddress,LocalPort,OwningProcess,RemoteAddress,RemotePort

                $combinedObj["NetTCPConnections"] = $NetTCPConnections
                $combinedObj["Modules"] = $associatedModules
                $combinedObj["Services"] = $associatedServices
                $combinedObj["Threads"] = $associatedThreads
            }

            $combinedData += $combinedObj
        }

        return $combinedData
    }

    return Get-ProcessAssociations -parentId $args[0] -depth $args[1]
}

$processData = Invoke-Command -ComputerName $computerName -ScriptBlock $scriptBlock -ArgumentList 0, 0

$processData | ForEach-Object {
    $process = $_
    if ($process) {
        $processInfo = $process.ProcessInfo
        $processId = $processInfo.ProcessId
        $processName = $processInfo.ProcessName
        $creationDate = $processInfo.CreationDate

        if ($processName -ne $null -and $processId -ne $null -and $creationDate -ne $null) {
            $html += "<div class='process-line' timestamp='$creationDate'>"
            $html += "<a onclick=`"showProcessDetails('$processId'); return false;`">$processName ($processId) - Created on: $creationDate</a>"
            $html += "<div id='$processId' class='process-info' style='display: none;'>"
        }

        if ($process.ProcessInfo) {
            $html += "<pre>" + ($process.ProcessInfo | ConvertTo-Json -Depth 100) + "</pre>"
        }

        if ($processName -ne $null -and $processId -ne $null -and $creationDate -ne $null) {
            $html += "</div></div>"
        }
    }
}

function Build-HTMLProcessList {
    param(
        [Parameter(Mandatory=$true)]
        [array]$processData
    )

    $html = ''

    foreach ($process in ($processData | Sort-Object -Property @{Expression = { $_.ProcessInfo.ProcessId }; Ascending = $true})) {
        if ($null -eq $process) {
            continue
        }
        $processName = $process.ProcessInfo.ProcessName
        $processId = $process.ProcessInfo.ProcessId
        $creationDate = $process.ProcessInfo.CreationDate

        if ($null -ne $processName -and $null -ne $processId -and $null -ne $creationDate) {
            $html += "<div class='process-line' timestamp='$creationDate'>"
            $html += "<a onclick=`"showProcessDetails('$processId'); return false;`">$processName ($processId) - Created on: $creationDate</a><br>"
            $html += "<div id='$processId' class='process-info' style='display: none;'>"
        }

        $processInfo = $process.ProcessInfo | ConvertTo-Json -Depth 100

        # Checking if $processInfo is an empty object
        if ($processInfo -ne "{}") {
            $html += $processInfo
        }

        if ($null -ne $processName -and $null -ne $processId -and $null -ne $creationDate) {

            if ($process.NetTCPConnections) {
                $html += "<br><a onclick=`"showNetworkConnections('$processId'); return false;`">Show Network Connections for Process ID: $processId</a><br>"
                $html += "<div id='net-$processId' class='network-info' style='display: none;'>"
                $html += ($process.NetTCPConnections | ConvertTo-Json -Depth 100)
                $html += "</div>"
            }
    
            if ($process.Modules) {
                $html += "<a onclick=`"showModules('$processId'); return false;`">Show Modules for Process ID: $processId</a><br>"
                $html += "<div id='mod-$processId' class='module-info' style='display: none;'>"
                $html += ($process.Modules | ConvertTo-Json -Depth 100)
                $html += "</div>"
            }
    
            if ($process.Services) {
                $html += "<a onclick=`"showServices('$processId'); return false;`">Show Services for Process ID: $processId</a><br>"
                $html += "<div id='ser-$processId' class='service-info' style='display: none;'>"
                $html += ($process.Services | ConvertTo-Json -Depth 100)
                $html += "</div>"
            }

            if ($process.ChildProcesses) {
                $html += "<a onclick=`"showChildProcesses('$processId'); return false;`">Show Child Processes for Process ID: $processId</a><br>"
                $html += "<div id='child-$processId' class='child-process-info' style='display: none;'>"
                $html += ($process.ChildProcesses | ConvertTo-Json -Depth 100)
                $html += "</div>"
            }
    
            $html += "</div></div>"
    }
}

    return $html
}


$processListHtml = Build-HTMLProcessList -processData $processData
$outFilePath = ".\Logs\Reports\$computerName\ProcessAssociations\ProcessList.html"
$outFileContent = @"
<!DOCTYPE html>
<html>
<head>
<title>Process List Explorer</title>
<style>
    /* Include your CSS styles here */
    body {
        font-family: Arial, sans-serif;
        font-size: 16px;
        background-color: #181818;
        color: #c0c0c0;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    a {
        color: #66ccff;
    }
    h1 {
        text-align: center;
    }
    .process-info {
        white-space: pre-wrap; /* CSS3 */
        word-wrap: break-word; /* Internet Explorer 5.5+ */
        font-size: 14px;
        color: #4dc5b5;
        background-color: #333333;
        border: solid 1px #666666;
        margin: 5px;
        padding: 10px;
        overflow-wrap: break-word; /* Adds support for Firefox and other browsers */
        max-width: 50vw; /* Limit width to 50% of the viewport width */
        word-break: break-all; /* Breaks the words at the limit of the container width */
        overflow: auto; /* Adds a scrollbar if the content exceeds the max-width */
    }
    .child-process-info {
        white-space: pre-wrap; /* CSS3 */
        word-wrap: break-word; /* Internet Explorer 5.5+ */
        font-size: 14px;
        color: #4dc5b5;
        background-color: #333333;
        margin: 5px;
        padding: 10px;
        overflow-wrap: break-word; /* Adds support for Firefox and other browsers */
        max-width: 50vw; /* Limit width to 50% of the viewport width */
        word-break: break-all; /* Breaks the words at the limit of the container width */
        overflow: auto; /* Adds a scrollbar if the content exceeds the max-width */
    }
    module-info {
        white-space: pre-wrap; /* CSS3 */
        word-wrap: break-word; /* Internet Explorer 5.5+ */
        font-size: 14px;
        color: #4dc5b5;
        background-color: #333333;
        margin: 5px;
        padding: 10px;
        overflow-wrap: break-word; /* Adds support for Firefox and other browsers */
        max-width: 50vw; /* Limit width to 50% of the viewport width */
        word-break: break-all; /* Breaks the words at the limit of the container width */
        overflow: auto; /* Adds a scrollbar if the content exceeds the max-width */
    }
    .service-info {
        white-space: pre-wrap; /* CSS3 */
        word-wrap: break-word; /* Internet Explorer 5.5+ */
        font-size: 14px;
        color: #4dc5b5;
        background-color: #333333;
        margin: 5px;
        padding: 10px;
        overflow-wrap: break-word; /* Adds support for Firefox and other browsers */
        max-width: 50vw; /* Limit width to 50% of the viewport width */
        word-break: break-all; /* Breaks the words at the limit of the container width */
        overflow: auto; /* Adds a scrollbar if the content exceeds the max-width */
    }
    .netowrk-info {
        white-space: pre-wrap; /* CSS3 */
        word-wrap: break-word; /* Internet Explorer 5.5+ */
        font-size: 14px;
        color: #4dc5b5;
        background-color: #333333;
        margin: 5px;
        padding: 10px;
        overflow-wrap: break-word; /* Adds support for Firefox and other browsers */
        max-width: 50vw; /* Limit width to 50% of the viewport width */
        word-break: break-all; /* Breaks the words at the limit of the container width */
        overflow: auto; /* Adds a scrollbar if the content exceeds the max-width */
    }
    .process-line {
        display: block;
        text-align: center;
    }
    .controls {
        position: fixed;
        top: 0;
        width: 100%;
        background-color: #181818;
        padding: 10px;
        z-index: 100;
        display: flex;
        justify-content: space-around;
    }
    .controls > div {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }
    .controls > div > div {
        display: flex;
        align-items: center;
        justify-content: flex-end;
        gap: 10px;
    }
    .controls > div > label {
        margin-bottom: 10px;
        font-size: 18px;
        font-weight: bold;
    }
    .content {
        margin-top: 100px; /* Adjust this value based on the height of your controls */
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
        padding: 15px;
        border-radius: 4px;
    }
    .top-button:hover {
        background-color: #444;
    }
</style>
<script>
function addSoftHyphens(str, interval) {
    var result = '';
    while (str.length > 0) {
        result += str.substring(0, interval) + '&shy;';
        str = str.substring(interval);
    }
    return result;
}
function showChildProcesses(processId) {
    var infoElem = document.getElementById('child-' + processId);
    if (infoElem.style.display === "none") {
        infoElem.style.display = "block";
    } else {
        infoElem.style.display = "none";
    }
}
function showNetworkConnections(processId) {
    var infoElem = document.getElementById('net-' + processId);
    if (infoElem.style.display === "none") {
        infoElem.style.display = "block";
    } else {
        infoElem.style.display = "none";
    }
}
function showModules(processId) {
    var infoElem = document.getElementById('mod-' + processId);
    if (infoElem.style.display === "none") {
        infoElem.style.display = "block";
    } else {
        infoElem.style.display = "none";
    }
}
function showServices(processId) {
    var infoElem = document.getElementById('ser-' + processId);
    if (infoElem.style.display === "none") {
        infoElem.style.display = "block";
    } else {
        infoElem.style.display = "none";
    }
}
        function showProcessDetails(processId) {
            var infoElem = document.getElementById(processId.toString());
            if (infoElem.style.display === "none") {
                infoElem.style.display = "block";
            } else {
                infoElem.style.display = "none";
            }
        }
        function openAllChildProcesses() {
            var coll = document.getElementsByClassName('child-process-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        
        function closeAllChildProcesses() {
            var coll = document.getElementsByClassName('child-process-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'none';
            }
        }

        function openAllProcesses() {
            var coll = document.getElementsByClassName('process-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        function closeAllProcesses() {
            var coll = document.getElementsByClassName('process-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'none';
            }
        }
        function filter() {
            var timeFrom = document.getElementById("timeFrom").value ? new Date(document.getElementById("timeFrom").value).getTime() : null;
            var timeTo = document.getElementById("timeTo").value ? new Date(document.getElementById("timeTo").value).getTime() : null;
            var keyword = document.getElementById("keyword").value;
        
            var coll = document.getElementsByClassName('process-line');
            for (var i = 0; i < coll.length; i++) {
                var timestamp = new Date(coll[i].getAttribute('timestamp')).getTime();
                var processName = coll[i].innerText;
        
                // Hide element initially
                coll[i].style.display = 'none';
        
                // Show if within the time window (or no time window set)
                if ((!timeFrom || timestamp >= timeFrom) && (!timeTo || timestamp <= timeTo)) {
                    // And show if keyword is empty or process name contains keyword
                    if (!keyword || processName.indexOf(keyword) !== -1) {
                        coll[i].style.display = 'block';
                    }
                }
            }
        }
        function reset() {
            document.getElementById("timeFrom").value = "";
            document.getElementById("timeTo").value = "";
            document.getElementById("keyword").value = "";
            var coll = document.getElementsByClassName('process-line');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        function openAllNetworks() {
            var coll = document.getElementsByClassName('network-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        function closeAllNetworks() {
            var coll = document.getElementsByClassName('network-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'none';
            }
        }
        function openAllModules() {
            var coll = document.getElementsByClassName('module-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        function closeAllModules() {
            var coll = document.getElementsByClassName('module-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'none';
            }
        }
        function openAllServices() {
            var coll = document.getElementsByClassName('service-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'block';
            }
        }
        function closeAllServices() {
            var coll = document.getElementsByClassName('service-info');
            for (var i = 0; i < coll.length; i++) {
                coll[i].style.display = 'none';
            }
        }
        function openAll() {
        openAllProcesses();
        openAllNetworks();
        openAllModules();
        openAllServices();
        openAllChildProcesses();
        }

        function closeAll() {
        closeAllProcesses();
        closeAllNetworks();
        closeAllModules();
        closeAllServices();
        closeAllChildProcesses();
        }

        window.onscroll = function() {scrollFunction()};

        function scrollFunction() {
          if (document.body.scrollTop > 20 || document.documentElement.scrollTop > 20) {
            document.getElementById("topBtn").style.display = "block";
          } else {
            document.getElementById("topBtn").style.display = "none";
          }
        }
        
        function topFunction() {
          document.body.scrollTop = 0;
          document.documentElement.scrollTop = 0;
        }
    </script>
</head>
<body>
<div class="controls">
    <h1>Process List</h1>
    <div>
        <label>Time Picker</label>
        <div>
            <label>From:</label>
            <input type="datetime-local" id="timeFrom">
        </div>
        <div>
            <label>To:</label>
            <input type="datetime-local" id="timeTo">
        </div>
    </div>
    <div>
        <label>Keyword</label>
        <input type="text" id="keyword">
        <div style="display: flex;">
            <button onclick="filter();">Apply Filter</button>
            <button onclick="reset();">Reset</button>
        </div>
    </div>
    <div>
        <label>Process Details</label>
        <button onclick="openAllProcesses();">Expand All</button>
        <button onclick="closeAllProcesses();">Collapse All</button>
    </div>
    <div>
        <label>Network Details</label>
        <button onclick="openAllNetworks();">Expand All</button>
        <button onclick="closeAllNetworks();">Collapse All</button>
    </div>
    <div>
        <label>Module Details</label>
        <button onclick="openAllModules();">Expand All</button>
        <button onclick="closeAllModules();">Collapse All</button>
    </div>
    <div>
        <label>Service Details</label>
        <button onclick="openAllServices();">Expand All</button>
        <button onclick="closeAllServices();">Collapse All</button>
    </div>
    <div>
        <label>Child Process Details</label>
        <button onclick="openAllChildProcesses();">Expand All</button>
        <button onclick="closeAllChildProcesses();">Collapse All</button>
    </div>
    <div>
        <label>All Details</label>
        <button onclick="openAll();">Expand All</button>
        <button onclick="closeAll();">Collapse All</button>
    </div>
</div>
<div class="content">
$processListHtml
<button onclick="topFunction()" id="topBtn" class="top-button">Return to Top</button>
</div>
</body>
</html>
"@

Set-Content -Path $outFilePath -Value $outFileContent


Invoke-Item $outFilePath


$colprocassEnd = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "Completed Mapping of Process Associations from $computerName at $colprocassEnd" -ForegroundColor Green
Log_Message -logfile $logfile -Message "Completed Mapping of Process Associations from $computerName at $colprocassEnd"  
$textboxResults.AppendText("Completed Mapping of Process Associations from $computerName at $colprocassEnd `r`n")