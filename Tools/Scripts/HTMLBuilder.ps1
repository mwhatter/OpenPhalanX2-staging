#.\HTMLBuilder.ps1 -JsonData '{"key": "value", "anotherKey": "anotherValue"}' -OutputFilePath "path_to_output.html"

param (
    [Parameter(Mandatory=$true)]
    [string]$JsonData,
    [Parameter(Mandatory=$true)]
    [string]$OutputFilePath
)

function Generate_HTML {
    param (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$DataObject
    )

    # Convert the data object into a JSON formatted string for display
    $jsonDataDisplay = $DataObject | ConvertTo-Json -Depth 10

    # Generate the HTML content
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>API Data Viewer</title>
    <style>
        /* Styles based on the provided example */
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
        .api-data-content {
            white-space: pre-wrap;
            font-size: 14px;
            color: #4dc5b5;
            background-color: #333333;
            border: solid 1px #666666;
            margin: 5px;
            padding: 10px;
            overflow-wrap: break-word;
            max-width: 80vw;
            word-break: break-all;
            overflow: auto;
        }
    </style>
</head>
<body>
<div class="api-data-content">
$jsonDataDisplay
</div>
</body>
</html>
"@

    return $htmlContent
}

# Convert the passed JSON string to an object
$jsonObject = $JsonData | ConvertFrom-Json

# Generate the HTML content
$htmlOutput = Generate_HTML -DataObject $jsonObject

# Save the HTML to the specified output file path
$htmlOutput | Out-File -Path $OutputFilePath

Write-Host "HTML file generated at $OutputFilePath"


