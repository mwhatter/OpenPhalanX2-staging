$ComputerName = if ($comboBoxComputerName.SelectedItem) {
    $comboBoxComputerName.SelectedItem.ToString()
} else {
    $comboBoxComputerName.Text
}

if (-not $ComputerName) {
    [System.Windows.Forms.MessageBox]::Show("Please enter a remote computer name.")
    return
}

function Update-ListView {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string] $path
    )
    $listView.Items.Clear()
    $currentPathTextBox.Text = $path
    $items = Get-ChildItem -Path $path -Force -ErrorAction SilentlyContinue
    foreach ($item in $items) {
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
        $listViewItem.ImageIndex = if ($item.PSIsContainer) { 0 } else { 1 }
        
        if (-not $item.PSIsContainer) {
            # For files
            $listViewItem.SubItems.Add($item.Length.ToString())  # Size
        } else {
            # For directories
            $listViewItem.SubItems.Add("")  # Size
        }
        
        $listViewItem.SubItems.Add($item.LastWriteTime.ToString())  # Last Modified
        $listViewItem.SubItems.Add($item.LastAccessTime.ToString())  # Last Access Time
        $listViewItem.SubItems.Add($item.CreationTime.ToString())  # Creation Time
        $listViewItem.SubItems.Add($item.Attributes.ToString())  # Attributes

        $listViewItem.Tag = $item.FullName
        $listView.Items.Add($listViewItem) | Out-Null
    }
}

Function Get-RemoteResources {
    param($computerName)
    try {
        $resources = @()
        Invoke-Command -ComputerName $computerName -ScriptBlock {
            $shares = Get-CimInstance -ClassName Win32_Share | ForEach-Object {
                "\\$using:computerName\$($_.Name)"
            }
            $shares
        } | ForEach-Object {
            $resources += $_
        }
        return $resources
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: Failed to retrieve resources from the remote computer.")
        return @()
    }
}

$remoteRootPath = ""
$script:currentPath = ""

$remoteRootPath = "\\$ComputerName\C$"
    $script:currentPath = $remoteRootPath

    $fileBrowserForm = New-Object System.Windows.Forms.Form
    $fileBrowserForm.Text = "Select a Remote File"
    $fileBrowserForm.ForeColor = [System.Drawing.Color]::lightseagreen   
    $fileBrowserForm.BackColor = [System.Drawing.Color]::Black
    $fileBrowserForm.Size = New-Object System.Drawing.Size(780, 590)
    $fileBrowserForm.FormBorderStyle = 'Fixed3D'
    $fileBrowserForm.StartPosition = "CenterScreen"

    $currentPathTextBox = New-Object System.Windows.Forms.TextBox
    $currentPathTextBox.Location = New-Object System.Drawing.Point(355, 510)
    $currentPathTextBox.Size = New-Object System.Drawing.Size(400, 23)
    $currentPathTextBox.ReadOnly = $true
    $fileBrowserForm.Controls.Add($currentPathTextBox)

    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.BackColor = [System.Drawing.Color]::Black
    $listView.ForeColor = [System.Drawing.Color]::lightseagreen
    $listView.Size = New-Object System.Drawing.Size(750, 500) 
    $listView.Location = New-Object System.Drawing.Point(5, 5)
    $listView.FullRowSelect = $true
    $listView.Columns.Add("Name", 150) | Out-Null
    $listView.Columns.Add("Size", 75) | Out-Null
    $listView.Columns.Add("Last Modified", 140) | Out-Null
    $listView.Columns.Add("Last Access Time", 140) | Out-Null 
    $listView.Columns.Add("Creation Time", 140) | Out-Null
    $listView.Columns.Add("Attributes", 350) | Out-Null

    $listView.Add_DoubleClick({
        $selectedItem = $listView.SelectedItems[0]
        $selectedFile = $selectedItem.Tag
        if ($selectedItem.ImageIndex -eq 0) { # Directory
            $script:currentPath = $selectedFile
            Update-ListView -path $script:currentPath
        } else { # File
            $remotePath = $selectedFile.Replace($remoteRootPath, "C:")
            $textboxRemoteFilePath.Text = $remotePath
            $fileBrowserForm.Close()
        }
    })

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Text = "Back"
    $backButton.Location = New-Object System.Drawing.Point(5, 510)
    $backButton.Size = New-Object System.Drawing.Size(100, 23)

    $backButton.Add_Click({
        if (-not [String]::Equals($script:currentPath, $remoteRootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $script:currentPath = Split-Path -Parent $script:currentPath
            Update-ListView -path $script:currentPath
        }
    })

    $driveComboBox = New-Object System.Windows.Forms.ComboBox
    $driveComboBox.Location = New-Object System.Drawing.Point(115, 510)
    $driveComboBox.Size = New-Object System.Drawing.Size(230, 23)
    $driveComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $driveComboBox.Add_SelectedIndexChanged({
    $script:currentPath = $driveComboBox.SelectedItem.ToString()
    Update-ListView -path $script:currentPath
    }) 

    $drives = Get-RemoteResources -computerName $ComputerName
    $driveComboBox.Items.Clear()
    $driveComboBox.Items.AddRange($drives)

    if ($driveComboBox.Items.Count -gt 0) {
        $driveComboBox.SelectedIndex = 0
    }

$fileBrowserForm.Controls.Add($driveComboBox)
$fileBrowserForm.Controls.Add($backButton)
$fileBrowserForm.Controls.Add($listView)
$fileBrowserForm.Add_Shown({ Update-ListView -path $remoteRootPath })
$fileBrowserForm.ShowDialog()

