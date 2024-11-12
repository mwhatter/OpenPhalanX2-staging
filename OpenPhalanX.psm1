# Import necessary assemblies and modules
Import-Module ImportExcel
Import-Module ActiveDirectory
Import-Module PSSQLite
Import-Module Microsoft.PowerShell.Utility

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.Net.Http
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Define all functions from Defending_Off_the_Land.ps1

function Log_Message {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [string]$LogFile
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "Time=[$timestamp] User=[$Username] Message=[$Message]"
}

function Get-AllComputers {
    $domains = (Get-ADForest).Domains
    $domaincomputers = $domains | ForEach-Object -Parallel {
        $domainController = (Get-ADDomainController -DomainName $_ -Discover -Service PrimaryDC).HostName
        if ($domainController -is [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
            $domainController = $domainController[0]
        }
        Get-ADComputer -Filter {Enabled -eq $true} -Server $domainController | Select-Object -ExpandProperty Name
    } | Sort-Object
    return $domaincomputers
}

function Show-DefendingOffTheLandGUI {
    $logstart = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $CWD = (Get-Location).Path
    New-Item -ItemType Directory -Path "$CWD\Logs\Audit" -Force | Out-Null
    $LogFile = "$CWD\Logs\Audit\$logstart.DOL.auditlog.txt"
    $Username = $env:USERNAME
    $CopiedFilesDir = "$CWD\CopiedFiles"
    $exportPath = "$CWD\Logs\Reports"
    $script:allComputers = @()
    $InformationPreference = 'SilentlyContinue'

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Defending Off the Land"
    $form.Size = New-Object System.Drawing.Size(590, 715)
    $form.StartPosition = "CenterScreen"
    $form.ForeColor = [System.Drawing.Color]::lightseagreen   
    $form.BackColor = [System.Drawing.Color]::Black
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Font = New-Object System.Drawing.Font("Times New Roman", 8, [System.Drawing.FontStyle]::Regular)
    $Form.Font = $Font

    $textboxResults = New-Object System.Windows.Forms.TextBox
    $textboxResults.Location = New-Object System.Drawing.Point(15, 380)
    $textboxResults.Size = New-Object System.Drawing.Size(540, 300)
    $textboxResults.Multiline = $true
    $textboxResults.ScrollBars = "Vertical"
    $textboxResults.WordWrap = $true
    $textboxResults.BackColor = [System.Drawing.Color]::Black
    $textboxResults.ForeColor = [System.Drawing.Color]::lightseagreen
    $textboxResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $form.Controls.Add($textboxResults)

    $ScriptStartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Log_Message -Message "Session started at $ScriptStartTime" -LogFile $LogFile
    $textboxResults.AppendText("Session started at $ScriptStartTime `r`n")

    $btnPopOut = New-Object System.Windows.Forms.Button
    $btnPopOut.Text = "Expand View"
    $btnPopOut.Width = 80
    $btnPopOut.Height = 20
    $btnPopOut.BackColor = [System.Drawing.Color]::Black
    $btnPopOut.ForeColor = [System.Drawing.Color]::lightseagreen
    $btnPopOut.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnPopOut.FlatAppearance.BorderSize = 1
    $btnPopOut.Location = New-Object System.Drawing.Point(440, 0) 
    $btnPopOut.Add_Click({
        $popupForm = New-Object System.Windows.Forms.Form
        $popupForm.Text = "Results"
        $popupForm.Width = 800
        $popupForm.Height = 600

        $layout = New-Object System.Windows.Forms.TableLayoutPanel
        $layout.Dock = [System.Windows.Forms.DockStyle]::Fill
        $layout.RowCount = 2
        $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 97)))
        $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 3)))

        $popupTextbox = New-Object System.Windows.Forms.TextBox
        $popupTextbox.Dock = [System.Windows.Forms.DockStyle]::Fill
        $popupTextbox.Multiline = $true
        $popupTextbox.ScrollBars = "Vertical"
        $popupTextbox.WordWrap = $true
        $popupTextbox.BackColor = $textboxResults.BackColor
        $popupTextbox.ForeColor = $textboxResults.ForeColor
        $popupTextbox.Text = $textboxResults.Text
        $popupTextbox.Font = New-Object System.Drawing.Font($popupTextbox.Font.FontFamily, 16)

        $layout.Controls.Add($popupTextbox, 0, 0)

        $buttonPanel = New-Object System.Windows.Forms.Panel
        $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = "Save"
        $btnSave.Location = New-Object System.Drawing.Point(100, 0)
        $btnSave.Add_Click({
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Set-Content -Path $saveFileDialog.FileName -Value $textboxResults.Text
            }
        })

        $buttonPanel.Controls.Add($btnSave)
        $layout.Controls.Add($buttonPanel, 0, 1)
        $popupForm.Controls.Add($layout)
        $popupForm.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        $popupForm.Show()
    })
    $textboxResults.Controls.Add($btnPopOut)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(250, 0)
    $statusLabel.Size = New-Object System.Drawing.Size(200, 20)
    $statusLabel.Text = ""
    $form.Controls.Add($statusLabel)

    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $helpButton.Location = New-Object System.Drawing.Point(15, 20) 
    $helpButton.Size = New-Object System.Drawing.Size(30, 20) 
    $helpButton.Text = "?"
    $helpButton.Add_Click({ & ".\Tools\Scripts\small_dol_help.ps1" })
    $form.Controls.Add($helpButton)

    $buttonSysInfo = New-Object System.Windows.Forms.Button
    $buttonSysInfo.Location = New-Object System.Drawing.Point(315, 20)
    $buttonSysInfo.Size = New-Object System.Drawing.Size(80, 40)
    $buttonSysInfo.Text = "Reconnoiter"
    $buttonSysInfo.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonSysInfo.Add_Click({
        $statusLabel.Text = "Working..."
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        & ".\Tools\Scripts\small_dol_recon.ps1"
        Log_Message -logfile $LogFile -Message "System and User AD Information exported from $computerName to Excel workbook"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonSysInfo)

    $buttonViewProcesses = New-Object System.Windows.Forms.Button
    $buttonViewProcesses.Location = New-Object System.Drawing.Point(395, 20)
    $buttonViewProcesses.Size = New-Object System.Drawing.Size(80, 40)
    $buttonViewProcesses.Text = "View Processes"
    $buttonViewProcesses.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonViewProcesses.Add_Click({ .\Tools\Scripts\small_dol_viewproc.ps1 })
    $form.Controls.Add($buttonViewProcesses)

    $buttonRestart = New-Object System.Windows.Forms.Button
    $buttonRestart.Location = New-Object System.Drawing.Point(475, 20)
    $buttonRestart.Size = New-Object System.Drawing.Size(80, 40)
    $buttonRestart.Text = "Restart Host"
    $buttonRestart.Anchor = [System.Windows.Forms.AnchorStyles]::Top 
    $buttonRestart.Add_Click({
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        .\Tools\Scripts\small_dol_restart.ps1
        Log_Message -logfile $LogFile -Message "Restarted $computerName" })
    $form.Controls.Add($buttonRestart)

    $comboboxComputerName = New-Object System.Windows.Forms.ComboBox
    $comboboxComputerName.Location = New-Object System.Drawing.Point(150, 30)
    $comboboxComputerName.Size = New-Object System.Drawing.Size(150, 20)
    $comboboxComputerName.BackColor = [System.Drawing.Color]::Black
    $comboboxComputerName.ForeColor = [System.Drawing.Color]::lightseagreen
    $comboBoxComputerName.AutoCompleteMode = 'Suggest'
    $comboBoxComputerName.AutoCompleteSource = 'ListItems'
    $comboBoxComputerName.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($comboboxComputerName)

    $labelComputerName = New-Object System.Windows.Forms.Label
    $labelComputerName.Location = New-Object System.Drawing.Point(53, 32)
    $labelComputerName.Size = New-Object System.Drawing.Size(95, 20)
    $labelComputerName.Text = "Remote Computer"
    $labelComputerName.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labelComputerName)

    $buttonGetHostList = New-Object System.Windows.Forms.Button
    $buttonGetHostList.Text = 'Get Host List'
    $buttonGetHostList.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonGetHostList.Location = New-Object System.Drawing.Point(53, 50)
    $buttonGetHostList.Size = New-Object System.Drawing.Size(90, 23)
    $buttonGetHostList.Add_Click({
        $statusLabel.Text = "Working..."
        $comboBoxComputerName.Items.Clear()
        $script:allComputers = Get-AllComputers
        $comboBoxComputerName.Items.AddRange($script:allComputers)
        $statusLabel.Text = "Done"
    })
    $Form.Controls.Add($buttonGetHostList)

    $buttonKillProcess = New-Object System.Windows.Forms.Button
    $buttonKillProcess.Location = New-Object System.Drawing.Point(315, 60)
    $buttonKillProcess.Size = New-Object System.Drawing.Size(80, 40)
    $buttonKillProcess.Text = "Kill Process"
    $buttonKillProcess.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonKillProcess.Add_Click({
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        $selectedProcess = if ($dropdownProcessId.SelectedItem) { $dropdownProcessId.SelectedItem.ToString() } else { $dropdownProcessId.Text }
        .\Tools\Scripts\small_dol_killproc.ps1
        Log_Message -logfile $LogFile -Message "Terminated $selectedProcess on $computerName"
    })
    $form.Controls.Add($buttonKillProcess)

    $buttonCopyBinaries = New-Object System.Windows.Forms.Button
    $buttonCopyBinaries.Location = New-Object System.Drawing.Point(395, 60)
    $buttonCopyBinaries.Size = New-Object System.Drawing.Size(80, 40)
    $buttonCopyBinaries.Text = "Copy   Modules"
    $buttonCopyBinaries.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonCopyBinaries.Add_Click({
        $statusLabel.Text = "Working..."
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($CopiedFilesDir)) {
            .\Tools\Scripts\small_dol_copymodules.ps1 -CopiedFilesPath $CopiedFilesDir
            Log_Message -logfile $LogFile -Message "Copied all modules from $computerName"
        }
        else { [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonCopyBinaries)

    $buttonShutdown = New-Object System.Windows.Forms.Button
    $buttonShutdown.Location = New-Object System.Drawing.Point(475, 60)
    $buttonShutdown.Size = New-Object System.Drawing.Size(80, 40)
    $buttonShutdown.Text = "Shutdown Host"
    $buttonShutdown.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonShutdown.Add_Click({
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        if (![string]::IsNullOrEmpty($computerName)) {
            .\Tools\Scripts\small_dol_shutdown.ps1 -computerName $computerName
            Log_Message -logfile $LogFile -Message "Shutdown $computerName"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($buttonShutdown)

    $dropdownProcessId = New-Object System.Windows.Forms.ComboBox
    $dropdownProcessId.Location = New-Object System.Drawing.Point(150, 85)
    $dropdownProcessId.Size = New-Object System.Drawing.Size(150, 20)
    $dropdownProcessId.BackColor = [System.Drawing.Color]::Black
    $dropdownProcessId.ForeColor = [System.Drawing.Color]::lightseagreen
    $dropdownProcessId.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($dropdownProcessId)

    $labelProcessId = New-Object System.Windows.Forms.Label
    $labelProcessId.Location = New-Object System.Drawing.Point(53, 87)
    $labelProcessId.Size = New-Object System.Drawing.Size(95, 20)
    $labelProcessId.Text = "Select a Process"
    $labelProcessId.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labelProcessId)

    $RapidTriageButton = New-Object System.Windows.Forms.Button
    $RapidTriageButton.Location = New-Object System.Drawing.Point(315, 100)
    $RapidTriageButton.Size = New-Object System.Drawing.Size(80, 40)
    $RapidTriageButton.Text = 'RapidTriage'
    $RapidTriageButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $RapidTriageButton.Add_Click({
        $statusLabel.Text = "Working..."
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        & ".\Tools\Scripts\small_dol_rapidtriage.ps1" -computerName $computerName -exportPath $exportPath
        Log_Message -logfile $LogFile -Message "RapidTriage completed on $computerName"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($RapidTriageButton)

    $WinEventalyzerButton = New-Object System.Windows.Forms.Button
    $WinEventalyzerButton.Location = New-Object System.Drawing.Point(395, 100)
    $WinEventalyzerButton.Size = New-Object System.Drawing.Size(80, 40)
    $WinEventalyzerButton.Text = 'Win-Eventalyzer'
    $WinEventalyzerButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $WinEventalyzerButton.Add_Click({
        $statusLabel.Text = "Working..."
        $computerName = if ($comboBoxComputerName.SelectedItem) { $comboBoxComputerName.SelectedItem.ToString() } else { $comboBoxComputerName.Text }
        try {
            .\Tools\Scripts\small_dol_wineventalyzer.ps1
            $statusLabel.Text = "Done"
            Log_Message -logfile $LogFile -Message "WinEventalyzed $computerName"
            write-host "WinEventalyzer completed on $computerName" -ForegroundColor Green
            $textboxResults.AppendText("WinEventalyzer completed on $computerName")
        }
        catch {
            $statusLabel.Text = "Error"
            Write-Host $_.Exception.Message
        }
    })
    $form.Controls.Add($WinEventalyzerButton)

    $buttonProcAsso = New-Object System.Windows.Forms.Button
    $buttonProcAsso.Location = New-Object System.Drawing.Point(475, 100)
    $buttonProcAsso.Size = New-Object System.Drawing.Size(80, 40)
    $buttonProcAsso.Text = "ProcAsso"
    $buttonProcAsso.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonProcAsso.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_procasso.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonProcAsso)

    $labelTractorBeam = New-Object System.Windows.Forms.Label
    $labelTractorBeam.Location = New-Object System.Drawing.Point(53, 120)
    $labelTractorBeam.Size = New-Object System.Drawing.Size(100, 20)
    $labelTractorBeam.Text = "Tractor Beam"
    $labelTractorBeam.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labelTractorBeam)

    $buttonIsolateHost = New-Object System.Windows.Forms.Button
    $buttonIsolateHost.Location = New-Object System.Drawing.Point(15, 140)
    $buttonIsolateHost.Size = New-Object System.Drawing.Size(70, 40)
    $buttonIsolateHost.Text = "Engage"
    $buttonIsolateHost.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonIsolateHost.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_isolate.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonIsolateHost)

    $buttonUndoIsolation = New-Object System.Windows.Forms.Button
    $buttonUndoIsolation.Location = New-Object System.Drawing.Point(85, 140)
    $buttonUndoIsolation.Size = New-Object System.Drawing.Size(70, 40)
    $buttonUndoIsolation.Text = "Release"
    $buttonUndoIsolation.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonUndoIsolation.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_release.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonUndoIsolation)

    $buttonInstallSysmon = New-Object System.Windows.Forms.Button
    $buttonInstallSysmon.Location = New-Object System.Drawing.Point(160, 140)
    $buttonInstallSysmon.Size = New-Object System.Drawing.Size(70, 40)
    $buttonInstallSysmon.Text = "Deploy Sysmon"
    $buttonInstallSysmon.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonInstallSysmon.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_sysmon.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonInstallSysmon)

    $labelRemoteFilePath = New-Object System.Windows.Forms.Label
    $labelRemoteFilePath.Location = New-Object System.Drawing.Point(260, 142)
    $labelRemoteFilePath.Size = New-Object System.Drawing.Size(50, 30)
    $labelRemoteFilePath.Text = "Remote File Path"
    $labelRemoteFilePath.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labelRemoteFilePath)

    $textboxremoteFilePath = New-Object System.Windows.Forms.TextBox
    $textboxremoteFilePath.Location = New-Object System.Drawing.Point(310, 150)
    $textboxremoteFilePath.Size = New-Object System.Drawing.Size(245, 20)
    $textboxremoteFilePath.BackColor = [System.Drawing.Color]::Black
    $textboxremoteFilePath.ForeColor = [System.Drawing.Color]::lightseagreen
    $textboxremoteFilePath.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($textboxremoteFilePath)

    $buttonSelectRemoteFile = New-Object System.Windows.Forms.Button
    $buttonSelectRemoteFile.Text = "Browse Remote Files"
    $buttonSelectRemoteFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonSelectRemoteFile.Size = New-Object System.Drawing.Size(80, 40)
    $buttonSelectRemoteFile.Location = New-Object System.Drawing.Point(245, 180)
    $buttonSelectRemoteFile.Add_Click({ & ".\Tools\Scripts\small_dol_remotefilebrowse.ps1" })
    $form.Controls.Add($buttonSelectRemoteFile)

    $buttonCopyFile = New-Object System.Windows.Forms.Button
    $buttonCopyFile.Location = New-Object System.Drawing.Point(330, 180)
    $buttonCopyFile.Size = New-Object System.Drawing.Size(75, 40)
    $buttonCopyFile.Text = "Copy"
    $buttonCopyFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonCopyFile.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_copy.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonCopyFile)

    $buttonHuntFile = New-Object System.Windows.Forms.Button
    $buttonHuntFile.Location = New-Object System.Drawing.Point(405, 180)
    $buttonHuntFile.Size = New-Object System.Drawing.Size(75, 40)
    $buttonHuntFile.Text = "Hunt File"
    $buttonHuntFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonHuntFile.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_hunt.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonHuntFile)

    $buttonDeleteFile = New-Object System.Windows.Forms.Button
    $buttonDeleteFile.Location = New-Object System.Drawing.Point(480, 180)
    $buttonDeleteFile.Size = New-Object System.Drawing.Size(75, 40)
    $buttonDeleteFile.Text = "Delete"
    $buttonDeleteFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonDeleteFile.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_delete.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonDeleteFile)

    $labelLocalFilePath = New-Object System.Windows.Forms.Label
    $labelLocalFilePath.Location = New-Object System.Drawing.Point(258, 230)
    $labelLocalFilePath.Size = New-Object System.Drawing.Size(50, 30)
    $labelLocalFilePath.Text = "Local  File Path"
    $labelLocalFilePath.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labelLocalFilePath)

    $comboboxlocalFilePath = New-Object System.Windows.Forms.ComboBox
    $comboboxlocalFilePath.Location = New-Object System.Drawing.Point(310, 237)
    $comboboxlocalFilePath.Size = New-Object System.Drawing.Size(245, 20)
    $comboboxlocalFilePath.BackColor = [System.Drawing.Color]::Black
    $comboboxlocalFilePath.ForeColor = [System.Drawing.Color]::lightseagreen
    $comboboxlocalFilePath.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $comboboxlocalFilePath.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $comboboxlocalFilePath.add_DrawItem({
        param($senderloc, $e)
        $e.DrawBackground()
        $text = $senderloc.Items[$e.Index]
        $textFormatFlags = [System.Windows.Forms.TextFormatFlags]::Right
        [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $text, $e.Font, $e.Bounds, $e.ForeColor, $textFormatFlags)
        $e.DrawFocusRectangle()
    })
    $comboboxlocalFilePath.add_MouseEnter({
        if ($comboboxlocalFilePath.SelectedItem) {
            $toolTip.SetToolTip($comboboxlocalFilePath, $comboboxlocalFilePath.SelectedItem.ToString())
        }
        else {
            $toolTip.SetToolTip($comboboxlocalFilePath, "")
        }
    })
    $comboboxlocalFilePath.add_SelectedIndexChanged({
        if ($comboboxlocalFilePath.SelectedItem) {
            $toolTip.SetToolTip($comboboxlocalFilePath, $comboboxlocalFilePath.SelectedItem.ToString())
        }
        else {
            $toolTip.SetToolTip($comboboxlocalFilePath, "")
        }
    })
    $form.Controls.Add($comboboxlocalFilePath)

    $buttonSelectLocalFile = New-Object System.Windows.Forms.Button
    $buttonSelectLocalFile.Text = "Browse Local Files"
    $buttonSelectLocalFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonSelectLocalFile.Size = New-Object System.Drawing.Size(80, 40)
    $buttonSelectLocalFile.Location = New-Object System.Drawing.Point(285, 315)
    $buttonSelectLocalFile.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "All files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFile = $openFileDialog.FileName
            $comboboxlocalFilePath.Text = $selectedFile
        }
    })
    $form.Controls.Add($buttonSelectLocalFile)

    $buttonListCopiedFiles = New-Object System.Windows.Forms.Button
    $buttonListCopiedFiles.Location = New-Object System.Drawing.Point(330, 270)
    $buttonListCopiedFiles.Size = New-Object System.Drawing.Size(75, 40)
    $buttonListCopiedFiles.Text = "View Copied Files"
    $buttonListCopiedFiles.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonListCopiedFiles.Add_Click({ & ".\Tools\Scripts\small_dol_listcopied.ps1" })
    $form.Controls.Add($buttonListCopiedFiles)

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Location = New-Object System.Drawing.Point(53, 190)
    $usernameLabel.Size = New-Object System.Drawing.Size(200, 15)
    $usernameLabel.Text = "Enter or Select Username"
    $usernameLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($usernameLabel)

    $comboboxUsername = New-Object System.Windows.Forms.ComboBox
    $comboboxUsername.Location = New-Object System.Drawing.Point(15, 205)
    $comboboxUsername.Size = New-Object System.Drawing.Size(215, 20)
    $comboboxUsername.BackColor = [System.Drawing.Color]::Black
    $comboboxUsername.ForeColor = [System.Drawing.Color]::lightseagreen
    $comboboxUsername.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($comboboxUsername)

    $buttonPWChange = New-Object System.Windows.Forms.Button
    $buttonPWChange.Location = New-Object System.Drawing.Point(15, 225)
    $buttonPWChange.Size = New-Object System.Drawing.Size(55, 40)
    $buttonPWChange.Text = 'PW Reset'
    $buttonPWChange.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonPWChange.Add_Click({ & ".\Tools\Scripts\small_dol_pwchange.ps1" })
    $form.Controls.Add($buttonPWChange)

    $buttonLogOff = New-Object System.Windows.Forms.Button
    $buttonLogOff.Location = New-Object System.Drawing.Point(70, 225)
    $buttonLogOff.Size = New-Object System.Drawing.Size(55, 40)
    $buttonLogOff.Text = "Logoff"
    $buttonLogOff.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonLogOff.Add_Click({ & ".\Tools\Scripts\small_dol_logoff.ps1" })
    $form.Controls.Add($buttonLogOff)

    $buttonDisableAcc = New-Object System.Windows.Forms.Button
    $buttonDisableAcc.Location = New-Object System.Drawing.Point(125, 225)
    $buttonDisableAcc.Size = New-Object System.Drawing.Size(55, 40)
    $buttonDisableAcc.Text = "Disable"
    $buttonDisableAcc.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonDisableAcc.Add_Click({ & ".\Tools\Scripts\small_dol_disableacc.ps1" })
    $form.Controls.Add($buttonDisableAcc)

    $buttonEnableAcc = New-Object System.Windows.Forms.Button
    $buttonEnableAcc.Location = New-Object System.Drawing.Point(180, 225)
    $buttonEnableAcc.Size = New-Object System.Drawing.Size(50, 40)
    $buttonEnableAcc.Text = "Enable"
    $buttonEnableAcc.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonEnableAcc.Add_Click({ & ".\Tools\Scripts\small_dol_enableacc.ps1" })
    $form.Controls.Add($buttonEnableAcc)

    $textboxaddargs = New-Object System.Windows.Forms.TextBox
    $textboxaddargs.Location = New-Object System.Drawing.Point(80, 275)
    $textboxaddargs.Size = New-Object System.Drawing.Size(245, 20)
    $textboxaddargs.BackColor = [System.Drawing.Color]::Black
    $textboxaddargs.ForeColor = [System.Drawing.Color]::lightseagreen
    $textboxaddargs.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($textboxaddargs)

    $labeladdargs = New-Object System.Windows.Forms.Label
    $labeladdargs.Location = New-Object System.Drawing.Point(15, 275)
    $labeladdargs.Size = New-Object System.Drawing.Size(60, 20)
    $labeladdargs.Text = "Arguments"
    $labeladdargs.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $form.Controls.Add($labeladdargs)

    $buttonIntelligizer = New-Object System.Windows.Forms.Button
    $buttonIntelligizer.Location = New-Object System.Drawing.Point(480, 270)
    $buttonIntelligizer.Size = New-Object System.Drawing.Size(75, 40)
    $buttonIntelligizer.Text = "Intelligazer"
    $buttonIntelligizer.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonIntelligizer.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_intelligazer.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonIntelligizer)

    $buttonPlaceFile = New-Object System.Windows.Forms.Button
    $buttonPlaceFile.Location = New-Object System.Drawing.Point(405, 270)
    $buttonPlaceFile.Size = New-Object System.Drawing.Size(75, 40)
    $buttonPlaceFile.Text = "Place File"
    $buttonPlaceFile.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonPlaceFile.Add_Click({ & ".\Tools\Scripts\small_dol_place.ps1" })
    $form.Controls.Add($buttonPlaceFile)

    $buttonPlaceAndRun = New-Object System.Windows.Forms.Button
    $buttonPlaceAndRun.Location = New-Object System.Drawing.Point(100, 305)
    $buttonPlaceAndRun.Size = New-Object System.Drawing.Size(75, 40)
    $buttonPlaceAndRun.Text = "Place and Run"
    $buttonPlaceAndRun.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonPlaceAndRun.Add_Click({
        $statusLabel.Text = "Working..."
        & ".\Tools\Scripts\small_dol_placerun.ps1"
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($buttonPlaceAndRun)

    $executeCommandButton = New-Object System.Windows.Forms.Button
    $executeCommandButton.Text = "Execute Oneliner"
    $executeCommandButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $executeCommandButton.Location = New-Object System.Drawing.Point(180, 305)
    $executeCommandButton.Size = New-Object System.Drawing.Size(75, 40)
    $executeCommandButton.Add_Click({
        $statusLabel.Text = "Working..."
        $computerName = if ($comboBoxComputerName.SelectedItem) {
            $comboBoxComputerName.SelectedItem.ToString()
        }
        else { $comboBoxComputerName.Text }
        $command = $textboxaddargs.Text
        $output = & ".\Tools\Scripts\small_dol_oneliner.ps1" -ComputerNameFromMain $computerName -CommandFromMain $command
        $textboxResults.AppendText($output)
        Log_Message -Message $output -LogFilePath $LogFile
        $statusLabel.Text = "Done"
    })
    $form.Controls.Add($executeCommandButton)

    $buttonReset = New-Object System.Windows.Forms.Button
    $buttonReset.Location = New-Object System.Drawing.Point(15, 305)
    $buttonReset.Size = New-Object System.Drawing.Size(80, 23)
    $buttonReset.Text = "Reset"
    $buttonReset.Anchor = [System.Windows.Forms.AnchorStyles]::Top
    $buttonReset.Add_Click({ & ".\Tools\Scripts\small_dol_resetform.ps1" -ScriptStartTime $ScriptStartTime })
    $form.Controls.Add($buttonReset)

    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($buttonViewProcesses, "Click to view the list of running processes on the selected remote computer. `r`n This button populates the Select a Process dropdown")
    $tooltip.SetToolTip($buttonCopyBinaries, "Click to copy all uniquely pathed modules currently running on the remote computer.")
    $tooltip.SetToolTip($buttonIsolateHost, "Click to isolate the remote computer by blocking all IP addresses except the local host.")
    $tooltip.SetToolTip($buttonHuntFile, "Click to hunt for a remote file from the remote computer.")
    $tooltip.SetToolTip($buttonPlaceFile, "Click to place the local file you want on the remote computer to the remote file path.")
    $tooltip.SetToolTip($buttonPlaceAndRun, "Click to place and run the local file you want on the remote computer to the remote file path. `r`n Use the 'Arguments' textbox if any are required for execution.")
    $tooltip.SetToolTip($executeCommandButton, "Enter the command you want to execute on the remote computer in the 'Arguments' text box, then click this button. `r`n Use set-location to execute outside of System32.")
    $tooltip.SetToolTip($buttonInstallSysmon, "Click to deploy Sysmon on the remote computer using the Olaf Hartong default configuration.")
    $tooltip.SetToolTip($RapidTriageButton, "Click to run the RapidTriage tool on the remote computer.")
    $tooltip.SetToolTip($buttonReset, "Click to reset the form and clear all input fields.")
    $tooltip.SetToolTip($helpButton, "Click to view a detailed explanation of each functionality of the script.")
    $tooltip.SetToolTip($WinEventalyzerButton, "Click to copy and threat hunt windows event logs from remote computer.")
    $tooltip.SetToolTip($buttonListCopiedFiles, "Click to view the list of files copied from the remote computer. `r`n This button populates the Local File Path dropdown.")
    $tooltip.SetToolTip($buttonGet
