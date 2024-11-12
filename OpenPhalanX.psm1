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

    $tooltip.SetToolTip($buttonProcAsso, "Click to associate activity with each running process on the remote host.")
    $tooltip.SetToolTip($buttonSelectRemoteFile, "Click to open a custom remote file system exporer.")
    $tooltip.SetToolTip($buttonSelectLocalFile, "Click to select a file from your local machine for use.")
    $tooltip.SetToolTip($buttonPWChange, "Click to force the specified user to change their password.")
    $tooltip.SetToolTip($buttonLogOff, "Click to force the spefified user off of the remote computer.")
    $tooltip.SetToolTip($buttonDisableAcc, "Click to disable the specified user's account")
    $tooltip.SetToolTip($buttonEnableAcc, "Click to enable the specified user's account.")
    $tooltip.SetToolTip($buttonSysInfo, "Click to retrieve AD info on the specified remote host and all users who have logged into this host. `r`n This buttton populates the Username dropdown.")
    $tooltip.SetToolTip($buttonRestart, "Click to command the remote computer to restart.")
    $tooltip.SetToolTip($buttonKillProcess, "Click to kill the selected process from the remote computer.")
    $tooltip.SetToolTip($buttonShutdown, "Click to command the remote computer to shutdown.")
    $tooltip.SetToolTip($buttonCopyFile, "Click to copy the file specified in the Remote File Path to the CopiedFiles directory on the local host.")
    $tooltip.SetToolTip($buttonDeleteFile, "Delete the file specified in the Remote File Path from the remote host.")
    $tooltip.SetToolTip($buttonIntelligizer, "Click to scrape indicators from collected reports and logs")
    $tooltip.SetToolTip($buttonUndoIsolation, "Click to remove firewall rules applied for isolation.")

    $Form.Add_FormClosing({
        $ScriptEndTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Log_Message -Message "Your session ended at $ScriptEndTime" -LogFile $LogFile
        $textboxResults.AppendText("Your session ended at $ScriptEndTime")
    })
    $Form.ShowDialog()
}

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

function processMatches($content, $file) {
    $patterns = @(
        'HTTP/S' = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
        'FTP' = 'ftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'SFTP' = 'sftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'SCP' = 'scp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'DATA' = 'data://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
        'SSH' = 'ssh://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?',
        'LDAP' = 'ldap://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'RFC 1918 IP Address' = '\b(10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3})\b',
        'Non-RFC 1918 IP Address' = '\b((?!10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|127\.\d{1,3}\.\d{1,3}\.\d{1,3})\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b',
        'Email' = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    )

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

function Copy-RemoteFile {
    param (
        [string]$computerName,
        [string]$filePath,
        [string]$CopiedFilesDir,
        [string]$logfile
    )

    $filecopy = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Copied $filePath from $computerName at $filecopy" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Copied $filePath from $computerName"

    if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($filePath)) {
        $driveLetter = $filePath.Substring(0, 1)
        $uncPath = "\\$computerName\$driveLetter$" + $filePath.Substring(2)
        try {
            if ((Test-Path -Path $uncPath) -and (Get-Item -Path $uncPath).PSIsContainer) {
                $copyDecision = [System.Windows.Forms.MessageBox]::Show("Do you want to copy the directory and all its child directories? Select No to only copy files from specified directory", "Copy Directory Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($copyDecision -eq "Yes") {
                    Copy-Item -Path $uncPath -Destination $CopiedFilesDir -Force -Recurse
                } elseif ($copyDecision -eq "No") {
                    Get-ChildItem -Path $uncPath -File | ForEach-Object {
                        Copy-Item -Path $_.FullName -Destination $CopiedFilesDir -Force
                    }
                } else {
                    return
                }
            } else {
                Copy-Item -Path $uncPath -Destination $CopiedFilesDir -Force
            }

            [System.Windows.Forms.MessageBox]::Show("File '$filePath' copied to local directory.", "Copy File Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $textboxResults.AppendText("File '$filePath' copied to local directory.`r`n")
            $filecopy = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "Copied $filePath from $computerName at $filecopy" -ForegroundColor Cyan 
            Log_Message -logfile $logfile -Message "Copied $filePath from $computerName"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error copying file '$filePath' to local directory: $_", "Copy File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $textboxResults.AppendText("Error copying file '$filePath' to local directory: $_`r`n")
        }
    }
}

function Copy-RemoteModules {
    param(
        [string]$computerName,
        [string]$CopiedFilesPath
    )

    function Get_RemoteDriveLetters {
        param(
            [string]$ComputerName
        )

        $driveLetters = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object { $_.DeviceID.TrimEnd(':') }
        }

        return $driveLetters
    }

    function CopyModules {
        param (
            [string]$ComputerName,
            [string]$CopiedFilesPath
        )

        $uniqueModules = Invoke-Command -ComputerName $computerName -ScriptBlock {
            Get-Process | ForEach-Object { $_.Modules } | Select-Object -Unique -ExpandProperty FileName
        }

        $uniqueModules | ForEach-Object -ThrottleLimit 100 -Parallel {
            $modulePath = $_
            if (![string]::IsNullOrEmpty($modulePath)) {
                $driveLetter = $modulePath.Substring(0, 1)
                $uncPath = "\\$using:computerName\$driveLetter$" + $modulePath.Substring(2)
                $modifiedModuleName = $modulePath.Replace('\', '_').Replace(':', '')
                $destinationPath = Join-Path -Path $using:CopiedFilesPath -ChildPath $modifiedModuleName
                Copy-Item $uncPath -Destination $destinationPath\$ComputerName -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $copytreestart = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $textboxResults.AppendText("Started copying all modules from $computerName at $copytreestart`r`n")
    write-host "Started copying all modules from $computerName at $copytreestart" -ForegroundColor Green

    try {
        CopyModules -computerName $computerName -CopiedFilesPath $CopiedFilesPath
        $copytreesend = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $textboxResults.AppendText("Copied all modules from $computerName at $copytreesend`r`n")
        write-host "Copied all modules from $computerName at $copytreesend" -ForegroundColor Green
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error while trying to copy modules to the CopiedFiles folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Delete-RemoteFile {
    param (
        [string]$computerName,
        [string]$filePath,
        [string]$logfile
    )

    if (![string]::IsNullOrEmpty($computerName) -and ![string]::IsNullOrEmpty($filePath)) {
        $filename = [System.IO.Path]::GetFileName($filePath)
        try {
            $result = [System.Windows.Forms.MessageBox]::Show("Do you want to delete '$filePath'?", "Delete File or Directory", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($result -eq "Yes") {
                Invoke-Command -ComputerName $computerName -ScriptBlock { param($path) Remove-Item -Path $path -Recurse -Force -ErrorAction Stop } -ArgumentList $filePath
                [System.Windows.Forms.MessageBox]::Show("'$filename' deleted from remote host.", "Delete Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $textboxResults.AppendText("'$filename' deleted from remote host.`r`n")
            }
            $filedelete = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-Host "Deleted $filePath from $computerName at $filedelete" -ForegroundColor Cyan 
            Log_Message -logfile $logfile -Message "Deleted $filePath from $computerName"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error deleting '$filename' from remote host: $_", "Delete Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $textboxResults.AppendText("Error deleting '$filename' from remote host: $_`r`n")
        }
    }
}

function Disable-RemoteAccount {
    param (
        [string]$selectedUser,
        [string]$Username,
        [string]$logfile
    )

    Disable-ADAccount -Identity $selectedUser  
    $gotdisabusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$Username disabled $selectedUser at $gotdisabusr" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "$Username disabled $selectedUser"
    $textboxResults.AppendText("$Username disabled $selectedUser at $gotdisabusr")
}

function Enable-RemoteAccount {
    param (
        [string]$selectedUser,
        [string]$Username,
        [string]$logfile
    )

    Enable-ADAccount -Identity $selectedUser
    $gotenabusr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$Username enabled $selectedUser at $gotenabusr" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "$Username enabled $selectedUser"
    $textboxResults.AppendText("$Username enabled $selectedUser at $gotenabusr")
}

function Get-FileScan {
    param (
        [string]$fileName,
        [string]$textboxResults
    )

    $Response = Invoke-RestMethod -Uri "https://api.threatstream.com/api/v1/submit/search/?q=$script:Note" -Headers @{"Authorization" = "apikey redacted"} -Method Get 
    if ($Response.meta.total_count -gt 0) {
        $textboxResults.AppendText("Total Count: " + $Response.meta.total_count + [Environment]::NewLine)
        foreach ($report in $Response.objects) {
            $textboxResults.AppendText(($report | ConvertTo-Json -Depth 100) + [Environment]::NewLine)
        }
    } else {
        $textboxResults.AppendText("No reports found for file: $fileName. Please wait and try again.")
    }
}

function Get-HostList {
    return Get-AllComputers
}

function Get-Intel {
    param (
        [string]$IndicatorFromMain,
        [string]$directoryForScanning,
        [string]$pythonExe
    )

    $integrationFiles = Get-ChildItem -Path $directoryForScanning -Filter "Integration_*_Intel.py"

    foreach ($file in $integrationFiles) {
        if ($file.Name -match "Integration_(.+?)_Intel\.py") {
            $providerName = $matches[1]
            $confirmation = [System.Windows.Forms.MessageBox]::Show(
                "Do you want to use the $providerName API provider to fetch intel for the indicator?", 
                "Confirm Provider Usage", 
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($confirmation -eq 'Cancel') {
                return "Action cancelled by user."
            } elseif ($confirmation -eq 'Yes') {
                $scriptPath = Join-Path $directoryForScanning $file.Name
                $output = & $pythonExe $scriptPath $IndicatorFromMain | Out-String

                if ($output -contains "specific-string") {
                }
                return "Python script output for $providerName $output"
            }
        }
    }
}

function Show-Help {
    param([string]$message, [string]$title)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(910, 600)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::Black
    $form.ForeColor = [System.Drawing.Color]::lightseagreen

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(880, 500) 
    $panel.Location = New-Object System.Drawing.Point(10, 10)
    $panel.AutoScroll = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.ForeColor = [System.Drawing.Color]::LightSeaGreen 
    $label.Font = New-Object System.Drawing.Font("Arial", 12)

    $updateButton = New-Object System.Windows.Forms.Button
    $updateButton.Text = "Update Phalanx"
    $updateButton.Location = New-Object System.Drawing.Point(350, 520)
    $updateButton.Size = New-Object System.Drawing.Size(200, 30)
    $updateButton.Add_Click({
        & ".\Tools\Scripts\Update_Phalanx.ps1"
    })

    $panel.Controls.Add($label)
    $form.Controls.Add($panel)
    $form.Controls.Add($updateButton)

    $form.ShowDialog()
}

function Hunt-RemoteFile {
    param (
        [string]$remoteComputer,
        [string]$remoteFilePath,
        [string]$logfile,
        [string]$textboxResults
    )

    $drivelessPath = Split-Path $remoteFilePath -NoQualifier
    $driveLetter = Split-Path $remoteFilePath -Qualifier
    $filehunt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Hunting for file: $remoteFilePath at $filehunt" -ForegroundColor Cyan
    $textboxResults.AppendText("Hunting for file: $remoteFilePath at $filehunt `r`n")
    Log_Message -logfile $logfile -Message "Hunting for file: $remoteFilePath"
    if (![string]::IsNullOrEmpty($remoteComputer) -and ![string]::IsNullOrEmpty($remoteFilePath)) {
        $localPath = '.\CopiedFiles'
        if (!(Test-Path $localPath)) {
            New-Item -ItemType Directory -Force -Path $localPath
        }

        try {
            $fileExists = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                param($path)
                Test-Path -Path $path
            } -ArgumentList $remoteFilePath -ErrorAction Stop
        } catch {
            Write-Error "Failed to check if the file exists on the remote computer: $_"
            return
        }

        if ($fileExists) {
            $destination = Join-Path -Path $localPath -ChildPath (Split-Path $remoteFilePath -Leaf)
            Copy-Item -Path "\\$remoteComputer\$($remoteFilePath.Replace(':', '$'))" -Destination $destination -Force
            $textboxResults.AppendText("$remoteFilePath copied from $remoteComputer.")
            Log_Message -Message "$remoteFilePath copied from $remoteComputer" -LogFilePath $LogFile
            Write-Host "$remoteFilePath copied from $remoteComputer" -ForegroundColor -Cyan
        } else {
            $foundInRecycleBin = $false

            try {
                $recycleBinItems = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                    Get-ChildItem 'C:\$Recycle.Bin' -Recurse -Force | Where-Object { $_.PSIsContainer -eq $false -and $_.Name -like "$I*" } | ForEach-Object {
                        $originalFilePath = Get-Content $_.FullName -ErrorAction SilentlyContinue | Select-Object -First 1
                        $originalFilePath = $originalFilePath -replace '^[^\:]*\:', ''
                        [PSCustomObject]@{
                            Name = $_.Name
                            FullName = $_.FullName
                            OriginalFilePath = $originalFilePath
                        }
                    }
                } -ErrorAction Stop
            } catch {
                Write-Error "Failed to retrieve the items from the recycle bin: $_ `r`n"
                return
            }

            $recycleBinItem = $recycleBinItems | Where-Object { $_.OriginalFilePath -eq "$drivelessPath" }
            if ($recycleBinItem) {
                Write-Host "Match found in Recycle Bin: $($recycleBinItem.Name)" -Foregroundcolor Cyan
                $textboxResults.AppendText("Match found in Recycle Bin: $($recycleBinItem.Name) `r`n")
                $foundInRecycleBin = $true
            }

            if (!$foundInRecycleBin) {
                $textboxResults.AppendText("File not found in the remote computer or recycle bin. `r`n")
            } else {
                try {
                    $vssServiceStatus = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                        $service = Get-Service -Name VSS
                        $status = $service.Status
                        return $status
                    } 
                    
                    $statusCodes = @{
                        1 = "Stopped"
                        2 = "Start Pending"
                        3 = "Stop Pending"
                        4 = "Running"
                        5 = "Continue Pending"
                        6 = "Pause Pending"
                        7 = "Paused"
                    }
                    
                    $vssServiceStatusName = $statusCodes[$vssServiceStatus]
                    Write-Host "VSS service status: $vssServiceStatusName" -ForegroundColor Cyan
                    $textboxResults.AppendText("VSS service status on $remoteComputer $vssServiceStatusName `r`n")

                    if ($vssServiceStatus -eq 'Running') {
                        $shadowCopyFileExists = Invoke-Command -ComputerName $remoteComputer -ScriptBlock {
                            param($path, $driveLetter)
                            $shadowCopies = vssadmin list shadows /for=$driveLetter | Where-Object { $_ -match 'GLOBALROOT\\Device\\HarddiskVolumeShadowCopy\\d+' }
                            foreach ($shadowCopy in $shadowCopies) {
                                $shadowCopyPath = $shadowCopy -replace '.*?(GLOBALROOT\\Device\\HarddiskVolumeShadowCopy\\d+).*', '$1'
                                if (Test-Path -Path "$shadowCopyPath\$path") {
                                    return $shadowCopyPath
                                }
                            }
                        } -ArgumentList $drivelessPath, $driveLetter -ErrorAction Stop

                        if ($shadowCopyFileExists) {
                            Write-Host "Shadow copy found: $shadowCopyFileExists" -ForegroundColor Cyan
                            $textboxResults.AppendText("Shadow copy found: $shadowCopyFileExists `r`n")
                        } else {
                            Write-Host "No shadow copy found for the file." -ForegroundColor Red
                            $textboxResults.AppendText("No shadow copy found for the file. `r`n")
                        }
                    }
                } catch {
                    Write-Error "Failed to check VSS service or shadow copies: $_"
                    return
                }
            $restorePoints = Invoke-Expression -Command "wmic /Namespace:\\root\default Path SystemRestore get * /format:list"
            if ($restorePoints) {
                Write-Host "System Restore points exist on $remoteComputer" -ForegroundColor Cyan
                $textboxResults.AppendText("System Restore points exist on $remoteComputer `r`n")
            } else {
                Write-Host "No System Restore points found." -ForegroundColor Red
                $textboxResults.AppendText("No System Restore points found. `r`n")
            }
            $lastBackup = Invoke-Expression -Command "wbadmin get versions"
            if ($lastBackup) {
                Write-Host "Backups exist on $remoteComputer" -ForegroundColor Cyan
                $textboxResults.AppendText("Backups exist on $remoteComputer `r`n")
            } else {
                Write-Host "No backups found." -ForegroundColor Red
                $textboxResults.AppendText("No backups found. `r`n")
            }
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a remote computer name and file path.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
    $fileput = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    write-host "File Hunt completed at $fileput" -ForegroundColor Green 
    Log_Message -logfile $logfile -Message "File Hunt completed"
    $textboxResults.AppendText("File Hunt completed at $fileput `r`n")
}

function Run-Intelligazer {
    param (
        [string]$computerName,
        [string]$CopiedFilesDir,
        [string]$logfile,
        [string]$textboxResults
    )

    $output = New-Object System.Collections.ArrayList
    $patterns = @(
        'HTTP/S' = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
        'FTP' = 'ftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'SFTP' = 'sftp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'SCP' = 'scp://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'DATA' = 'data://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',
        'SSH' = 'ssh://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?',
        'LDAP' = 'ldap://(?:[a-zA-Z0-9]+:[a-zA-Z0-9]+@)?(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?::[0-9]{1,5})?(?:/[^\s]*)?',
        'RFC 1918 IP Address' = '\b(10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3})\b',
        'Non-RFC 1918 IP Address' = '\b((?!10\.\d{1,3}\.\d{1,3}\.\d{1,3}|172\.(1[6-9]|2[0-9]|3[0-1])\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|127\.\d{1,3}\.\d{1,3}\.\d{1,3})\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b',
        'Email' = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    )

    $includeFilePath = [System.Windows.Forms.MessageBox]::Show("Do you want to include File Paths?", "Include File Paths", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($includeFilePath -eq 'Yes') {
        $patterns['File Path'] = '(file:///)?(?![sSpP]:)[a-zA-Z]:[\\\\/].+?\.[a-zA-Z0-9]{2,5}(?=[\s,;]|$)'
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
                        $excel.Visible = $false
                        
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
                            $output.AddRange(@($matchess))
                        }
                    }
                }
            }
        }
    }     

    $copiedFiles = Get-ChildItem $CopiedFilesDir -Recurse -File | Select-Object -ExpandProperty FullName
    foreach ($file in $copiedFiles) {
        $sha256 = Get-FileHashSHA256 -FilePath $file
        [void]$output.Add((New-Object PSObject -Property @{
            'Source File' = $file
            'Data' = $sha256
            'Type' = 'SHA256'
        }))
    }

    $output = $output | Sort-Object 'Data' -Unique

    $Indicatordone = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $count = $output.Count
    Write-Host "Indicators extracted at $Indicatordone. Total count: $count" -ForegroundColor Cyan
    Log_Message -logfile $logfile -Message "Indicators extracted "
    $textboxResults.AppendText("Indicators extracted at $Indicatordone. Total count: $count `r`n")

    $htmlFile = New-Item -Path ".\Logs\Reports\$computerName\indicators.html" -ItemType File -Force

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
    padding: 0;
    margin: 0;
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
    z-index: 2;
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
        order: 1;
        align-self: center; 
        z-index: 1;
    }
    #indicators {
        margin-top: 60px;
        width: 60vw;
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
        order: 2;
        height: 325px;
        width: auto;
        font-size: 16px;
        margin-top: 10px;
        position: absolute;
        right: 0;
        top: 100%;
        z-index: 2;
    }
    #filter-button {
        height: 35px;
        font-size: 20px;
    }
    #reset-button {
        height: 35px;
        font-size: 20px;
    }
    #scroll-button {
        height: 35px;
        font-size: 20px;
    }
    .indicator-data {
        color: #66ccff;
        font-size: 18px;
        word-wrap: break-word;
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
    document.body.scrollTop = 0;
    document.documentElement.scrollTop = 0;
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
}

function Isolate-RemoteHost {
    param (
        [string]$computerName,
        [string]$logfile
    )

    if (![string]::IsNullOrEmpty($computerName)) {
        Invoke-Command -ComputerName $computerName -ScriptBlock {
            $localHostIPs = (Get-NetIPAddress -AddressFamily IPv4 -CimSession localhost).IPAddress | Where-Object { $_ -notlike "127.0.0.*" }
            $firewallProfiles = Get-NetFirewallProfile
            if ($firewallProfiles.Enabled -contains $false) {
                $firewallProfiles | Set-NetFirewallProfile -Enabled:True
            }
            $isolationRule = Get-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -ErrorAction SilentlyContinue
            if (!$isolationRule) {
                New-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -Direction Outbound -RemoteAddress $localHostIPs -Action Allow -Enabled:True
            } else {
                Set-NetFirewallRule -DisplayName "ISOLATION: Allowed Hosts" -RemoteAddress $localHostIPs
            }
            $firewallProfiles | Set-NetFirewallProfile -DefaultOutboundAction Block

            Write-Host "$computerName isolated at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
            Log_Message -logfile $logfile -Message "$computerName isolated"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid computer name.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Kill-RemoteProcess {
    param (
        [string]$computerName,
        [string]$selectedProcess,
        [string]$logfile
    )

    Invoke-Command -ComputerName $computerName -ScriptBlock {
        param ($selectedProcess)
        Stop-Process -Id $selectedProcess -Force
    } -ArgumentList $selectedProcess

    $killproc = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "Killed process $selectedProcess on $computerName at $killproc" -ForegroundColor Cyan 
    Log_Message -logfile $logfile -Message "Killed process $selectedProcess on $computerName"
    $textboxResults.AppendText("Killed process $selectedProcess on $computerName at $killproc")
}

function List-CopiedFiles {
    param (
        [string]$CopiedFilesPath,
        [string]$textboxResults,
        [System.Windows.Forms.ComboBox]$comboboxlocalFilePath
    )

    if (![string]::IsNullOrEmpty($CopiedFilesPath)) {
        try {
            $files = Get-ChildItem -Path $CopiedFilesPath -File -Recurse | Select-Object -ExpandProperty FullName
            $textboxResults.AppendText(($files -join "`r`n") + "`r`n")
            $comboboxlocalFilePath.Items.Clear()
            foreach ($file in $files) {
                $comboboxlocalFilePath.Items.Add($file)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error while trying to list files in the CopiedFiles directory: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid CopiedFiles path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function Log-Message {
    param (
        [string]$logfile,
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logfile -Value "Time=[$timestamp] User=[$Username] Message=[$Message]"
}
