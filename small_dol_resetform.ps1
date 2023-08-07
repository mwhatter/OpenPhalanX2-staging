function Reset-Form {
    $textboxURL.Clear()
    $textboxResults.Clear()
    $buttonSubmitUrl.Enabled = $true
    $buttonCheckStatus.Enabled = $false
    $buttonCheckStatus.Text = "Awaiting Job"
    $buttonRetrieveReport.Enabled = $false
    $buttonSubmitUrl.Enabled = $true
    $buttonCheckStatusfile.Enabled = $false
    $buttonCheckStatusfile.Text = "Awaiting Job"
    $buttonRetrieveReportfile.Enabled = $false
    $textboxaddargs.Clear()
    $comboboxlocalFilePath.Items.Clear()
    $comboboxlocalFilePath.SelectedIndex = -1 
    $comboboxlocalFilePath.Text = ""
    $textboxremoteFilePath.Clear()
    $comboboxComputerName.SelectedIndex = -1 
    $comboboxComputerName.Text = ""
    $dropdownProcessId.Items.Clear()
    $dropdownProcessId.SelectedIndex = -1 
    $dropdownProcessId.Text = ""
    $script:SubmissionId = ""
    $script:SubmissionId = ""
    $buttonCheckStatus.ForeColor = [System.Drawing.Color]::lightseagreen   
    $buttonCheckStatus.BackColor = [System.Drawing.Color]::Black 
    $buttonCheckStatusfile.ForeColor = [System.Drawing.Color]::lightseagreen   
    $buttonCheckStatusfile.BackColor = [System.Drawing.Color]::Black 
}

