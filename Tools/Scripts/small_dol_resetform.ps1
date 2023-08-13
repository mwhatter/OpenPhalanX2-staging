param (
        $ScriptStartTime
    )

function Reset-Form {
    $statusLabel.Text = ""
    $textboxURL.Clear()
    $textboxResults.Clear()
    $textboxResults.Text = "Session started at $ScriptStartTime `r`n"
    $buttonSubmitUrl.Enabled = $true
    $buttonRetrieveReport.Enabled = $false
    $buttonSubmitUrl.Enabled = $true
    $buttonRetrieveReportfile.Enabled = $false
    $textboxaddargs.Text = ""
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
    $comboboxUsername.Items.Clear()
    $comboboxUsername.SelectedIndex = -1 
    $comboboxUsername.Text = ""
}

Reset-Form