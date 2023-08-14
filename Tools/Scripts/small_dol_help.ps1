function Show-ColoredMessage {
    param([string]$message, [string]$title)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(980, 600)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = [System.Drawing.Color]::Black
    $form.ForeColor = [System.Drawing.Color]::lightseagreen

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(950, 500) 
    $panel.Location = New-Object System.Drawing.Point(10, 10)
    $panel.AutoScroll = $true

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.ForeColor = [System.Drawing.Color]::LightSeaGreen 
    $label.Font = New-Object System.Drawing.Font("Arial", 12) # Increase font size to 12

    # Create the Update button
    $updateButton = New-Object System.Windows.Forms.Button
    $updateButton.Text = "Update Phalanx"
    $updateButton.Location = New-Object System.Drawing.Point(350, 520) # Adjust the location to be below the panel
    $updateButton.Size = New-Object System.Drawing.Size(200, 30)
    $updateButton.Add_Click({
        # Execute the update script when the button is clicked
        & ".\Update_Phalanx.ps1"
    })

    $panel.Controls.Add($label)
    $form.Controls.Add($panel)
    $form.Controls.Add($updateButton) # Add the button to the form

    $form.ShowDialog()
}


$helpText = @"
    Enter Remote Computer Name or Get Host List:

    Enter the name of the remote computer you want to interact with in the provided text box, or click on 
    "Get Host List" to retrieve a list of all hosts on the domain. The list will paginate and match characters 
    as you type. The selected or typed hostname will be the target for any remote actions.


    View Processes Button:

    Clicking this button will display the remote computer's process list in the lower display box. It also 
    populates the dropdown for selecting processes to terminate. The list includes the process ID, name, 
    and status.


    Copy All Process Binaries Button:

    This button will copy every uniquely pathed binary that is currently running a process on the remote machine. 
    The binaries are stored in the "OpenPhalanx\CopiedFiles" directory on your local host. The contents of this 
    directory are displayed when "View Copied Files" is clicked.


    Tractor Beam:

    Engaging the "Tractor Beam" will assign firewall rules to the remote host which prevent any IP addresses 
    other than those of the local host from communicating with the remote host. This is useful for isolating 
    a compromised system.


    Hunt File Button:

    This button will attempt to locate the file specified in the "Remote File Path" text box. If the file is not 
    found in the specified remote file path the recycle bin is queried. The status of the volume shadow copy 
    service is checked. The existence of system restore points and backups is also checked.


    Place File Button:

    This button takes the file specified in the "Local File Path" dropdown and places it at the location 
    specified in the "Remote File Path" text box.


    Place and Run Button:

    This button will place the file specified in the "Local File Path" dropdown at the location specified in the 
    "Remote File Path" text box and then execute it. If the "Arguments" text box is filled, those arguments will 
    be used when executing the file.


    Execute Oneliner Button:

    This button will attempt to execute the command specified in the "Arguments" text box on the remote machine.


    Deploy Sysmon Button:

    This button deploys Sysmon, a powerful system monitoring tool, on the remote host using the Olaf Hartong 
    default configuration.


    RapidTriage Button:

    This button runs the RapidTriage tool, which collects a wide range of data from the remote host and outputs it 
    in an xlsx workbook. The workbook includes several worksheets with different types of data, such as running 
    processes, network connections, and more.


    View Copied Files Button:

    This button displays the contents of the "OpenPhalanx\CopiedFiles" directory on your local host in the lower 
    display box. The "Local File Path" dropdown is populated with the copied files.


    Reset Button:

    This button clears all input fields and resets the state of the form. It also disables the "Retrieve Report" 
    and "Check Status" buttons until a new job is submitted.


    Submit URL Button:

    This button submits the URL specified in the "URL" text box to the VMRay sandbox for analysis. After 
    submission, the "Check Status" button is enabled.


    Retrieve Report Button:

    This button retrieves the report of the analysis of the submitted URL. The report includes detailed information 
    about the URL, such as its threat score, threat level, and more.


    Sandbox Local File Button:

    This button submits the file specified in the "Local File Path" dropdown to the VMRay sandbox for analysis. 
    After submission, the "Check Status" button is enabled.


    Retrieve Report (File) Button:

    This button retrieves the report of the analysis of the submitted file. The report includes detailed information 
    about the file, such as its threat score, threat level, and more.


    WinEventalyzer Button:

    This button runs the WinEventalyzer tool, which collects Windows event logs from the remote host and analyzes them 
    using several threat hunting tools, including DeepBlueCLI, and Hayabusa. The output includes timelines, 
    summaries, and metrics that can be used for threat hunting.

    
    USN Journal Button:

    This button retrieves the USN Journal from the remote host. The USN Journal is a log of all changes to files on the 
    remote host and can be used for forensic analysis.


    List Copied Files Button:

    This button lists all the files that have been copied from the remote host to the "OpenPhalanx\CopiedFiles" directory 
    on your local host. The list is displayed in the lower display box, and the "Local File Path" dropdown is populated 
    with the copied files.


    Select Remote File Button:

    This button opens a custom remote file system explorer. You can navigate through the file system of the remote computer 
    and select a file. The selected file's path will be displayed in the "Remote File Path" text box.


    Select Local File Button:

    This button opens a file dialog that allows you to select a file from your local machine. The selected file's path 
    will be displayed in the "Local File Path" text box.


    Force Password Change Button:

    This button forces the specified user to change their password at the next logon. This is useful for ensuring that 
    users regularly update their passwords.


    Log Off User Button:

    This button forces the specified user to log off from the remote computer. This can be useful for ending a user's 
    session without shutting down the computer.


    Disable Account Button:

    This button disables the specified user's account on the remote computer. This can be useful for preventing a user 
    from logging in to the computer.


    Enable Account Button:

    This button enables the specified user's account on the remote computer. This can be useful for allowing a user 
    who was previously disabled to log in to the computer.


    Retrieve System Info Button:

    This button retrieves Active Directory information about the specified remote host and all users who have logged into 
    this host. This information is displayed in the lower display box, and the "Username" dropdown is populated with the 
    usernames.


    Restart Button:

    This button sends a command to the remote computer to restart. This can be useful for applying updates or changes 
    that require a restart.


    Kill Process Button:

    This button kills the selected process on the remote computer. This can be useful for stopping a process that is 
    not responding or that is using too many resources.


    Shutdown Button:

    This button sends a command to the remote computer to shut down. This can be useful for turning off the computer 
    remotely.


    Copy File Button:

    This button copies the file specified in the "Remote File Path" text box from the remote computer to the 
    "OpenPhalanx\CopiedFiles" directory on your local host.


    Delete File Button:

    This button deletes the file specified in the "Remote File Path" text box from the remote computer. This can be 
    useful for removing unwanted or unnecessary files from the computer.


    Intelligizer Button:

    This button scrapes indicators from collected reports and logs for further analysis. This can be useful for 
    identifying patterns or anomalies in the data.


    BoxEmAll Button:

    This button submits all files in the "OpenPhalanx\CopiedFiles" directory to the VMRay sandbox for analysis. This 
    can be useful for analyzing multiple files at once.


    Get Intel Button:

    This button retrieves intelligence on the specified indicator. This can be useful for getting more information 
    about a potential threat.


    Undo Isolation Button:

    This button removes the firewall rules that were applied for isolation. This can be useful for restoring 
    communication with the remote host after it has been isolated.


"@

Show-ColoredMessage -message "$helpText" -title "Help - Defending Off the Land"
