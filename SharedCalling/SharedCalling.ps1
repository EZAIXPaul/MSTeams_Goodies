# Import necessary assemblies for the GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Teams Shared Calling Management Tool"
$form.Size = New-Object System.Drawing.Size(900, 750)
$form.StartPosition = "CenterScreen"

# Create Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$form.Controls.Add($tabControl)

# Create Tabs
$tabViewPolicies = New-Object System.Windows.Forms.TabPage
$tabViewPolicies.Text = "View Policies"
$tabControl.TabPages.Add($tabViewPolicies)

$tabCreatePolicy = New-Object System.Windows.Forms.TabPage
$tabCreatePolicy.Text = "Create Policy"
$tabControl.TabPages.Add($tabCreatePolicy)

$tabAssignPolicy = New-Object System.Windows.Forms.TabPage
$tabAssignPolicy.Text = "Assign Policy"
$tabControl.TabPages.Add($tabAssignPolicy)

$tabViewUsers = New-Object System.Windows.Forms.TabPage
$tabViewUsers.Text = "View Users"
$tabControl.TabPages.Add($tabViewUsers)

# =====================================
# Tab 1: View Policies
# =====================================

# Button to load shared calling policies
$buttonLoadShared = New-Object System.Windows.Forms.Button
$buttonLoadShared.Text = "Load Shared Calling Policies"
$buttonLoadShared.Location = New-Object System.Drawing.Point(10,10)
$buttonLoadShared.Size = New-Object System.Drawing.Size(200,30)
$tabViewPolicies.Controls.Add($buttonLoadShared)

# Button to export policies to CSV
$buttonExportPolicies = New-Object System.Windows.Forms.Button
$buttonExportPolicies.Text = "Export to CSV"
$buttonExportPolicies.Location = New-Object System.Drawing.Point(220,10)
$buttonExportPolicies.Size = New-Object System.Drawing.Size(100,30)
$tabViewPolicies.Controls.Add($buttonExportPolicies)

# ListView to display shared calling policies
$listViewPolicies = New-Object System.Windows.Forms.ListView
$listViewPolicies.Location = New-Object System.Drawing.Point(10,50)
$listViewPolicies.Size = New-Object System.Drawing.Size(860, 620)
$listViewPolicies.View = 'Details'
$listViewPolicies.FullRowSelect = $true
$listViewPolicies.GridLines = $true
$listViewPolicies.Columns.Add("Policy Name", 150)
$listViewPolicies.Columns.Add("Resource Account", 200)
$listViewPolicies.Columns.Add("Phone Number", 150)
$listViewPolicies.Columns.Add("Emergency Numbers", 300)
$tabViewPolicies.Controls.Add($listViewPolicies)

# =====================================
# Tab 2: Create Policy
# =====================================

# Label for Shared Calling Policy Name
$labelSCPolicyName = New-Object System.Windows.Forms.Label
$labelSCPolicyName.Text = "Shared Calling Policy Name:"
$labelSCPolicyName.Location = New-Object System.Drawing.Point(10,20)
$labelSCPolicyName.Size = New-Object System.Drawing.Size(200,20)
$tabCreatePolicy.Controls.Add($labelSCPolicyName)

# Textbox for Shared Calling Policy Name
$textSCPolicyName = New-Object System.Windows.Forms.TextBox
$textSCPolicyName.Location = New-Object System.Drawing.Point(210,20)
$textSCPolicyName.Size = New-Object System.Drawing.Size(550,20)
$tabCreatePolicy.Controls.Add($textSCPolicyName)

# Label for Resource Account UPN
$labelResourceAccount = New-Object System.Windows.Forms.Label
$labelResourceAccount.Text = "Resource Account (Auto Attendant) UPN:"
$labelResourceAccount.Location = New-Object System.Drawing.Point(10,60)
$labelResourceAccount.Size = New-Object System.Drawing.Size(250,20)
$tabCreatePolicy.Controls.Add($labelResourceAccount)

# Textbox for Resource Account UPN
$textResourceAccount = New-Object System.Windows.Forms.TextBox
$textResourceAccount.Location = New-Object System.Drawing.Point(260,60)
$textResourceAccount.Size = New-Object System.Drawing.Size(500,20)
$tabCreatePolicy.Controls.Add($textResourceAccount)

# Label for Emergency Dial Strings
$labelEmergencyDialStrings = New-Object System.Windows.Forms.Label
$labelEmergencyDialStrings.Text = "Emergency Dial Strings (comma-separated):"
$labelEmergencyDialStrings.Location = New-Object System.Drawing.Point(10,100)
$labelEmergencyDialStrings.Size = New-Object System.Drawing.Size(300,20)
$tabCreatePolicy.Controls.Add($labelEmergencyDialStrings)

# Textbox for Emergency Dial Strings
$textEmergencyDialStrings = New-Object System.Windows.Forms.TextBox
$textEmergencyDialStrings.Location = New-Object System.Drawing.Point(310,100)
$textEmergencyDialStrings.Size = New-Object System.Drawing.Size(450,20)
$tabCreatePolicy.Controls.Add($textEmergencyDialStrings)

# Label for Emergency Callback Numbers (ECBNs)
$labelEmergencyNumbers = New-Object System.Windows.Forms.Label
$labelEmergencyNumbers.Text = "Emergency Callback Numbers (ECBNs, comma-separated):"
$labelEmergencyNumbers.Location = New-Object System.Drawing.Point(10,140)
$labelEmergencyNumbers.Size = New-Object System.Drawing.Size(350,20)
$tabCreatePolicy.Controls.Add($labelEmergencyNumbers)

# Textbox for Emergency Callback Numbers
$textEmergencyNumbers = New-Object System.Windows.Forms.TextBox
$textEmergencyNumbers.Location = New-Object System.Drawing.Point(360,140)
$textEmergencyNumbers.Size = New-Object System.Drawing.Size(400,20)
$tabCreatePolicy.Controls.Add($textEmergencyNumbers)

# Checkbox for Allow Enhanced Emergency Services
$checkAllowE911 = New-Object System.Windows.Forms.CheckBox
$checkAllowE911.Text = "Allow Enhanced Emergency Services"
$checkAllowE911.Location = New-Object System.Drawing.Point(10,180)
$checkAllowE911.Size = New-Object System.Drawing.Size(250,20)
$tabCreatePolicy.Controls.Add($checkAllowE911)

# Button to create Shared Calling Policy
$buttonCreateSCPolicy = New-Object System.Windows.Forms.Button
$buttonCreateSCPolicy.Text = "Create Shared Calling Policy"
$buttonCreateSCPolicy.Location = New-Object System.Drawing.Point(10,220)
$buttonCreateSCPolicy.Size = New-Object System.Drawing.Size(200,30)
$tabCreatePolicy.Controls.Add($buttonCreateSCPolicy)

# =====================================
# Tab 3: Assign Policy
# =====================================

# Label for Policy Selection
$labelSelectPolicy = New-Object System.Windows.Forms.Label
$labelSelectPolicy.Text = "Select Shared Calling Policy:"
$labelSelectPolicy.Location = New-Object System.Drawing.Point(10,20)
$labelSelectPolicy.Size = New-Object System.Drawing.Size(200,20)
$tabAssignPolicy.Controls.Add($labelSelectPolicy)

# ComboBox for Shared Calling Policies
$comboPolicies = New-Object System.Windows.Forms.ComboBox
$comboPolicies.Location = New-Object System.Drawing.Point(210,20)
$comboPolicies.Size = New-Object System.Drawing.Size(550,20)
$tabAssignPolicy.Controls.Add($comboPolicies)

# Button to load shared calling policies
$buttonLoadPoliciesAssign = New-Object System.Windows.Forms.Button
$buttonLoadPoliciesAssign.Text = "Load Policies"
$buttonLoadPoliciesAssign.Location = New-Object System.Drawing.Point(10,60)
$buttonLoadPoliciesAssign.Size = New-Object System.Drawing.Size(100,30)
$tabAssignPolicy.Controls.Add($buttonLoadPoliciesAssign)

# Label for User Principal Name
$labelUserUPN = New-Object System.Windows.Forms.Label
$labelUserUPN.Text = "User Principal Name (email):"
$labelUserUPN.Location = New-Object System.Drawing.Point(10,110)
$labelUserUPN.Size = New-Object System.Drawing.Size(200,20)
$tabAssignPolicy.Controls.Add($labelUserUPN)

# Textbox for User Principal Name
$textUserUPN = New-Object System.Windows.Forms.TextBox
$textUserUPN.Location = New-Object System.Drawing.Point(210,110)
$textUserUPN.Size = New-Object System.Drawing.Size(550,20)
$tabAssignPolicy.Controls.Add($textUserUPN)

# Button to assign policy to single user
$buttonAssignPolicySingle = New-Object System.Windows.Forms.Button
$buttonAssignPolicySingle.Text = "Assign Policy to User"
$buttonAssignPolicySingle.Location = New-Object System.Drawing.Point(10,150)
$buttonAssignPolicySingle.Size = New-Object System.Drawing.Size(150,30)
$tabAssignPolicy.Controls.Add($buttonAssignPolicySingle)

# Label for CSV Upload
$labelCSV = New-Object System.Windows.Forms.Label
$labelCSV.Text = "Upload CSV File (with UPN column):"
$labelCSV.Location = New-Object System.Drawing.Point(10,200)
$labelCSV.Size = New-Object System.Drawing.Size(200,20)
$tabAssignPolicy.Controls.Add($labelCSV)

# Textbox for CSV File Path
$textCSV = New-Object System.Windows.Forms.TextBox
$textCSV.Location = New-Object System.Drawing.Point(210,200)
$textCSV.Size = New-Object System.Drawing.Size(450,20)
$tabAssignPolicy.Controls.Add($textCSV)

# Button to browse CSV File
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse"
$buttonBrowse.Location = New-Object System.Drawing.Point(670,195)
$buttonBrowse.Size = New-Object System.Drawing.Size(90,30)
$tabAssignPolicy.Controls.Add($buttonBrowse)

# Button to assign policy to multiple users
$buttonAssignPolicyCSV = New-Object System.Windows.Forms.Button
$buttonAssignPolicyCSV.Text = "Assign Policy to CSV Users"
$buttonAssignPolicyCSV.Location = New-Object System.Drawing.Point(10,240)
$buttonAssignPolicyCSV.Size = New-Object System.Drawing.Size(200,30)
$tabAssignPolicy.Controls.Add($buttonAssignPolicyCSV)

# =====================================
# Tab 4: View Users
# =====================================

# Button to load users with Shared Calling Policy
$buttonLoadUsers = New-Object System.Windows.Forms.Button
$buttonLoadUsers.Text = "Load Users with Shared Calling Policy"
$buttonLoadUsers.Location = New-Object System.Drawing.Point(10,10)
$buttonLoadUsers.Size = New-Object System.Drawing.Size(250,30)
$tabViewUsers.Controls.Add($buttonLoadUsers)

# Button to export users to CSV
$buttonExportUsers = New-Object System.Windows.Forms.Button
$buttonExportUsers.Text = "Export to CSV"
$buttonExportUsers.Location = New-Object System.Drawing.Point(270,10)
$buttonExportUsers.Size = New-Object System.Drawing.Size(100,30)
$tabViewUsers.Controls.Add($buttonExportUsers)

# ListView to display users
$listViewUsers = New-Object System.Windows.Forms.ListView
$listViewUsers.Location = New-Object System.Drawing.Point(10,50)
$listViewUsers.Size = New-Object System.Drawing.Size(860, 620)
$listViewUsers.View = 'Details'
$listViewUsers.FullRowSelect = $true
$listViewUsers.GridLines = $true
$listViewUsers.Columns.Add("User Principal Name", 250)
$listViewUsers.Columns.Add("Shared Calling Policy", 250)
$listViewUsers.Columns.Add("Voice Routing Policy", 250)
$tabViewUsers.Controls.Add($listViewUsers)

# =====================================
# Event Handlers and Functions
# =====================================

# Global variables to store data for export
$global:PoliciesData = @()
$global:UsersData = @()

# Load Shared Calling Policies (View Policies Tab)
$buttonLoadShared.Add_Click({
    try {
        $sharedPolicies = Get-CsTeamsSharedCallingRoutingPolicy
        $listViewPolicies.Items.Clear()
        $global:PoliciesData = @() # Clear previous data
        foreach ($policy in $sharedPolicies) {
            $policyName = $policy.Identity

            # Handle ResourceAccount
            $resourceAccountId = $policy.ResourceAccount
            if ($resourceAccountId -ne $null) {
                # Get the resource account details
                $resourceAccount = Get-CsOnlineApplicationInstance -Identity $resourceAccountId
                if ($resourceAccount -ne $null) {
                    $resourceAccountName = $resourceAccount.DisplayName
                    $resourceAccountUPN = $resourceAccount.UserPrincipalName

                    # Retrieve the phone number assigned to the resource account using Get-CsPhoneNumberAssignment
                    try {
                        $phoneNumberAssignment = Get-CsPhoneNumberAssignment -AssignedPstnTargetId $resourceAccount.ObjectId
                        if ($phoneNumberAssignment -and $phoneNumberAssignment.TelephoneNumber) {
                            $phoneNumber = $phoneNumberAssignment.TelephoneNumber
                        } else {
                            $phoneNumber = ""
                        }
                    } catch {
                        $phoneNumber = ""
                    }
                } else {
                    $resourceAccountName = "Unknown Resource Account"
                    $phoneNumber = ""
                }
            } else {
                $resourceAccountName = "No Resource Account Assigned"
                $phoneNumber = ""
            }

            # Handle EmergencyNumbers
            $emergencyNumbers = $policy.EmergencyNumbers
            if ($emergencyNumbers -ne $null) {
                $emergencyNumbersText = $emergencyNumbers -join ", "
            } else {
                $emergencyNumbersText = ""
            }

            $listItem = New-Object System.Windows.Forms.ListViewItem($policyName)
            $listItem.SubItems.Add($resourceAccountName)
            $listItem.SubItems.Add($phoneNumber)
            $listItem.SubItems.Add($emergencyNumbersText)
            $listViewPolicies.Items.Add($listItem)

            # Add to global data for export
            $global:PoliciesData += [PSCustomObject]@{
                'Policy Name'       = $policyName
                'Resource Account'  = $resourceAccountName
                'Phone Number'      = $phoneNumber
                'Emergency Numbers' = $emergencyNumbersText
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Shared Calling Policies Loaded!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading shared calling policies: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Export Policies to CSV function
$buttonExportPolicies.Add_Click({
    if ($global:PoliciesData.Count -gt 0) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save Policies Data as CSV"
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $saveFileDialog.ShowDialog() | Out-Null
        $csvPath = $saveFileDialog.FileName
        if ($csvPath) {
            try {
                $global:PoliciesData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $csvPath", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please load the policies first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Create Shared Calling Policy function
$buttonCreateSCPolicy.Add_Click({
    $policyName = $textSCPolicyName.Text.Trim()
    $resourceAccountUPN = $textResourceAccount.Text.Trim()
    $emergencyDialStringsText = $textEmergencyDialStrings.Text.Trim()
    $emergencyNumbersText = $textEmergencyNumbers.Text.Trim()
    $allowE911 = $checkAllowE911.Checked

    if ($policyName -and $resourceAccountUPN -and $emergencyDialStringsText) {
        try {
            # Get the resource account
            $resourceAccount = Get-CsOnlineApplicationInstance -Identity $resourceAccountUPN
            if ($null -ne $resourceAccount) {
                # Process emergency numbers
                $emergencyDialStrings = $emergencyDialStringsText.Split(",") | ForEach-Object { $_.Trim() }
                $emergencyNumbers = @()
                foreach ($dialString in $emergencyDialStrings) {
                    if ($emergencyNumbersText) {
                        $ecbnList = $emergencyNumbersText.Split(",") | ForEach-Object { $_.Trim() }
                        $emergencyNumber = New-CsTeamsEmergencyNumber -EmergencyDialString $dialString -EmergencyCallbackNumber $ecbnList[0]
                    } else {
                        $emergencyNumber = New-CsTeamsEmergencyNumber -EmergencyDialString $dialString
                    }
                    $emergencyNumbers += $emergencyNumber
                }

                # Create Emergency Call Routing Policy
                $ecrpName = "$policyName-ECRP"
                New-CsTeamsEmergencyCallRoutingPolicy -Identity $ecrpName -EmergencyNumbers @{add=$emergencyNumbers} -AllowEnhancedEmergencyServices $allowE911

                # Create the Shared Calling Policy
                if ($emergencyNumbersText) {
                    $ecbnArray = $emergencyNumbersText.Split(",") | ForEach-Object { $_.Trim() }
                    New-CsTeamsSharedCallingRoutingPolicy -Identity $policyName -ResourceAccount $resourceAccount.Identity -EmergencyNumbers @{add=$ecbnArray}
                } else {
                    New-CsTeamsSharedCallingRoutingPolicy -Identity $policyName -ResourceAccount $resourceAccount.Identity
                }

                [System.Windows.Forms.MessageBox]::Show("Shared Calling Policy '$policyName' created successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                [System.Windows.Forms.MessageBox]::Show("Resource Account '$resourceAccountUPN' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating shared calling policy: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter the Policy Name, Resource Account UPN, and Emergency Dial Strings.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Load Shared Calling Policies (Assign Policy Tab)
$buttonLoadPoliciesAssign.Add_Click({
    try {
        $sharedPolicies = Get-CsTeamsSharedCallingRoutingPolicy
        $comboPolicies.Items.Clear()
        foreach ($policy in $sharedPolicies) {
            $comboPolicies.Items.Add($policy.Identity)
        }
        [System.Windows.Forms.MessageBox]::Show("Shared Calling Policies Loaded!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading shared calling policies: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Assign Shared Calling Policy to single user function
$buttonAssignPolicySingle.Add_Click({
    $policyName = $comboPolicies.SelectedItem
    $userPrincipalName = $textUserUPN.Text.Trim()
    if ($policyName -and $userPrincipalName) {
        try {
            # Assign Shared Calling Policy
            Grant-CsTeamsSharedCallingRoutingPolicy -Identity $userPrincipalName -PolicyName $policyName

            # Assign Emergency Call Routing Policy
            $ecrpName = "$policyName-ECRP"
            Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $userPrincipalName -PolicyName $ecrpName

            [System.Windows.Forms.MessageBox]::Show("Shared Calling Policy '$policyName' and Emergency Call Routing Policy '$ecrpName' assigned to user '$userPrincipalName' successfully!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error assigning policies: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a policy and enter the User Principal Name (email).", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Browse CSV File function
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $result = $openFileDialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $textCSV.Text = $openFileDialog.FileName
    }
})

# Assign Shared Calling Policy to multiple users via CSV function
$buttonAssignPolicyCSV.Add_Click({
    $policyName = $comboPolicies.SelectedItem
    $csvFilePath = $textCSV.Text
    if ($policyName) {
        if (Test-Path $csvFilePath) {
            try {
                $ecrpName = "$policyName-ECRP"
                $users = Import-Csv -Path $csvFilePath
                foreach ($user in $users) {
                    $userPrincipalName = $user.UPN
                    if ($userPrincipalName) {
                        # Assign Shared Calling Policy
                        Grant-CsTeamsSharedCallingRoutingPolicy -Identity $userPrincipalName -PolicyName $policyName
                        # Assign Emergency Call Routing Policy
                        Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $userPrincipalName -PolicyName $ecrpName
                    }
                }
                [System.Windows.Forms.MessageBox]::Show("Shared Calling Policy '$policyName' and Emergency Call Routing Policy '$ecrpName' assigned to users from CSV!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error assigning policies: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV file.", "Invalid File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a policy to assign.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Load Users with Shared Calling Policy
$buttonLoadUsers.Add_Click({
    try {
        $listViewUsers.Items.Clear()
        $global:UsersData = @() # Clear previous data
        $users = Get-CsOnlineUser -Filter {TeamsSharedCallingRoutingPolicy -ne $null}
        foreach ($user in $users) {
            $userUPN = $user.UserPrincipalName

            # Handle SharedCallingPolicy
            $sharedCallingPolicy = $user.TeamsSharedCallingRoutingPolicy
            if ($sharedCallingPolicy -ne $null) {
                $sharedCallingPolicyName = [string]$sharedCallingPolicy
            } else {
                $sharedCallingPolicyName = ""
            }

            # Handle VoiceRoutingPolicy
            $voiceRoutingPolicy = $user.VoiceRoutingPolicy
            if ($voiceRoutingPolicy -ne $null) {
                $voiceRoutingPolicyName = [string]$voiceRoutingPolicy
            } else {
                $voiceRoutingPolicyName = ""
            }

            $listItem = New-Object System.Windows.Forms.ListViewItem($userUPN)
            $listItem.SubItems.Add($sharedCallingPolicyName)
            $listItem.SubItems.Add($voiceRoutingPolicyName)
            $listViewUsers.Items.Add($listItem)

            # Add to global data for export
            $global:UsersData += [PSCustomObject]@{
                'User Principal Name'    = $userUPN
                'Shared Calling Policy'  = $sharedCallingPolicyName
                'Voice Routing Policy'   = $voiceRoutingPolicyName
            }
        }
        [System.Windows.Forms.MessageBox]::Show("Users with Shared Calling Policy Loaded!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading users: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Export Users to CSV function
$buttonExportUsers.Add_Click({
    if ($global:UsersData.Count -gt 0) {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveFileDialog.Title = "Save Users Data as CSV"
        $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $saveFileDialog.ShowDialog() | Out-Null
        $csvPath = $saveFileDialog.FileName
        if ($csvPath) {
            try {
                $global:UsersData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $csvPath", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No data to export. Please load the users first.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

# Show the form
[void]$form.ShowDialog()
