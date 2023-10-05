# Check if the Active Directory module is available
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    # Active Directory module is not available

    # Create a Windows Forms message box
    [System.Windows.Forms.MessageBox]::Show(
        "The Active Directory module is not available. Please install the Remote Server Administration Tools (RSAT) or enable the Active Directory module.",
        "Active Directory Module Not Found",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    # Exit the script or take appropriate action
    Exit
}


# Load Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Computer Info"
$form.Width = 800
$form.Height = 600

# Create a label
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Size = New-Object System.Drawing.Size(200, 20)
$label.Text = "Select OS Version:"
$form.Controls.Add($label)

# Create a ComboBox for OS version
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10, 30)
$comboBox.Size = New-Object System.Drawing.Size(200, 20)
$comboBox.Items.AddRange(@(
    "Windows 10,11,2022,2019,2016 (10.0*)",     
    "Windows 8.1, 2012R2 (6.3*)",     
    "Windows 8,2012 (6.2)",       
    "Windows 7 (6.1)",  
    "Windows Server 2008 R2 (6.1)",   
    "Windows Server 2008, Vista (6.0)",   
    "Windows Server 2003, XP64 (5.2)",         
    "Windows XP32 (5.1)",    
    "Windows 2000 (5.0)"
))
$form.Controls.Add($comboBox)

# Create a DataGridView
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 60)
$dataGridView.Size = New-Object System.Drawing.Size(760, 400)
$dataGridView.AllowUserToAddRows = $false
$form.Controls.Add($dataGridView)

# Create columns for the DataGridView
$dataGridView.Columns.Add("ComputerName", "ComputerName")
$dataGridView.Columns.Add("ResolvesToIP", "ResolvesToIP")
$dataGridView.Columns.Add("IPResolvesTo", "IPResolvesTo")
$dataGridView.Columns.Add("PingStatus", "PingStatus")
$dataGridView.Columns.Add("OSName", "OSName")
$dataGridView.Columns.Add("OU", "OU")

# Create a "Run" button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Location = New-Object System.Drawing.Point(225, 30)
$runButton.Size = New-Object System.Drawing.Size(75, 20)
$runButton.Text = "Run"
$runButton.Add_Click({
    $selectedItem = $comboBox.SelectedItem.ToString()
    # Extract the version number from the selected item (text between the parentheses)
    $version = [regex]::Match($selectedItem, '\((.*?)\)').Groups[1].Value
    Execute-Script -version $version
})
$form.Controls.Add($runButton)

# Create a ProgressBar to show script progress
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 470)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$form.Controls.Add($progressBar)

# Create a label for the CSV file link
$labelCSV = New-Object System.Windows.Forms.LinkLabel
$labelCSV.Location = New-Object System.Drawing.Point(10, 500)
$labelCSV.Size = New-Object System.Drawing.Size(760, 20)
$labelCSV.Visible = $false  # Initially hide the link
$form.Controls.Add($labelCSV)

# Create a "Close" button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(10, 530)
$closeButton.Size = New-Object System.Drawing.Size(75, 20)
$closeButton.Text = "Close"
$closeButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($closeButton)

# Create a "Reset" button
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(95, 530)
$resetButton.Size = New-Object System.Drawing.Size(75, 20)
$resetButton.Text = "Reset"
$resetButton.Add_Click({
    $dataGridView.Rows.Clear()
    $labelCSV.Visible = $false
    $progressBar.Value = 0
})
$form.Controls.Add($resetButton)

# Function to execute the script
function Execute-Script {
    param (
        [string]$version
    )

    # Import Active Directory and set up the path while suppressing errors
    New-Item -ItemType Directory -Path C:\BGInfo -ErrorAction SilentlyContinue
    $ErrorActionPreference = 'SilentlyContinue'
    Import-Module ActiveDirectory

    # Gather versions of operating systems from AD
    $versionWildcard = "$version*"
    $computers = Get-ADComputer -Filter * -Property OperatingSystemVersion, Name, DistinguishedName, OperatingSystem | 
        Where-Object { $_.OperatingSystemVersion -like $versionWildcard } | 
        Select-Object Name, OperatingSystem, DistinguishedName 

    # Remove spacing from file and create clean
    $computers | ForEach-Object { $_.OperatingSystem = $_.OperatingSystem.Trim() }
    $computers | Export-Csv C:\BGInfo\computers.csv -NoTypeInformation

    # Clear the DataGridView
    $dataGridView.Rows.Clear()

    # Take the file and scan each hostname to see if it is active currently
    Update-Status "Starting HOST Network Test for Windows version $version"

    $processedComputers = @{}  # Track processed computers

    $complist = Get-Content "C:\BGInfo\computers.csv" | ConvertFrom-Csv
    $counter = 1

    $results = @()  # Collect results here

    foreach ($compInfo in $complist) {
        $compName = $compInfo.Name

        # Check if the computer is already processed
        if ($processedComputers.ContainsKey($compName)) {
            continue
        }

        $compDistinguishedName = $compInfo.DistinguishedName

        $pingtest = Test-Connection -ComputerName $compName -Quiet -Count 1 -ErrorAction SilentlyContinue
        $pingStatus = $null

        if ($pingtest) {
            $pingStatus = "Host Online"
            $pingStatusColor = [System.Drawing.Color]::Green
        } else {
            $pingStatus = "Not Reachable"
            $pingStatusColor = [System.Drawing.Color]::Red
        }

        $osName = $compInfo.OperatingSystem
        $ou = $compInfo.DistinguishedName -replace '^CN=.*?,OU=|,DC=.*$', ''

        $result = [PSCustomObject]@{
            ComputerName = $compName
            ResolvesToIP = if ($pingtest) { [System.Net.Dns]::GetHostAddresses($compName) -join "," } else { "" }
            IPResolvesTo = if ($pingtest) { ([System.Net.Dns]::GetHostEntry($compName)).HostName } else { "" }
            PingStatus = $pingStatus
            OSName = $osName
            OU = $ou
        }

        $results += $result

        $dataGridView.Rows.Add(
            $compName,
            $result.ResolvesToIP,
            $result.IPResolvesTo,
            $result.PingStatus,
            $result.OSName,
            $result.OU
        )

        $processedComputers[$compName] = $true

        $cell = $dataGridView.Rows[$counter - 1].Cells[3]
        $cell.Style.BackColor = $pingStatusColor

        $counter++
        $progressBar.Value = [Math]::Min(($counter / $complist.Count) * 100, 100)
        $form.Refresh()
    }

    # Force DataGridView refresh
    $dataGridView.Refresh()

    # This script takes the results and gives you the IP and its status in a CSV
    $dnsResults = "C:\BGInfo\host-ip.csv"
    $results | Export-Csv $dnsResults -NoTypeInformation

    # Set the Tag property of the label to store the file path
    $labelCSV.Tag = $dnsResults

    # Update the label link to open the CSV file
    $labelCSV.Text = "Click here to open the CSV file: C:\BGInfo\host-ip.csv"
    $labelCSV.Visible = $true
    $labelCSV.Add_LinkClicked({
        Start-Process -FilePath $labelCSV.Tag
    })

    Update-Status "Script execution completed."

    # Show a dialog box to indicate completion
    [System.Windows.Forms.MessageBox]::Show("Search is complete!", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# Function to update the TextBox with script execution progress
function Update-Status {
    param (
        [string]$text
    )

    $textBox.AppendText("$text`r`n")
}

# Create a TextBox to display the script text
$scriptTextBox = New-Object System.Windows.Forms.TextBox
$scriptTextBox.Location = New-Object System.Drawing.Point(10, 560)
$scriptTextBox.Size = New-Object System.Drawing.Size(760, 20)
$scriptTextBox.Multiline = $true
$form.Controls.Add($scriptTextBox)

# Load and display the script text
$scriptText = Get-Content -Path $MyInvocation.MyCommand.Path
$scriptTextBox.Text = $scriptText

# Show the form
$form.ShowDialog()
