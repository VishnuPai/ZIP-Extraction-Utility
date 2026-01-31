Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================================
# GLOBAL STATE
# =========================================
$script:IsProcessing = $false
$script:CancelRequested = $false
$LogFolder = "C:\ZipUtilityLogs"

if (!(Test-Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null
}

# =========================================
# FORM
# =========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "ZIP Extraction Utility"
$form.Size = New-Object System.Drawing.Size(850,620)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI",9)

# =========================================
# SOURCE
# =========================================
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Folder:"
$lblSource.Location = New-Object System.Drawing.Point(20,20)
$lblSource.AutoSize = $true
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Size = New-Object System.Drawing.Size(650,25)
$txtSource.Location = New-Object System.Drawing.Point(20,45)
$form.Controls.Add($txtSource)

$btnSource = New-Object System.Windows.Forms.Button
$btnSource.Text = "Browse"
$btnSource.Location = New-Object System.Drawing.Point(690,43)
$form.Controls.Add($btnSource)

# =========================================
# OUTPUT
# =========================================
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Output Folder:"
$lblOutput.Location = New-Object System.Drawing.Point(20,80)
$lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Size = New-Object System.Drawing.Size(650,25)
$txtOutput.Location = New-Object System.Drawing.Point(20,105)
$form.Controls.Add($txtOutput)

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = "Browse"
$btnOutput.Location = New-Object System.Drawing.Point(690,103)
$form.Controls.Add($btnOutput)

# =========================================
# FILTER GROUP
# =========================================
$group = New-Object System.Windows.Forms.GroupBox
$group.Text = "Filter Options"
$group.Size = New-Object System.Drawing.Size(800,180)
$group.Location = New-Object System.Drawing.Point(20,145)
$form.Controls.Add($group)

$rbDate = New-Object System.Windows.Forms.RadioButton
$rbDate.Text = "Filter by Date Range"
$rbDate.Location = New-Object System.Drawing.Point(20,30)
$rbDate.AutoSize = $true
$rbDate.Checked = $true
$group.Controls.Add($rbDate)

$rbZip = New-Object System.Windows.Forms.RadioButton
$rbZip.Text = "Filter by ZIP Files"
$rbZip.Location = New-Object System.Drawing.Point(220,30)
$rbZip.AutoSize = $true
$group.Controls.Add($rbZip)

# Date controls
$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = "Start Date:"
$lblStart.Location = New-Object System.Drawing.Point(20,60)
$lblStart.AutoSize = $true
$group.Controls.Add($lblStart)

$dtStart = New-Object System.Windows.Forms.DateTimePicker
$dtStart.Format = "Custom"
$dtStart.CustomFormat = "yyyy-MM-dd HH:mm:ss"
$dtStart.Location = New-Object System.Drawing.Point(100,55)
$group.Controls.Add($dtStart)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = "End Date:"
$lblEnd.Location = New-Object System.Drawing.Point(350,60)
$lblEnd.AutoSize = $true
$group.Controls.Add($lblEnd)

$dtEnd = New-Object System.Windows.Forms.DateTimePicker
$dtEnd.Format = "Custom"
$dtEnd.CustomFormat = "yyyy-MM-dd HH:mm:ss"
$dtEnd.Location = New-Object System.Drawing.Point(420,55)
$group.Controls.Add($dtEnd)

# ZIP controls
$lblZip = New-Object System.Windows.Forms.Label
$lblZip.Text = "Enter ZIP Names (comma separated, no .zip):"
$lblZip.Location = New-Object System.Drawing.Point(20,95)
$lblZip.AutoSize = $true
$lblZip.Visible = $false
$group.Controls.Add($lblZip)

$txtZipNames = New-Object System.Windows.Forms.TextBox
$txtZipNames.Size = New-Object System.Drawing.Size(600,25)
$txtZipNames.Location = New-Object System.Drawing.Point(20,120)
$txtZipNames.Visible = $false
$group.Controls.Add($txtZipNames)

$btnSelectZip = New-Object System.Windows.Forms.Button
$btnSelectZip.Text = "Browse ZIP Files..."
$btnSelectZip.Size = New-Object System.Drawing.Size(150,30)
$btnSelectZip.Location = New-Object System.Drawing.Point(640,118)
$btnSelectZip.Visible = $false
$group.Controls.Add($btnSelectZip)

# Toggle visibility
$rbDate.Add_CheckedChanged({
    $dtStart.Visible = $rbDate.Checked
    $dtEnd.Visible = $rbDate.Checked
    $lblStart.Visible = $rbDate.Checked
    $lblEnd.Visible = $rbDate.Checked

    $lblZip.Visible = !$rbDate.Checked
    $txtZipNames.Visible = !$rbDate.Checked
    $btnSelectZip.Visible = !$rbDate.Checked
})

$rbZip.Add_CheckedChanged({
    $dtStart.Visible = !$rbZip.Checked
    $dtEnd.Visible = !$rbZip.Checked
    $lblStart.Visible = !$rbZip.Checked
    $lblEnd.Visible = !$rbZip.Checked

    $lblZip.Visible = $rbZip.Checked
    $txtZipNames.Visible = $rbZip.Checked
    $btnSelectZip.Visible = $rbZip.Checked
})

# =========================================
# PROGRESS
# =========================================
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,340)
$progressBar.Size = New-Object System.Drawing.Size(800,25)
$form.Controls.Add($progressBar)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20,370)
$lblStatus.Size = New-Object System.Drawing.Size(800,20)
$form.Controls.Add($lblStatus)

# =========================================
# STATUS BOX
# =========================================
$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.Location = New-Object System.Drawing.Point(20,400)
$txtStatus.Size = New-Object System.Drawing.Size(800,140)
$form.Controls.Add($txtStatus)

# =========================================
# BUTTONS
# =========================================
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start"
$btnStart.Location = New-Object System.Drawing.Point(280,550)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(400,550)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Location = New-Object System.Drawing.Point(520,550)
$form.Controls.Add($btnExit)

# =========================================
# DIALOGS
# =========================================
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "ZIP files (*.zip)|*.zip"
$openFileDialog.Multiselect = $true

$btnSource.Add_Click({
    if ($folderDialog.ShowDialog() -eq "OK") {
        $txtSource.Text = $folderDialog.SelectedPath
    }
})

$btnOutput.Add_Click({
    if ($folderDialog.ShowDialog() -eq "OK") {
        $txtOutput.Text = $folderDialog.SelectedPath
    }
})

$btnSelectZip.Add_Click({
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $names = $openFileDialog.FileNames |
            ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_) }
        $txtZipNames.Text = ($names -join ", ")
    }
})

# =========================================
# SAFE COPY
# =========================================
function Copy-SafeFile {
    param ($SourceFile, $DestinationFolder)
    $fileName = [System.IO.Path]::GetFileName($SourceFile)
    $destPath = Join-Path $DestinationFolder $fileName
    $counter = 1
    while (Test-Path $destPath) {
        $nameOnly = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        $ext = [System.IO.Path]::GetExtension($fileName)
        $destPath = Join-Path $DestinationFolder "$nameOnly`_$counter$ext"
        $counter++
    }
    Copy-Item $SourceFile $destPath
}

# =========================================
# PROCESSING
# =========================================
$btnCancel.Add_Click({ $script:CancelRequested = $true })

$btnStart.Add_Click({

    if ($script:IsProcessing) { return }

    try {

        $Source = $txtSource.Text
        $Output = $txtOutput.Text

        if (!(Test-Path $Source)) { throw "Invalid Source Folder" }
        if (!(Test-Path $Output)) { New-Item -ItemType Directory -Path $Output | Out-Null }

        if ($rbDate.Checked) {
            $ZipFiles = Get-ChildItem $Source -Filter *.zip |
                Where-Object { $_.LastWriteTime -ge $dtStart.Value -and $_.LastWriteTime -le $dtEnd.Value }
        } else {
            if ([string]::IsNullOrWhiteSpace($txtZipNames.Text)) {
                throw "Enter ZIP names or select ZIP files."
            }

            $names = $txtZipNames.Text.Split(",") |
                ForEach-Object { "$($_.Trim()).zip" }

            $ZipFiles = foreach ($n in $names) {
                Get-ChildItem $Source -Filter $n -ErrorAction SilentlyContinue
            }
        }

        if (!$ZipFiles) { throw "No ZIP files found." }

        $script:IsProcessing = $true
        $script:CancelRequested = $false
        $btnStart.Enabled = $false
        $btnCancel.Enabled = $true

        $total = $ZipFiles.Count
        $current = 0

        foreach ($zip in $ZipFiles) {

            if ($script:CancelRequested) { break }

            $temp = Join-Path $env:TEMP ([guid]::NewGuid())
            New-Item -ItemType Directory -Path $temp | Out-Null
            Expand-Archive $zip.FullName $temp -Force

            $files = Get-ChildItem $temp -Recurse -File
            foreach ($file in $files) {
                Copy-SafeFile $file.FullName $Output
            }

            Remove-Item $temp -Recurse -Force

            $current++
            $percent = [math]::Round(($current / $total) * 100)
            $progressBar.Value = $percent
            $lblStatus.Text = "Processing $current of $total ($percent%)"
            $txtStatus.AppendText("Processed: $($zip.Name)`r`n")
            [System.Windows.Forms.Application]::DoEvents()
        }

        Add-Content (Join-Path $LogFolder "ZipUtility.log") "$(Get-Date) | User:$env:USERNAME | ZIPs:$current"

        if ($script:CancelRequested) {
            $lblStatus.Text = "Operation Cancelled"
        } else {
            $lblStatus.Text = "Completed Successfully"
        }

    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_)
    }
    finally {
        $script:IsProcessing = $false
        $btnStart.Enabled = $true
        $btnCancel.Enabled = $false
    }
})

# =========================================
# EXIT
# =========================================
$btnExit.Add_Click({
    if ($script:IsProcessing) {
        [System.Windows.Forms.MessageBox]::Show("Processing is running. Please cancel first.")
        return
    }
    $form.Close()
})

# REQUIRED FOR ps2exe
[System.Windows.Forms.Application]::Run($form)
