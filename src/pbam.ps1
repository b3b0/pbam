# if i'm not admin, make me admin

If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{   
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

# path variables

$installdir = "C:\Program Files\PBAM"
$cache = "$installdir\.btlckr_cache"
$exepath = "$installdir\lib\wkhtmltopdf.exe"
$configpath = "$installdir\conf\pbam.conf"
$configargs = Get-Content $configpath

if (! (Test-Path $cache))
{
    New-Item -ItemType directory -Path $cache | Out-Null
}

if (! (Test-Path $configpath))
{
    New-Item -ItemType file -Path $configpath | Out-Null
}

# random filename variables - it's done this way because the filenames are not important

$randomdatacache = [guid]::NewGuid()
$randomcsv = [guid]::NewGuid()
$randomhtmlfile = [guid]::NewGuid()
$randomchart = [guid]::NewGuid()
$randompdf = [guid]::NewGuid()
$outfile = "$cache\$randomdatacache.dat"
$outcsv = "$cache\$randomcsv.csv"
$outhtml = "$cache\$randomhtmlfile.html"
$outchart = "$cache\$randomchart.png"
$outpdf = "$cache\$randompdf.pdf"

# function to pull bitlocker data from AD

function puller($arbi)
{
    # tell the user to wait while you get a steaming hot cuppa data

    Add-Type -AssemblyName System.Windows.Forms 
    $waiter = New-Object system.Windows.Forms.Form
    $Label = New-Object System.Windows.Forms.Label
    $waiter.Controls.Add($Label)
    $Label.Text = "Getting data! Be patient!"
    $Label.AutoSize = $True
    $waiter.Visible = $True
    $waiter.Update()

    # get the bitlocker info and shoot it into a .dat file
    
    Import-Module ActiveDirectory
    Get-ADComputer -Filter 'ObjectClass -eq "computer"' -SearchBase $configargs  | foreach-object {
        $Computer = $_.name
        $Computer_Object = Get-ADComputer -Filter {cn -eq $Computer} -Property msTPM-OwnerInformation, msTPM-TpmInformationForComputer
        if($null -eq $Computer_Object)
        {
            Write-Host "Error..."
        }
        $Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $Computer_Object.DistinguishedName -Properties 'msFVE-RecoveryPassword' | Select-Object -Last 1
        if($Bitlocker_Object.'msFVE-RecoveryPassword')
        {
            $Bitlocker_Key = $Bitlocker_Object.'msFVE-RecoveryPassword'
        }
        else
        {
            $Bitlocker_Key = "none"
        }
        $strToReport = $Computer + "," + $Bitlocker_Key  
        $strToReport | Out-File $outfile -append
    }

    # start crafting the html report

    $css = @"
    <style>
    h1, h2, h5, th { text-align: center; font-family: Segoe UI; }
    table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
    th { background: #0046c3; color: #fff; max-width: 400px; padding: 5px 10px; }
    td { font-size: 11px; padding: 5px 20px; color: #000; }
    tr { background: #b8d1f3; }
    tr:nth-child(even) { background: #dae5f4; }
    tr:nth-child(odd) { background: #b8d1f3; }
    </style>
"@

    # counts of enabled, disabled and total devices

    $enableddevices = (Get-Content $outfile | Select-String "none" -NotMatch | Measure-Object -line ).Lines
    $disableddevices = (Get-Content $outfile | Select-String "none" | Measure-Object -line ).Lines
    $totaldevices = (Get-Content $outfile | Measure-Object -line ).Lines

    # let's chart this shit out

    $bitlockerstatus = @()
    $properties = @{ status = ""
                    count = 0}

    # add the first object to the array

    $currentcount = New-Object -TypeName psobject -Property $properties
    $currentcount.status = "Published to AD"
    $currentcount.count = [int]$enableddevices
    $bitlockerstatus+= $currentcount

    # add the second object to the array

    $currentcount = New-Object -TypeName psobject -Property $properties
    $currentcount.status = "Unpublished to AD"
    $currentcount.count = [int]$disableddevices
    $bitlockerstatus+= $currentcount

    # use the script to make the image

    & $installdir\lib\Get-Corpchart-LightEdition -data $bitlockerstatus  -obj_key "status" -obj_value "count" -FilePath $outchart -type Doughnut -Show_percentage_pie

    # list our devices in order, shoot it into a CSV

    Get-Content $outfile | Sort-Object > $outcsv
    @("Computer,Key") + (Get-Content $outcsv) | Set-Content $outcsv

    # make the html page

    Import-CSV $outcsv | ConvertTo-Html -head $css -body "<h1>Executive Report</h1>`n<h2>Disk Encryption Enforced By Bitlocker</h2>`n<h5>Generated on $(Get-Date) by $env:UserName </h5>`n<img src='$randomchart.png'>`n<br><b>$enableddevices out of $totaldevices are protected.</b>`n"  | Out-File $outhtml
    
    # use our little python script to turn the html page into a pdf
    
    python $installdir\lib\pdffer.py $outhtml $outpdf $exepath
    
    $waiter.Close()
}

#click function to return information

function go
{
    $searchterm = $specificTextBox.Text.Trim()
    $OutputBox.Clear()
    if ($enabledButton.Checked)
    {
        $OutputBox.text += "`r`n" + "BITLOCKER ENABLED DEVICES:"
        $OutputBox.text += "`r`n" + "#########################"
        Foreach ($match in (Get-Content $outfile | Select-String "none" -NotMatch))
        {
            $OutputBox.text += "`r`n" + $match
        }
        $OutputBox.text += "`r`n" + "#########################"
    }
    if ($disabledButton.Checked)
    {
        $OutputBox.text += "`r`n" + "BITLOCKER DISABLED DEVICES:"
        $OutputBox.text += "`r`n" + "#########################"

        Foreach ($match in (Get-Content $outfile | Select-String "none"))
        {
            $OutputBox.text += "`r`n" + $match
        }
        $OutputBox.text += "`r`n" + "#########################"
    }
    if ($specificButton.Checked)
    {
        $OutputBox.text += "`r`n" + "RESULTS FOR " + $searchterm
        $OutputBox.text += "`r`n" + "#########################"

        Foreach ($match in (Get-Content $outfile | Select-String $searchterm))
        {
            $OutputBox.text += "`r`n" + $match
        }
        $OutputBox.text += "`r`n" + "#########################"
    }
    if ($reportButton.Checked)
    {
        reporter $outhtml $outpdf
    }
}

function reporter($html,$pdf)
{
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $reportform                            = New-Object system.Windows.Forms.Form
    $reportform.ClientSize                 = '400,156'
    $reportform.text                       = "Executive Report Puller"
    $reportform.TopMost                    = $false

    $pleasechoose                          = New-Object system.Windows.Forms.Label
    $pleasechoose.text                     = "Please choose your report format."
    $pleasechoose.AutoSize                 = $false
    $pleasechoose.width                    = 350
    $pleasechoose.height                   = 15
    $pleasechoose.location                 = New-Object System.Drawing.Point(102,49)
    $pleasechoose.Font                     = 'Microsoft Sans Serif,10'

    $htmlbutton                         = New-Object system.Windows.Forms.Button
    $htmlbutton.text                    = "HTML"
    $htmlbutton.width                   = 60
    $htmlbutton.height                  = 30
    $htmlbutton.location                = New-Object System.Drawing.Point(8,111)
    $htmlbutton.Font                    = 'Microsoft Sans Serif,10'

    $pdfbutton                         = New-Object system.Windows.Forms.Button
    $pdfbutton.text                    = "PDF"
    $pdfbutton.width                   = 60
    $pdfbutton.height                  = 30
    $pdfbutton.location                = New-Object System.Drawing.Point(80,111)
    $pdfbutton.Font                    = 'Microsoft Sans Serif,10'

    $cancelbutton                         = New-Object system.Windows.Forms.Button
    $cancelbutton.text                    = "Cancel"
    $cancelbutton.width                   = 60
    $cancelbutton.height                  = 30
    $cancelbutton.location                = New-Object System.Drawing.Point(330,111)
    $cancelbutton.Font                    = 'Microsoft Sans Serif,10'

    $reportform.controls.AddRange(@($pleasechoose,$htmlbutton,$pdfbutton,$cancelbutton))
    $htmlbutton.Add_Click({Invoke-Item $html})
    $pdfbutton.Add_Click({Invoke-Item $pdf})
    $cancelbutton.Add_Click({$reportform.Close()})
    [void]$reportform.ShowDialog()
}

# make the gui

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$MainWindow                            = New-Object system.Windows.Forms.Form
$MainWindow.ClientSize                 = '700,400'
$MainWindow.text                       = "PBAM: PowerShell Bitlocker Administration & Monitoring"
$MainWindow.TopMost                    = $false
$MainWindow.FormBorderStyle            = 'Fixed3D'
$MainWindow.MaximizeBox                = $false

$OutputBox                        = New-Object system.Windows.Forms.TextBox
$OutputBox.multiline              = $true
$OutputBox.width                  = 672
$OutputBox.height                 = 308
$OutputBox.location               = New-Object System.Drawing.Point(16,92)
$OutputBox.Font                   = 'Microsoft Sans Serif,10'
$OutputBox.ScrollBars             = "Vertical"

$enabledButton                    = New-Object system.Windows.Forms.RadioButton
$enabledButton.text               = "Enabled"
$enabledButton.AutoSize           = $true
$enabledButton.width              = 104
$enabledButton.height             = 20
$enabledButton.location           = New-Object System.Drawing.Point(16,1)
$enabledButton.Font               = 'Microsoft Sans Serif,10'

$disabledButton                    = New-Object system.Windows.Forms.RadioButton
$disabledButton.text               = "Disabled"
$disabledButton.AutoSize           = $true
$disabledButton.width              = 104
$disabledButton.height             = 20
$disabledButton.location           = New-Object System.Drawing.Point(16,21)
$disabledButton.Font               = 'Microsoft Sans Serif,10'

$specificButton                    = New-Object system.Windows.Forms.RadioButton
$specificButton.text               = "Specific:"
$specificButton.AutoSize           = $true
$specificButton.width              = 104
$specificButton.height             = 20
$specificButton.location           = New-Object System.Drawing.Point(16,41)
$specificButton.Font               = 'Microsoft Sans Serif,10'

$specificTextBox                        = New-Object system.Windows.Forms.TextBox
$specificTextBox.multiline              = $false
$specificTextBox.width                  = 155
$specificTextBox.height                 = 5
$specificTextBox.location               = New-Object System.Drawing.Point(99,41)
$specificTextBox.Font                   = 'Microsoft Sans Serif,10'

$reportButton                    = New-Object system.Windows.Forms.RadioButton
$reportButton.text               = "Executive Report"
$reportButton.AutoSize           = $true
$reportButton.width              = 104
$reportButton.height             = 20
$reportButton.location           = New-Object System.Drawing.Point(16,61)
$reportButton.Font               = 'Microsoft Sans Serif,10'

$runButton                         = New-Object system.Windows.Forms.Button
$runButton.text                    = "Run"
$runButton.width                   = 60
$runButton.height                  = 30
$runButton.location                = New-Object System.Drawing.Point(630,33)
$runButton.Font                    = 'Microsoft Sans Serif,10'
$MainWindow.controls.AddRange(@($OutputBox,$enabledButton,$disabledButton,$specificButton,$specificTextBox,$reportButton,$runButton))
$runButton.Add_Click({go})

# if there's SOMETHING in the $configpath then let's get crackin' - if not warn the user to put something into it.

if (! (Get-Content $configpath))
{
    $a = new-object -comobject wscript.shell
    $a.popup("Please add your searchbase to $configpath",0,"Invalid Configuration",4)
}

if (Get-Content $configpath)
{
    puller(1)
    [void]$MainWindow.ShowDialog()
}

# delete the cache after we're done

Remove-Item -Path "$cache\*" -Recurse -Force
