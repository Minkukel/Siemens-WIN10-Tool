# Opent het script als admin als dit nodig is.
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
        Exit
    }
}

function Startup{
#Reset alle variabelen
    rv * -ea SilentlyContinue; rmo *; $error.Clear(); cls
#Voegt windows form componenten toe
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
#Laat weten of script goed start
    Write-Host "Script Started Startup"
#Opent keuze menu om .csv bestand met software instellingen in te laden.
    $global:CSVfile = Get-Database
# Als .csv path niet leeg in geeft die weer dat CSV correct ingeladen is, en ander geeft die foutmelding
    if ($CSVfile -ne ""){
        Write-Host "CSV Correct geladen"
        MainScreen
    } else {
        Write-Host "CSV bestand error"
    }
}

Function Get-Database($initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    #Laat weten of script goed start
    Write-Host "Script Started Get-Database"
    $DatabaseDialog = New-Object System.Windows.Forms.OpenFileDialog
    $DatabaseDialog.initialDirectory = $initialDirectory
    $DatabaseDialog.filter = "CSV (*.csv)| *.csv"
#Zet loop aan
    $loop = $true
    while ($loop) {
        if ($DatabaseDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
        {
            $DatabaseDialog.filename
            $loop = $false
        }else {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "cancel"){
                [System.Windows.Forms.MessageBox]::Show("Je moet een database kiezen omdit programma te gebruiken")
                throw "Je moet een database kiezen omdit programma te gebruiken"
        }
    }
    $DatabaseDialog.Dispose()
    }
}


function MainScreen{
    $Main_Form                          = New-Object system.Windows.Forms.Form
    $Main_Form.ClientSize               = '824,584'
    $Main_Form.text                     = "Siemens WIN10 TOOLBOX by Rein Veenstra   JC-ELECTRONICS"
    $Main_Form.AutoSize                 = $true
    $Main_Form.AutoSizeMode             = "GrowAndShrink"
    $Main_Form.MaximizeBox              = $false
    $Main_Form.WindowState              = "Normal"
                                            # Maximized, Minimized, Normal
    $Main_Form.SizeGripStyle            = "Hide"
    $Main_Form.StartPosition            = "CenterScreen"
                                            # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
    $Main_Form.TopMost                  = $true


    $SoftwareLabel                      = New-Object system.Windows.Forms.Label
    $SoftwareLabel.text                 = "Software:"
    $SoftwareLabel.AutoSize             = $true
    $SoftwareLabel.width                = 25
    $SoftwareLabel.height               = 10
    $SoftwareLabel.location             = New-Object System.Drawing.Point(5,20)
    $SoftwareLabel.Font                 = 'Microsoft Sans Serif,10'

    $SoftwareBox                        = New-Object system.Windows.Forms.ComboBox
    $SoftwareBox.text                   = "Selecteer welke software je wilt installeren."
    $SoftwareBox.width                  = 328
    $SoftwareBox.height                 = 39
    $SoftwareBox.location               = New-Object System.Drawing.Point(76,19)
    $SoftwareBox.Font                   = 'Microsoft Sans Serif,10'

    $NetInstallStatusButton             = New-Object system.Windows.Forms.RadioButton
    $NetInstallStatusButton.text        = "Geinstalleerd!"
    $NetInstallStatusButton.AutoSize    = $true
    $NetInstallStatusButton.width       = 104
    $NetInstallStatusButton.height      = 20
    $NetInstallStatusButton.enabled     = $false
    $NetInstallStatusButton.location    = New-Object System.Drawing.Point(715,19)
    $NetInstallStatusButton.Font        = 'Microsoft Sans Serif,10'

    $ModusLabel                         = New-Object system.Windows.Forms.Label
    $ModusLabel.text                    = "Windows compatibileits modus:"
    $ModusLabel.AutoSize                = $true
    $ModusLabel.width                   = 25
    $ModusLabel.height                  = 10
    $ModusLabel.location                = New-Object System.Drawing.Point(5,67)
    $ModusLabel.Font                    = 'Microsoft Sans Serif,10'

    $ModusTextBox                       = New-Object system.Windows.Forms.TextBox
    $ModusTextBox.multiline             = $false
    $ModusTextBox.width                 = 202
    $ModusTextBox.height                = 20
    $ModusTextBox.location              = New-Object System.Drawing.Point(202,63)
    $ModusTextBox.Font                  = 'Microsoft Sans Serif,10'

    $SetupDirLabel                      = New-Object system.Windows.Forms.Label
    $SetupDirLabel.text                 = "Setup Exe:"
    $SetupDirLabel.AutoSize             = $true
    $SetupDirLabel.width                = 25
    $SetupDirLabel.height               = 10
    $SetupDirLabel.location             = New-Object System.Drawing.Point(5,224)
    $SetupDirLabel.Font                 = 'Microsoft Sans Serif,10'

    $SetupDirTextBox                    = New-Object system.Windows.Forms.TextBox
    $SetupDirTextBox.multiline          = $false
    $SetupDirTextBox.width              = 117
    $SetupDirTextBox.height             = 20
    $SetupDirTextBox.location           = New-Object System.Drawing.Point(85,220)
    $SetupDirTextBox.Font               = 'Microsoft Sans Serif,10'

    $InstallDirLabel                    = New-Object system.Windows.Forms.Label
    $InstallDirLabel.text               = "Installatie bestanden map:"
    $InstallDirLabel.AutoSize           = $true
    $InstallDirLabel.width              = 25
    $InstallDirLabel.height             = 10
    $InstallDirLabel.location           = New-Object System.Drawing.Point(5,258)
    $InstallDirLabel.Font               = 'Microsoft Sans Serif,10'

    $SetupDirButton                     = New-Object system.Windows.Forms.Button
    $SetupDirButton.text                = "Browse"
    $SetupDirButton.width               = 60
    $SetupDirButton.height              = 30
    $SetupDirButton.location            = New-Object System.Drawing.Point(220,213)
    $SetupDirButton.Font                = 'Microsoft Sans Serif,10'

    $InstallDirButton                   = New-Object system.Windows.Forms.Button
    $InstallDirButton.text              = "Browse"
    $InstallDirButton.width             = 60
    $InstallDirButton.height            = 30
    $InstallDirButton.location          = New-Object System.Drawing.Point(421,248)
    $InstallDirButton.Font              = 'Microsoft Sans Serif,10'

    $NetInstallLabel                    = New-Object system.Windows.Forms.Label
    $NetInstallLabel.text               = ".NET3.5 installeren"
    $NetInstallLabel.AutoSize           = $true
    $NetInstallLabel.width              = 25
    $NetInstallLabel.height             = 10
    $NetInstallLabel.location           = New-Object System.Drawing.Point(6,314)
    $NetInstallLabel.Font               = 'Microsoft Sans Serif,10,style=Bold'

    $NetInstallButton                   = New-Object system.Windows.Forms.Button
    $NetInstallButton.text              = "Install"
    $NetInstallButton.width             = 60
    $NetInstallButton.height            = 30
    $NetInstallButton.location          = New-Object System.Drawing.Point(177,304)
    $NetInstallButton.Font              = 'Microsoft Sans Serif,10'

    $NetInstallStatusLabel              = New-Object system.Windows.Forms.Label
    $NetInstallStatusLabel.text         = "Is .NET3.5 geinstalleerd?"
    $NetInstallStatusLabel.AutoSize     = $true
    $NetInstallStatusLabel.width        = 25
    $NetInstallStatusLabel.height       = 10
    $NetInstallStatusLabel.location     = New-Object System.Drawing.Point(537,20)
    $NetInstallStatusLabel.Font         = 'Microsoft Sans Serif,10,style=Bold'

    $VCInstallStatusLabel               = New-Object system.Windows.Forms.Label
    $VCInstallStatusLabel.text          = "Zijn de VC+ pakketten geinstalleerd?"
    $VCInstallStatusLabel.AutoSize      = $true
    $VCInstallStatusLabel.width         = 25
    $VCInstallStatusLabel.height        = 10
    $VCInstallStatusLabel.location      = New-Object System.Drawing.Point(470,63)
    $VCInstallStatusLabel.Font          = 'Microsoft Sans Serif,10,style=Bold'

    $VCInstallStatusButton              = New-Object system.Windows.Forms.RadioButton
    $VCInstallStatusButton.text         = "Geinstalleerd!"
    $VCInstallStatusButton.AutoSize     = $true
    $VCInstallStatusButton.width        = 104
    $VCInstallStatusButton.height       = 20
    $VCInstallStatusButton.enabled      = $false
    $VCInstallStatusButton.location     = New-Object System.Drawing.Point(716,63)
    $VCInstallStatusButton.Font         = 'Microsoft Sans Serif,10'

    $VCInstallLabel                     = New-Object system.Windows.Forms.Label
    $VCInstallLabel.text                = "VC+ pakketten installeren"
    $VCInstallLabel.AutoSize            = $true
    $VCInstallLabel.width               = 25
    $VCInstallLabel.height              = 10
    $VCInstallLabel.location            = New-Object System.Drawing.Point(5,352)
    $VCInstallLabel.Font                = 'Microsoft Sans Serif,10,style=Bold'

    $VCInstallButton                    = New-Object system.Windows.Forms.Button
    $VCInstallButton.text               = "Install"
    $VCInstallButton.width              = 60
    $VCInstallButton.height             = 30
    $VCInstallButton.location           = New-Object System.Drawing.Point(177,342)
    $VCInstallButton.Font               = 'Microsoft Sans Serif,10'

    $InstallButton                      = New-Object system.Windows.Forms.Button
    $InstallButton.text                 = "Start Install"
    $InstallButton.width                = 105
    $InstallButton.height               = 55
    $InstallButton.location             = New-Object System.Drawing.Point(421,307)
    $InstallButton.Font                 = 'Microsoft Sans Serif,10,style=Bold'

    $InstallDirTextBox                  = New-Object system.Windows.Forms.TextBox
    $InstallDirTextBox.multiline        = $false
    $InstallDirTextBox.width            = 232
    $InstallDirTextBox.height           = 20
    $InstallDirTextBox.location         = New-Object System.Drawing.Point(176,254)
    $InstallDirTextBox.Font             = 'Microsoft Sans Serif,10'

    $LogTextBox                         = New-Object system.Windows.Forms.TextBox
    $LogTextBox.multiline               = $true
    $LogTextBox.width                   = 519
    $LogTextBox.height                  = 138
    $LogTextBox.location                = New-Object System.Drawing.Point(11,422)
    $LogTextBox.Font                    = 'Microsoft Sans Serif,10'

    $LogLabel                           = New-Object system.Windows.Forms.Label
    $LogLabel.text                      = "Log Output:"
    $LogLabel.AutoSize                  = $true
    $LogLabel.width                     = 25
    $LogLabel.height                    = 10
    $LogLabel.location                  = New-Object System.Drawing.Point(14,400)
    $LogLabel.Font                      = 'Microsoft Sans Serif,10'

    $ExtraButton                        = New-Object system.Windows.Forms.Button
    $ExtraButton.text                   = "Extra Screen"
    $ExtraButton.width                  = 85
    $ExtraButton.height                 = 35
    $ExtraButton.location               = New-Object System.Drawing.Point(725,94)
    $ExtraButton.Font                   = 'Microsoft Sans Serif,10'


    $Main_Form.controls.AddRange(@($SoftwareLabel,$SoftwareBox,$NetInstallStatusButton,$ModusLabel,$ModusTextBox,$SetupDirLabel,$SetupDirTextBox,$InstallDirLabel,$SetupDirButton,$InstallDirButton,$NetInstallLabel,$NetInstallButton,$NetInstallStatusLabel,$VCInstallStatusLabel,$VCInstallStatusButton,$VCInstallLabel,$VCInstallButton,$InstallButton,$InstallDirTextBox,$LogTextBox,$LogLabel,$ExtraButton))


    #Write your logic code here
    Write-Host "Script Started MainScreen"

    #Variabelen
    $Software_list = Import-CSV $CSVfile
    $Software_Names = $Software_list | select -ExpandProperty Software
    $Exe_location = $SetupDirTextBox.Text
    $Global:Selected_Program = ""


    #NET3.5 Status laden
    $NetStatus = (Get-WindowsOptionalFeature -Online -FeatureName NETFX3).State

    #Verwijzingen van de buttons.
    $InstallButton.Add_Click({ Install-Program })
    $InstallDirButton.Add_Click({ Get-InstallPath })
    $SetupDirButton.Add_Click({ Get-InstallEXEFileName })
    $ExtraButton.Add_Click({ Get-NET3.5Info })


    Foreach ($Software_Names in $Software_Names)
    {
        $SoftwareBox.Items.Add($Software_Names);
    }

    #Map kiezen waar de installatie bestanden in staan
    $SoftwareBox.add_SelectedIndexChanged({GetModusSelected})

    function GetModusSelected
    {
        $ModusTextBox.Text = ($Software_List | where Software -eq $SoftwareBox.SelectedItem).OS_Version_Exe
        $Global:Selected_Program = $SoftwareBox.SelectedItem
        $InstallDirButton.Enabled = $true
        $SetupDirButton.Enabled = $true
    } #end GetModusSelected

    function Get-InstallPath
    {
        $InstallDirTextBox.Text = Get-Folder
    }

    Function Get-Folder
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null
        $foldername = New-Object System.Windows.Forms.FolderBrowserDialog
        $foldername.Description = "Selecteer de map waar de bestanden van", "$Selected_Program", "staan!"
        $foldername.rootfolder = "MyComputer"
        if($foldername.ShowDialog() -eq "OK")
        {
            $folder += $foldername.SelectedPath
        }
        return $folder
        $foldername.Dispose()
    }

    function Get-InstallEXEFileName
    {
        $SetupDirTextBox.Text = Get-FileName -initialDirectory  $InstallDirTextBox.Text -FileName ($Software_List | where Software -eq $SoftwareBox.SelectedItem).Exe_name
    }

    Function Get-FileName($initialDirectory, $FileName)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.FileName = $FileName
        $OpenFileDialog.filter = "EXE (*.EXE)| *.EXE"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
        $OpenFileDialog.Dispose()
    }

    Function Install-Program()
    {
        $LogTextBox.text = $installDirTextBox.Text
    }


    #Aanduiden als .NET3.5 geinstalleerd is!
    if ($NetStatus -eq "Enabled")
    {
        $NetInstallStatusButton.checked = $True
        $NetInstallStatusButton.enabled = $True
        $NetInstallButton.Enabled = $false
    }

    if ($Selected_Program -eq "")
    {
        $InstallDirButton.Enabled = $false
        $SetupDirButton.Enabled = $false
    }

    Function Get-NET3.5Info
    {
        $LogTextBox.text = (Get-WindowsOptionalFeature -Online -FeatureName NETFX3).state
    }

    $Main_Form.ShowDialog()
}
Startup