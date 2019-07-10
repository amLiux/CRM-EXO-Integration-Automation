#CRM On prem integration EXCO
#CRM-EXO-Integration-Automation
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


#! *************************************** LOGIC ***************************************!#
Function InitializingPSSession(){
    Write-Host "Me estoy ejecutando"  
    Invoke-Expression -Command Enable-PSRemoting 
    Invoke-Expression -Command New-PSSession
}

Function ImportingModules(){
    Write-Host "Me estoy ejecutando x2"  
    Import-Module MSOnline -Force
    Import-Module MSOnlineExt -Force
}

Function GettingCredentials(){
    Write-Host "Me estoy ejecutando x3"
    $msolcred = get-credential
    connect-msolservice -credential $msolcred
}

Function STSCertCreate($PrivateFinalPath, $Password, $PublicFinalPath){
    $STSCertificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $PrivateFinalPath, $Password
    $PFXCertificateBin = $STSCertificate.GetRawCertData()
    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $Certificate.Import(“$PublicFinalPath”)
    $CERCertificateBin = $Certificate.GetRawCertData()
    return $CredentialValue = [System.Convert]::ToBase64String($CERCertificateBin)
}

Function RootDomainChanges($RootDomain1, $CredentialValue){
    Write-Host "Me estoy ejecutando x5"
    $RootDomain = $RootDomain1
    $CRMAppId = "00000007-0000-0000-c000-000000000000" 
    New-MsolServicePrincipalCredential -AppPrincipalId $CRMAppId -Type asymmetric -Usage Verify -Value $CredentialValue
    $CRM = Get-MsolServicePrincipal -AppPrincipalId $CRMAppId
    $ServicePrincipalName = $CRM.ServicePrincipalNames
    $ServicePrincipalName.Remove("$CRMAppId/$RootDomain")
    $ServicePrincipalName.Add("$CRMAppId/$RootDomain")
    Set-MsolServicePrincipal -AppPrincipalId $CRMAppId -ServicePrincipalNames $ServicePrincipalName
}

Function AddingNewSettingToCloud(){
    Write-Host "Me estoy ejecutando x6"
    Add-PSSnapin Microsoft.Crm.PowerShell 
    $setting = New-Object "Microsoft.Xrm.Sdk.Deployment.ConfigurationEntity"
    $setting.LogicalName = "ServerSettings"
    $setting.Attributes = New-Object "Microsoft.Xrm.Sdk.Deployment.AttributeCollection"
    $attribute1 = New-Object "System.Collections.Generic.KeyValuePair[String, Object]" ("S2SDefaultAuthorizationServerPrincipalId", "00000001-0000-0000-c000-000000000000")
    $setting.Attributes.Add($attribute1)
    $attribute2 = New-Object "System.Collections.Generic.KeyValuePair[String, Object]" ("S2SDefaultAuthorizationServerMetadataUrl", "https://accounts.accesscontrol.windows.net/metadata/json/1")
    $setting.Attributes.Add($attribute2)
    Set-CrmAdvancedSetting -Entity $setting
}


#! *************************************** FORMS ***************************************!#
# this creates the first form that checks if the modules MSOnline and MSOnlineExt are installed
$Form = New-Object System.Windows.Forms.Form
    $Form.Text = "CRM + EXO Integration"
    $Form.Size = New-Object System.Drawing.Size(272, 160)
    #You can use Form.Height and Form Width
    $Form.FormBorderStyle = "FixedDialog"
    $Form.TopMost = $true
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    $Form.ControlBox = $true
    $Form.StartPosition = "CenterScreen"
    $Form.Font = "Segoe UI"

#this creates the second form that asks for the path of the pfx certificate and the password.
$Form2 = New-Object System.Windows.Forms.Form
    $Form2.Text = "Parameters for the Integration"
    #$Form2.Size = New-Object System.Drawing.Size(300,200)
    $Form2.Height = 470
    $Form2.Width = 420
    $Form2.TopMost = $true
    $Form2.StartPosition = 'CenterScreen'
    $Form2.FormBorderStyle = 3
    $Form2.Font = "Segoe UI"

#* agregar label del root domain
#! *************************************** LABELS ***************************************!#
#label for path private cert selection Form #2
$label1_Form2= New-Object System.Windows.Forms.Label
    $label1_Form2.Location = New-Object System.Drawing.Point(10,20)
    $label1_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label1_Form2.Text = 'Please select the path for the .pfx certificate:'
    $Form2.Controls.Add($label1_Form2)

#label for password input Form #2
$label2_Form2= New-Object System.Windows.Forms.Label
    $label2_Form2.Location = New-Object System.Drawing.Point(10,100)
    $label2_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label2_Form2.Text = 'Please enter the password for the .pfx certificate:'
    $Form2.Controls.Add($label2_Form2)

#label for public path certificate Form #2
$label2_Form2= New-Object System.Windows.Forms.Label
    $label2_Form2.Location = New-Object System.Drawing.Point(10,180)
    $label2_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label2_Form2.Text = 'Please enter the path for the .cer certificate:'
    $Form2.Controls.Add($label2_Form2)

#label for root domain Form #2
$label3_Form2= New-Object System.Windows.Forms.Label
    $label3_Form2.Location = New-Object System.Drawing.Point(10,260)
    $label3_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label3_Form2.Text = 'Please enter the root domain of the organization:'
    $Form2.Controls.Add($label3_Form2)

#label for form #1
$label_Form = New-Object System.Windows.Forms.Label
    $label_Form.Location = New-Object System.Drawing.Size(8,8)
    $label_Form.Size = New-Object System.Drawing.Size(240, 32)
    $label_Form.TextAlign = "MiddleCenter"
    $label_Form.Text = "We will check if we have the right Azure modules installed"
    $Form.Controls.Add($label_Form)

#* agregar input del root domain
#! *************************************** INPUTS ***************************************!#
$PathInput = New-Object System.Windows.Forms.TextBox
    $PathInput.Location = New-Object System.Drawing.Point(10,50)
    $PathInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($PathInput)

$PasswordInput = New-Object System.Windows.Forms.MaskedTextBox
    $PasswordInput.Location = New-Object System.Drawing.Point(10,130)
    $PasswordInput.Size = New-Object System.Drawing.Size(260,40)
    $PasswordInput.PasswordChar = '*'
    $Form2.Controls.Add($PasswordInput)

$PubPathInput = New-Object System.Windows.Forms.TextBox
    $PubPathInput.Location = New-Object System.Drawing.Point(10,210)
    $PubPathInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($PubPathInput)

$RootInput = New-Object System.Windows.Forms.TextBox
    $RootInput.Location = New-Object System.Drawing.Point(10,290)
    $RootInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($RootInput)

#! *************************************** PROGRESS BAR***************************************!#

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Style="Continuous"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = $width - 40
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size
$progressBar1.Left = 5
$progressBar1.Top = 40

#! *************************************** BUTTONS ***************************************!#
$Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Size(8, 80)
    $Button.Size= New-Object System.Drawing.Size(240, 32)
    $Button.TextAlign = "MiddleCenter"
    $Button.Text = "Test the modules"
    $Button.add_Click({
        $check1 = $false
        $check2 = $false
        if (Get-Module -ListAvailable -Name MSOnline ) {
            $check1 = $true
        } 
        
        if (Get-Module -ListAvailable -Name MSOnlineExt ) {
            $check2 = $true
        } 

       if ($check1 -and $check2){
           $label_Form.Text = "It seems that both MSOnline and MSOnlineExt modules are installed."
           $Button.Text = "Next"
           $Button.add_Click({
               $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $false -and $check2 -eq $false){
	        $label_Form.Text = "We are installing both modules it will take a second."
       	    Install-Module -Name MSOnline
	        Install-Module -Name MSOnlineExt
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $false -and $check2 -eq $true){
            $label_Form.Text = "We are installing one module it will take a second."
            Install-Module -Name MSOnline
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $true -and $check2 -eq $false){
            $label_Form.Text = "We are installing one module it will take a second."
            Install-Module -Name MSOnlineExt
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }
    })
    $Form.Controls.Add($Button)

$PathButton = New-Object System.Windows.Forms.Button
    $PathButton.Location = New-Object System.Drawing.Size(280, 50)
    $PathButton.Size= New-Object System.Drawing.Size(100, 20)
    $PathButton.TextAlign = "MiddleCenter"
    $PathButton.Text = "Open"
    $PathButton.add_Click({
        $PFXCertPath = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $false # Multiple files can be chosen
            Filter = 'Certificates (*.pfx)|*.pfx;' # Specified file types
        }
        [void]$PFXCertPath.ShowDialog()
        $PFXCertPath.SelectedPath
        [string]$FinalPath = $PFXCertPath."FileName"
        $PathInput.Text = $FinalPath
    })
    $Form2.Controls.Add($PathButton)

$PubPathButton = New-Object System.Windows.Forms.Button
    $PubPathButton.Location = New-Object System.Drawing.Size(280, 210)
    $PubPathButton.Size= New-Object System.Drawing.Size(100, 20)
    $PubPathButton.TextAlign = "MiddleCenter"
    $PubPathButton.Text = "Open"
    $PubPathButton.add_Click({
        $CerCertPath = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $false # Multiple files can be chosen
            Filter = 'Certificates (*.cer)|*.cer;' # Specified file types
        }
        [void]$CerCertPath.ShowDialog()
        $CerCertPath.SelectedPath
        [string]$FinalPath = $CerCertPath."FileName"
        $PubPathInput.Text = $FinalPath
    })
    $Form2.Controls.Add($PubPathButton)

$Form2NextButton = New-Object System.Windows.Forms.Button
    $Form2NextButton.Location = New-Object System.Drawing.Size(150, 370)
    $Form2NextButton.Size= New-Object System.Drawing.Size(100, 30)
    $Form2NextButton.TextAlign = "MiddleCenter"
    $Form2NextButton.Text = "Next"
    $Form2NextButton.add_Click({
        $Form2.Controls.Add($progressBar1)
        #gets text from inputs and stored them in these variables
        $Password = $PasswordInput.Text
        $PrivateFinalPath = $PathInput.Text
        $PublicFinalPath = $PubPathInput.Text
        $RootDomain = $RootInput.Text
        #---------------------------------------------------

        #Pulls out the account that is running the MSCRMAsyncService and saves it on a variable
        $ServiceAccount = Get-WmiObject win32_service -filter "name= 'MSCRMAsyncService'" | Format-List StartName | Out-String 
        $ServiceAccount  = $ServiceAccount.Replace(‘StartName :’,  ’’)
        $ServiceAccount = $ServiceAccount.Trim()
        #--------------------------------------------------------------------------------

        if ('C:\Program Files\Microsoft Dynamics CRM\tools' | Test-Path) { 
            Set-Location 'C:\Program Files\Microsoft Dynamics CRM\tools'
        }elseif('C:\Program Files\Dynamics 365\Tools'  | Test-Path){
            Set-Location 'C:\Program Files\Dynamics 365\Tools' 
        }

        Try{
            #dot sources certificate reconfiguration and passes the parameters that were asked before
            # . ".\CertificateReconfiguration.ps1"  -certificateFile $PrivateFinalPath -password $Password -certificateType S2STokenIssuer -updateCrm -serviceAccount $ServiceAccount -storeFindType FindBySubjectDistinguishedName
            #---------------------------------------------------------------------------------
            InitializingPSSession 
            ImportingModules
            GettingCredentials
            $CredValue = STSCertCreate -PrivateFinalPath $PrivateFinalPath -Password $Password -PublicFinalPath $PublicFinalPath
            RootDomainChanges -RootDomain1 $RootDomain -CredentialValue $CredValue

        }   
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host "An error has occured $ErrorMessage  in $FailedItem" -ForegroundColor Red -BackgroundColor White
        }
    })#CRM On prem integration EXCO
#CRM-EXO-Integration-Automation
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")


#! *************************************** LOGIC ***************************************!#
Function InitializingPSSession(){
    Write-Host "Me estoy ejecutando"  
    Invoke-Expression -Command Enable-PSRemoting 
    Invoke-Expression -Command New-PSSession
}

Function ImportingModules(){
    Write-Host "Me estoy ejecutando x2"  
    Import-Module MSOnline -Force
    Import-Module MSOnlineExt -Force
}

Function GettingCredentials(){
    Write-Host "Me estoy ejecutando x3"
    $msolcred = get-credential
    connect-msolservice -credential $msolcred
}

Function STSCertCreate($PrivateFinalPath, $Password, $PublicFinalPath){
    $STSCertificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList $PrivateFinalPath, $Password
    $PFXCertificateBin = $STSCertificate.GetRawCertData()
    $Certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $Certificate.Import(“$PublicFinalPath”)
    $CERCertificateBin = $Certificate.GetRawCertData()
    return $CredentialValue = [System.Convert]::ToBase64String($CERCertificateBin)
}

Function RootDomainChanges($RootDomain1, $CredentialValue){
    Write-Host "Me estoy ejecutando x5"
    $RootDomain = $RootDomain1
    $CRMAppId = "00000007-0000-0000-c000-000000000000" 
    New-MsolServicePrincipalCredential -AppPrincipalId $CRMAppId -Type asymmetric -Usage Verify -Value $CredentialValue
    $CRM = Get-MsolServicePrincipal -AppPrincipalId $CRMAppId
    $ServicePrincipalName = $CRM.ServicePrincipalNames
    $ServicePrincipalName.Remove("$CRMAppId/$RootDomain")
    $ServicePrincipalName.Add("$CRMAppId/$RootDomain")
    Set-MsolServicePrincipal -AppPrincipalId $CRMAppId -ServicePrincipalNames $ServicePrincipalName
}

Function AddingNewSettingToCloud(){
    Write-Host "Me estoy ejecutando x6"
    Add-PSSnapin Microsoft.Crm.PowerShell 
    $setting = New-Object "Microsoft.Xrm.Sdk.Deployment.ConfigurationEntity"
    $setting.LogicalName = "ServerSettings"
    $setting.Attributes = New-Object "Microsoft.Xrm.Sdk.Deployment.AttributeCollection"
    $attribute1 = New-Object "System.Collections.Generic.KeyValuePair[String, Object]" ("S2SDefaultAuthorizationServerPrincipalId", "00000001-0000-0000-c000-000000000000")
    $setting.Attributes.Add($attribute1)
    $attribute2 = New-Object "System.Collections.Generic.KeyValuePair[String, Object]" ("S2SDefaultAuthorizationServerMetadataUrl", "https://accounts.accesscontrol.windows.net/metadata/json/1")
    $setting.Attributes.Add($attribute2)
    Set-CrmAdvancedSetting -Entity $setting
}


#! *************************************** FORMS ***************************************!#
# this creates the first form that checks if the modules MSOnline and MSOnlineExt are installed
$Form = New-Object System.Windows.Forms.Form
    $Form.Text = "CRM + EXO Integration"
    $Form.Size = New-Object System.Drawing.Size(272, 160)
    #You can use Form.Height and Form Width
    $Form.FormBorderStyle = "FixedDialog"
    $Form.TopMost = $true
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    $Form.ControlBox = $true
    $Form.StartPosition = "CenterScreen"
    $Form.Font = "Segoe UI"

#this creates the second form that asks for the path of the pfx certificate and the password.
$Form2 = New-Object System.Windows.Forms.Form
    $Form2.Text = "Parameters for the Integration"
    #$Form2.Size = New-Object System.Drawing.Size(300,200)
    $Form2.Height = 470
    $Form2.Width = 420
    $Form2.TopMost = $true
    $Form2.StartPosition = 'CenterScreen'
    $Form2.FormBorderStyle = 3
    $Form2.Font = "Segoe UI"

#* agregar label del root domain
#! *************************************** LABELS ***************************************!#
#label for path private cert selection Form #2
$label1_Form2= New-Object System.Windows.Forms.Label
    $label1_Form2.Location = New-Object System.Drawing.Point(10,20)
    $label1_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label1_Form2.Text = 'Please select the path for the .pfx certificate:'
    $Form2.Controls.Add($label1_Form2)

#label for password input Form #2
$label2_Form2= New-Object System.Windows.Forms.Label
    $label2_Form2.Location = New-Object System.Drawing.Point(10,100)
    $label2_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label2_Form2.Text = 'Please enter the password for the .pfx certificate:'
    $Form2.Controls.Add($label2_Form2)

#label for public path certificate Form #2
$label2_Form2= New-Object System.Windows.Forms.Label
    $label2_Form2.Location = New-Object System.Drawing.Point(10,180)
    $label2_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label2_Form2.Text = 'Please enter the path for the .cer certificate:'
    $Form2.Controls.Add($label2_Form2)

#label for root domain Form #2
$label3_Form2= New-Object System.Windows.Forms.Label
    $label3_Form2.Location = New-Object System.Drawing.Point(10,260)
    $label3_Form2.Size = New-Object System.Drawing.Size(280,15)
    $label3_Form2.Text = 'Please enter the root domain of the organization:'
    $Form2.Controls.Add($label3_Form2)

#label for form #1
$label_Form = New-Object System.Windows.Forms.Label
    $label_Form.Location = New-Object System.Drawing.Size(8,8)
    $label_Form.Size = New-Object System.Drawing.Size(240, 32)
    $label_Form.TextAlign = "MiddleCenter"
    $label_Form.Text = "We will check if we have the right Azure modules installed"
    $Form.Controls.Add($label_Form)

#* agregar input del root domain
#! *************************************** INPUTS ***************************************!#
$PathInput = New-Object System.Windows.Forms.TextBox
    $PathInput.Location = New-Object System.Drawing.Point(10,50)
    $PathInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($PathInput)

$PasswordInput = New-Object System.Windows.Forms.MaskedTextBox
    $PasswordInput.Location = New-Object System.Drawing.Point(10,130)
    $PasswordInput.Size = New-Object System.Drawing.Size(260,40)
    $PasswordInput.PasswordChar = '*'
    $Form2.Controls.Add($PasswordInput)

$PubPathInput = New-Object System.Windows.Forms.TextBox
    $PubPathInput.Location = New-Object System.Drawing.Point(10,210)
    $PubPathInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($PubPathInput)

$RootInput = New-Object System.Windows.Forms.TextBox
    $RootInput.Location = New-Object System.Drawing.Point(10,290)
    $RootInput.Size = New-Object System.Drawing.Size(260,40)
    $Form2.Controls.Add($RootInput)

#! *************************************** PROGRESS BAR***************************************!#

$progressBar1 = New-Object System.Windows.Forms.ProgressBar
$progressBar1.Name = 'progressBar1'
$progressBar1.Value = 0
$progressBar1.Style="Continuous"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = $width - 40
$System_Drawing_Size.Height = 20
$progressBar1.Size = $System_Drawing_Size
$progressBar1.Left = 5
$progressBar1.Top = 40

#! *************************************** BUTTONS ***************************************!#
$Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Size(8, 80)
    $Button.Size= New-Object System.Drawing.Size(240, 32)
    $Button.TextAlign = "MiddleCenter"
    $Button.Text = "Test the modules"
    $Button.add_Click({
        $check1 = $false
        $check2 = $false
        if (Get-Module -ListAvailable -Name MSOnline ) {
            $check1 = $true
        } 
        
        if (Get-Module -ListAvailable -Name MSOnlineExt ) {
            $check2 = $true
        } 

       if ($check1 -and $check2){
           $label_Form.Text = "It seems that both MSOnline and MSOnlineExt modules are installed."
           $Button.Text = "Next"
           $Button.add_Click({
               $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $false -and $check2 -eq $false){
	        $label_Form.Text = "We are installing both modules it will take a second."
       	    Install-Module -Name MSOnline
	        Install-Module -Name MSOnlineExt
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $false -and $check2 -eq $true){
            $label_Form.Text = "We are installing one module it will take a second."
            Install-Module -Name MSOnline
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }elseif($check1 -eq $true -and $check2 -eq $false){
            $label_Form.Text = "We are installing one module it will take a second."
            Install-Module -Name MSOnlineExt
            $label_Form.Text = "Installed."
            $Button.Text = "Next"
            $Button.add_Click({
                $Form2.Add_Shown({$Form2.Activate()})
                [void] $Form2.ShowDialog()
            })
       }
    })
    $Form.Controls.Add($Button)

$PathButton = New-Object System.Windows.Forms.Button
    $PathButton.Location = New-Object System.Drawing.Size(280, 50)
    $PathButton.Size= New-Object System.Drawing.Size(100, 20)
    $PathButton.TextAlign = "MiddleCenter"
    $PathButton.Text = "Open"
    $PathButton.add_Click({
        $PFXCertPath = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $false # Multiple files can be chosen
            Filter = 'Certificates (*.pfx)|*.pfx;' # Specified file types
        }
        [void]$PFXCertPath.ShowDialog()
        $PFXCertPath.SelectedPath
        [string]$FinalPath = $PFXCertPath."FileName"
        $PathInput.Text = $FinalPath
    })
    $Form2.Controls.Add($PathButton)

$PubPathButton = New-Object System.Windows.Forms.Button
    $PubPathButton.Location = New-Object System.Drawing.Size(280, 210)
    $PubPathButton.Size= New-Object System.Drawing.Size(100, 20)
    $PubPathButton.TextAlign = "MiddleCenter"
    $PubPathButton.Text = "Open"
    $PubPathButton.add_Click({
        $CerCertPath = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            Multiselect = $false # Multiple files can be chosen
            Filter = 'Certificates (*.cer)|*.cer;' # Specified file types
        }
        [void]$CerCertPath.ShowDialog()
        $CerCertPath.SelectedPath
        [string]$FinalPath = $CerCertPath."FileName"
        $PubPathInput.Text = $FinalPath
    })
    $Form2.Controls.Add($PubPathButton)

$Form2NextButton = New-Object System.Windows.Forms.Button
    $Form2NextButton.Location = New-Object System.Drawing.Size(150, 370)
    $Form2NextButton.Size= New-Object System.Drawing.Size(100, 30)
    $Form2NextButton.TextAlign = "MiddleCenter"
    $Form2NextButton.Text = "Next"
    $Form2NextButton.add_Click({
        $Form2.Controls.Add($progressBar1)
        #gets text from inputs and stored them in these variables
        $Password = $PasswordInput.Text
        $PrivateFinalPath = $PathInput.Text
        $PublicFinalPath = $PubPathInput.Text
        $RootDomain = $RootInput.Text
        #---------------------------------------------------

        #Pulls out the account that is running the MSCRMAsyncService and saves it on a variable
        $ServiceAccount = Get-WmiObject win32_service -filter "name= 'MSCRMAsyncService'" | Format-List StartName | Out-String 
        $ServiceAccount  = $ServiceAccount.Replace(‘StartName :’,  ’’)
        $ServiceAccount = $ServiceAccount.Trim()
        #--------------------------------------------------------------------------------

        if ('C:\Program Files\Microsoft Dynamics CRM\tools' | Test-Path) { 
            Set-Location 'C:\Program Files\Microsoft Dynamics CRM\tools'
        }elseif('C:\Program Files\Dynamics 365\Tools'  | Test-Path){
            Set-Location 'C:\Program Files\Dynamics 365\Tools' 
        }

        Try{
            #dot sources certificate reconfiguration and passes the parameters that were asked before
            . ".\CertificateReconfiguration.ps1"  -certificateFile $PrivateFinalPath -password $Password -certificateType S2STokenIssuer -updateCrm -serviceAccount $ServiceAccount -storeFindType FindBySubjectDistinguishedName
            #---------------------------------------------------------------------------------
            InitializingPSSession 
            ImportingModules
            GettingCredentials
            $CredValue = STSCertCreate -PrivateFinalPath $PrivateFinalPath -Password $Password -PublicFinalPath $PublicFinalPath
            RootDomainChanges -RootDomain1 $RootDomain -CredentialValue $CredValue
            AddingNewSettingToCloud 

        }   
        Catch{
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host "An error has occured $ErrorMessage  in $FailedItem" -ForegroundColor Red -BackgroundColor White
        }
    })
    $Form2.Controls.Add($Form2NextButton)


$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()
    $Form2.Controls.Add($Form2NextButton)


$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()