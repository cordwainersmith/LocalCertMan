<#
                                                                                                                     
                                          ,,                                                                         
`7MMF'                                  `7MM   .g8"""bgd                   mm   `7MMM.     ,MMF'                     
  MM                                      MM .dP'     `M                   MM     MMMb    dPMM                       
  MM         ,pW"Wq.   ,p6"bo   ,6"Yb.    MM dM'       ` .gP"Ya `7Mb,od8 mmMMmm   M YM   ,M MM   ,6"Yb.  `7MMpMMMb.  
  MM        6W'   `Wb 6M'  OO  8)   MM    MM MM         ,M'   Yb  MM' "'   MM     M  Mb  M' MM  8)   MM    MM    MM  
  MM      , 8M     M8 8M        ,pm9MM    MM MM.        8M""""""  MM       MM     M  YM.P'  MM   ,pm9MM    MM    MM  
  MM     ,M YA.   ,A9 YM.    , 8M   MM    MM `Mb.     ,'YM.    ,  MM       MM     M  `YM'   MM  8M   MM    MM    MM  
.JMMmmmmMMM  `Ybmd9'   YMbmd'  `Moo9^Yo..JMML. `"bmmmd'  `Mbmmd'.JMML.     `Mbmo.JML. `'  .JMML.`Moo9^Yo..JMML  JMML.

version 0.0.7                                                                                                                     
                                                                                                                     

a PowerShell script that uses the Windows Forms framework to create a graphical user interface for performing different certificate management tasks on a loca Windows machine.

* Project page: 
https://github.com/cordwainersmith/LocalCertMan

* Features

- Displays details of all installed personal and root certificates, highlighting certificates that are about to expire.
- Exports certificate lists, including the DNS names and port bindings, to CSV files.
- Installs password-protected PFX personal certificates.
- Installs trusted root certificates to the local machine certificate store.
- Merges CER and KEY files to a password-protected PFX.
- Binds a certificate to the IIS default website.
- Views certificate details from a URL.

* License: MIT

Copyright (c) 2023 Liran Baba

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


* Prerequisites

- PowerShell version 5.1 or later must be installed on the system.
- The PowerShell WebAdministration module must be installed to allow SSL bindings.
- The PowerShell execution policy must be set to RemoteSigned or Unrestricted to run this tool.
- .NET Framework version 4.5 or later must be installed on the system.
- The tool needs access to the LocalMachine certificates store to be able to export certificates and add the root CA to the trust store, and should be run with administrative privileges.

* Contact
questions or issues, contact cordwainer@alpharalpha.net

#>
#########################################################################################################################################################################################
# Form Settings
#########################################################################################################################################################################################
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "LocalCertMan"
#$form.Size = New-Object System.Drawing.Size(1190, 1150)
$form.ClientSize = New-Object System.Drawing.Size(1190, 1080)
$form.ShowIcon = $false
$form.StartPosition = "CenterScreen"
# Set the form to preview key events
$form.KeyPreview = $true
# Add an event handler for the KeyDown event
$form.Add_KeyDown({
        # If the key pressed is the escape key, close the form
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $form.Close()
        }
    })

# Style settings
$font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$form.font = $font

#########################################################################################################################################################################################
# About menu
###############################################
$helpMenu = New-Object System.Windows.Forms.MenuItem
$helpMenu.Text = "About"
$version = '0.0.8'
$aboutMenuItem = New-Object System.Windows.Forms.MenuItem
$aboutMenuItem.Text = "About LocalCertMan"
$aboutMenuItem.add_Click({
        # Show a new window with the name Test when About is clicked
        $aboutForm = New-Object System.Windows.Forms.Form
        $aboutForm.Text = "About"
        $aboutForm.Size = New-Object System.Drawing.Size(470, 400)
        $aboutForm.StartPosition = "CenterScreen"
        $aboutForm.ShowIcon = $false
        $aboutForm.KeyPreview = $true

        # Add an event handler for the KeyDown event
        $aboutForm.Add_KeyDown({
                # If the key pressed is the escape key, close the form
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                    $aboutForm.Close()
                }
            })


        # Add a close button
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Close"
        $closeButton.Size = New-Object System.Drawing.Size(100, 30)
        $closeButton.Location = New-Object System.Drawing.Point(185, 300)
    
        $closeButton.add_Click({
                $aboutForm.Close()
            })
        $aboutForm.Controls.Add($closeButton)

        # Add a label 
        # Add a label with a clickable email field
        $label = New-Object System.Windows.Forms.Label
        $label.Text = @"
LocalCertMan
Version  $version
Display lists for Personal and Trusted Root CA Certifcates
Bind SSL/TLS certificates to Internet Information Services (IIS)
Install PFX (Personal Information Exchange) files and root Certificate Authority (CA) certificates in the local certificate store
Convert CER (Certificate) and KEY (Private Key) files to PFX format
Export Personal and Trusted root certificate lists to CSV files
View certificate details from URL
"@
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(110, 300)

        $email = "cordwainer@alpharalpha.net"
        $emailLink = New-Object System.Windows.Forms.LinkLabel
        $emailLink.Text = $email
        $emailLink.Location = New-Object System.Drawing.Point(150, 240)
        $emailLink.AutoSize = $true
        $emailLink.Add_Click({
                $mailto = "mailto:$email"
                Start-Process $mailto
            })
        $label.Controls.Add($emailLink)

        $label.Size = New-Object System.Drawing.Size(450, 300)
        $label.Location = New-Object System.Drawing.Point(1, 20)
        $label.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
        $label.AutoSize = $false
    
        $aboutForm.Controls.Add($label)

        $aboutForm.ShowDialog()
    })

$helpMenu.MenuItems.Add($aboutMenuItem)
# Add the menus to the main form
$form.Menu = New-Object System.Windows.Forms.MainMenu
$form.Menu.MenuItems.AddRange(@($helpMenu))
#########################################################################################################################################################################################
# Controls
#########################################################################################################################################################################################
#########################################################################################
# Controls: Buttons
# Get Url
###############################################

$GetUrlButton = New-Object System.Windows.Forms.Button
$GetUrlButton.Text = "Get Certificate from URL"
$GetUrlButton.Location = New-Object System.Drawing.Point(730, 862)
$GetUrlButton.Size = New-Object System.Drawing.Size(450, 25)


$textBoxURL = New-Object System.Windows.Forms.TextBox
$textBoxURL.Width = 500
$textBoxURL.Height = 30
$textBoxURL.Location = New-Object System.Drawing.Point(220, 864)
$textBoxURL.Text = "https://" # Set the text to "https://"


#########################################################################################
# Event handler for the GET certificate from URL 
#########################################################################################
$GetUrlButton.Add_Click({
        $url = $textBoxURL.Text
        try {
            $req = [Net.HttpWebRequest]::Create($url)
            $req.GetResponse()
            $cert = $req.ServicePoint.Certificate
            $issuer = $cert.Issuer
            $subject = $cert.Subject
            $notAfter = $cert.GetExpirationDateString()
            $notBefore = $cert.GetEffectiveDateString()
            $thumbprint = $cert.GetCertHashString()

            $log = "URL: $url`nSubject: $subject`nNot After: $notAfter`nNot Before: $notBefore`nIssuer: $issuer`nThumbprint: $thumbprint"
            #Write-Host $log
            #[System.Windows.Forms.MessageBox]::Show("Subject: $subject`nNot After: $notAfter`nNot Before: $notBefore`nIssuer: $issuer")
        
            $today = Get-Date
            $certNotAfter = [DateTime]::Parse($notAfter)
        
            if (($certNotAfter - $today).Days -le 90) {
                $startIndex = $textBoxURL.Text.IndexOf($notAfter)
                if ($startIndex -ge 0) {
                    $textBoxURL.Select($startIndex, $notAfter.Length)
                    $textBoxURL.SelectionColor = [System.Drawing.Color]::Red
                }
            }

            $newRow = $GetUrlGrid.Rows.Add()
            $GetUrlGrid.Rows[$newRow].Cells["URL"].Value = $url
            $GetUrlGrid.Rows[$newRow].Cells["Subject"].Value = $subject
            $GetUrlGrid.Rows[$newRow].Cells["NotAfter"].Value = $notAfter
            $GetUrlGrid.Rows[$newRow].Cells["Issuer"].Value = $issuer
            $GetUrlGrid.Rows[$newRow].Cells["Thumbprint"].Value = $thumbprint
        
        }
        catch [System.Net.WebException] {
            $message = "Cannot establish trust - ROOT CA is not trusted"
            Write-Host $message
            [System.Windows.Forms.MessageBox]::Show($message)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show($_.Exception.Message)
        }
    })

#########################################################################################
# Controls: Buttons
# Refresh
###############################################
$button = New-Object System.Windows.Forms.Button
$button.Text = "Refresh List"
$button.Location = New-Object System.Drawing.Size(10, 10)
$button.Size = New-Object System.Drawing.Size(150, 25)

#########################################################################################
# Event handler for Refresh button
#########################################################################################
$button.Add_Click({
        # Define a function to run when the button is clicked 
        $grid.Rows.Clear() 
        RefreshPersonalDataGridView
        RefreshRootDataGridView   
        RefreshIISBox
    })

#########################################################################################
# Controls: Buttons
# Import PFX
###############################################
$buttonimport = New-Object System.Windows.Forms.Button
$buttonimport.Location = New-Object System.Drawing.Size(180, 10)
$buttonimport.Size = New-Object System.Drawing.Size(150, 25)
$buttonimport.Text = "Import New PFX"
#########################################################################################
# Event handler for Import button
#########################################################################################
# Event handler for the import button in the second tab 
$buttonimport.Add_Click({
    
        # Display a new text box to input the password
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.PasswordChar = "*"
        $passwordTextBox.Width = 300
        $passwordTextBox.Location = New-Object System.Drawing.Point(10, 10)
        $form = New-Object System.Windows.Forms.Form
        $form.Controls.Add($passwordTextBox)
        $form.Text = "Enter PFX Password"
        $form.ShowIcon = $false 
        $form.Size = New-Object System.Drawing.Size(350, 120)
        $form.StartPosition = "CenterScreen"
        $form.TopMost = $true
        $form.Add_Shown({ $passwordTextBox.Select() })

        # Add an "OK" button to the form
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(130, 50)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)
        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Get the entered password
            $password = $passwordTextBox.Text

            # Display a file open dialog to select the certificate file
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "PFX Files (*.pfx)|*.pfx|All files (*.*)|*.*"
            $openFileDialog.Title = "Select Certificate File"
            $openFileDialog.ShowDialog() | Out-Null

            # Get the selected file path
            $filePath = $openFileDialog.FileName
            # Check if a file was selected
            if ($filePath) {
                # Import the certificate using the Import-PfxCertificate cmdlet and the entered password
                $certificate = $null
                Try {
                    $certificate = Import-PfxCertificate -FilePath $filePath -CertStoreLocation "Cert:\LocalMachine\My" -Password (ConvertTo-SecureString -String $password -AsPlainText -Force)
                    # Get the certificate thumbprint
                    #$thumbprint = $certificate.Thumbprint

                    # Get the issuer name of the certificate
                    $issuerName = $certificate.IssuerName.Name

                    # Check if the root CA certificate is installed on the local machine trusted root store
                    $rootCA = Get-ChildItem -Path Cert:\LocalMachine\Root | Where-Object { $_.Subject -eq $issuerName }

                    if ($rootCA) {
                        # If the root CA certificate is found, display a success message with ROOT CA issuer name and thumbprint
                        [System.Windows.Forms.MessageBox]::Show("Import successful! `nRoot CA Validation PASSED: The root CA certificate of your PFX is installed on this machine.`n`nCA thumbprint:$($rootCA.Thumbprint) `nIssuer:$($rootCA.Issuer) `n", "Import Certificate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        return
                    }
                    else {
                        # If the root CA certificate is not found, display a message to inform the user with ROOT CA issuer name and thumbprint
                        [System.Windows.Forms.MessageBox]::Show("Import successful! `n`n [Warning:`]`nRoot CA Validation FAILED: root CA with issuer name: $issuerName `CA thumbprint:$($rootCA.Thumbprint) is not installed on the local machine trusted root store.", "Import Certificate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        return
                    }
                }
                Catch {
                    # If the import fails, check the error message
                    if ($_.Exception.Message -like "*different password*") {
                        # If the error is "The PFX file you are trying to import requires a password", display a message to inform the user
                        [System.Windows.Forms.MessageBox]::Show("The private key password is incorrect. Please try again.", "Import Certificate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)                        
                        return
                    }
                    else {
                        # If the error is not recognized, display the full error message with ROOT CA issuer name and thumbprint
                        [System.Windows.Forms.MessageBox]::Show("$($_.Exception.Message) The root CA certificate with thumbprint $($rootCA.Thumbprint) and issuer name $($rootCA.Issuer) is installed on the local machine trusted root store.", "Import Certificate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        return
                    }
                }
                
            }
            RefreshRootDataGridView
            RefreshPersonalDataGridView
        }
    })


#########################################################################################
# Controls: Buttons
# Select Cert to Bind
###############################################
$selectButton = New-Object System.Windows.Forms.Button
$selectButton.Location = New-Object System.Drawing.Size(350, 10)
$selectButton.Size = New-Object System.Drawing.Size(150, 25)
$selectButton.Text = "Select Cert to Bind"
#########################################################################################
# Event handler for the select button choose a certificate from a list to bind to the IIS
#########################################################################################
$selectButton.Add_Click({
        try {
            # Get the certificates in the Local Machine Personal store
            $certificates = Get-ChildItem cert:\LocalMachine\my | Select-Object Subject, NotAfter, Issuer, HasPrivateKey, DnsNamelist
            # Display a list of the certificates to the user
            $selectedCertificate = $certificates | Out-GridView -Title "Select Certificate" -OutputMode Single

            # Check if a certificate was selected
            if ($selectedCertificate) {
                # Update the certificate label to show the selected certificate
                $certificateLabel.Text = $selectedCertificate.Subject
                # Enable the bind button
                $bindButton.Enabled = $true
            }
        }
        catch {
            # Show the error to the user
            [System.Windows.Forms.MessageBox]::Show("Error selecting the certificate: $($Error[0].Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 
        }
    })

########################################################################################
# Controls: Buttons
# Bind Button
###############################################
# Create the bind certificate button which is disabled by default, enabled when a certificate is selected
$bindButton = New-Object System.Windows.Forms.Button
$bindButton.Location = New-Object System.Drawing.Size(520, 10)
$bindButton.Size = New-Object System.Drawing.Size(150, 25)
$bindButton.Text = "Bind Certificate to IIS"
$bindButton.Enabled = $false
#########################################################################################
# Event handler for the bind button to bind a certificate to the IIS
#########################################################################################
$bindButton.Add_Click({
        try {
            $thumbprint = $selectedCertificate.Thumbprint
            $command = Get-Item cert:\LocalMachine\MY\
            # Import the WebAdministration module
            Import-Module WebAdministration

            # Remove any existing HTTPS bindings
            $site = Get-WebBinding -Name "Default Web Site" -Protocol "https"
            if ($site) {
                Remove-WebBinding -Name "Default Web Site" -Protocol "https"
            }

            # Find the selected certificate by subject
            $certificate = Get-ChildItem cert:\ -Recurse | Where-Object { $_.Subject -match $certificateLabel.Text } | Select-Object -First 1

            # Add a new HTTPS binding using the selected certificate
            $binding = New-WebBinding -Name "Default Web Site" -IPAddress "*" -Port 443 -Protocol "https"
            cd IIS:\SslBindings
            Remove-Item -Path "IIS:\SslBindings\0.0.0.0!443"
            New-Item -Path "IIS:\SslBindings\0.0.0.0!443" -Value $certificate

            # Display a message box to confirm the import and bind
            $message = "Certificate '$($certificate.Subject)' with thumbprint '$($certificate.Thumbprint)' bound successfully!"
            [System.Windows.Forms.MessageBox]::Show($message, "Bind Certificate", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            # Show the error to the user
            [System.Windows.Forms.MessageBox]::Show("Error binding the certificate: $($Error[0].Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) 
        }
        RefreshPersonalDataGridView
        RefreshIISBox })



########################################################################################
# Controls: Buttons
# Convert Cer and KEY to pfx Button
###############################################
$ConvertButton = New-Object System.Windows.Forms.Button
$ConvertButton.Location = New-Object System.Drawing.Size(690, 10)
$ConvertButton.Size = New-Object System.Drawing.Size(150, 25)
$ConvertButton.Text = "Convert CER to PFX"

#########################################################################################
# Event handler for the convert button to convert .cer and .key to pfx using certutil
#########################################################################################
$ConvertButton.Add_Click({
        # Display a new text box to input the password
        $passwordTextBox = New-Object System.Windows.Forms.TextBox
        $passwordTextBox.PasswordChar = "*"
        $passwordTextBox.Width = 320
        $passwordTextBox.Location = New-Object System.Drawing.Point(10, 90)
        $form = New-Object System.Windows.Forms.Form
        $form.Controls.Add($passwordTextBox)
        $form.Text = "Convert CER+KEY file to PFX"
        $form.ShowIcon = $false 
        $form.Size = New-Object System.Drawing.Size(380, 220)
        $form.StartPosition = "CenterScreen"
        $form.TopMost = $true
        $form.Add_Shown({ $passwordTextBox.Select() })
    
        # Add an "OK" button to the form
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(110, 130)
        $okButton.Size = New-Object System.Drawing.Size(135, 33)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $okButton
        $form.Controls.Add($okButton)
    
        # Create the details label in the new form for convert
        $convertlabel = New-Object System.Windows.Forms.Label
        $convertlabel.Location = New-Object System.Drawing.Size(10, 10)
        $convertlabel.Size = New-Object System.Drawing.Size(350, 200)              
        $convertlabel.Text = ".CER and .KEY files Must be placed in the same directory and MUST have the same base file name (file name excluding extension) provide a PFX password and click OK"
        
        $form.Controls.Add($convertlabel)  
    
    
        $result = $form.ShowDialog()
    
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # Get the entered password
            $password = $passwordTextBox.Text         
           
    
            # Show a file open dialog and set the text box to the selected file path
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "CER files (*.cer)|*.cer|All files (*.*)|*.*"
            $openFileDialog.Title = "Select CER file"
            $openFileDialog.ShowDialog() | Out-Null
       
            $cerFile = $openFileDialog.FileName
       
            # Get the directory of the CER file
            $cerDirectory = Split-Path $cerFile    
            
            # Get the name of the CER file without the directory or extension
            $cerName = (Split-Path -Leaf $cerFile)
    
            # Remove the .cer extension from the CER file name
            $cerName = $cerName.Substring(0, $cerName.Length - 4)
            # Set the PFX file path to the same directory as the CER file, with the name of the CER file and a .pfx extension
            $pfxFile = "$cerDirectory\$cerName.pfx"  
    
    
            # Run the certutil command to convert the CER file to a PFX file
            certutil -mergepfx -p "$password,$password" $cerFile $pfxFile     
    
            # Show a message box to indicate that the conversion was successful
            [System.Windows.Forms.MessageBox]::Show("Successfully generated PFX file: $pfxFile")
        }
        else {
            # Show a message box to indicate that the CER file and password fields are required
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid CER file path and password.")
        }
        
    })

########################################################################################
# Controls: Buttons
# Install Root cert
###############################################
$InstallRoot = New-Object System.Windows.Forms.Button
$InstallRoot.Location = New-Object System.Drawing.Size(860, 10)
$InstallRoot.Size = New-Object System.Drawing.Size(150, 25)
$InstallRoot.Text = "Install Root CA"
#########################################################################################
# Event handler for the installRoot button to install a root cert and refresh the root certs grid once done
#########################################################################################
$InstallRoot.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Certificate Files (*.pem;*.cer)|*.pem;*.cer"
        $dialog.Multiselect = $false
        $dialog.InitialDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($dialog.FileName)
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::Root, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
                $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
                $store.Add($cert)
                $store.Close()
                [System.Windows.Forms.MessageBox]::Show("Certificate installed successfully.", "Success")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Certificate installation failed, Are you sure this is a valid certificate file? :`n$($_.Exception.Message)", "Error")
            }
        }
        RefreshRootDataGridView })

########################################################################################
# Controls: Buttons
# Export to CSV
###############################################
$ExportButton = New-Object System.Windows.Forms.Button
$ExportButton.Location = New-Object System.Drawing.Size(1030, 10)
$ExportButton.Size = New-Object System.Drawing.Size(150, 25)
$ExportButton.Text = "Export to CSV"
#########################################################################################
# Event handler for the export button to export root and personal grid results to a csv files on the users desktop
#########################################################################################
$ExportButton.add_Click({
        # Export certificates and DNS names to a CSV file when Export is clicked
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $certs = $store.Certificates
        $store.Close()
        # Export root store certificates to a CSV file when Export is clicked
        $rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
        $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        $rootCerts = $rootStore.Certificates
        $rootStore.Close()
        $dateString = Get-Date -Format "yyyyMMdd_HHmmss"
        $hostname = $env:COMPUTERNAME
        $csvFile = "$env:USERPROFILE\Desktop\$dateString`_$hostname`_personal_certs.csv"
        $output = @()

        foreach ($cert in $certs) {
            $dnsNames = ""
            if ($cert.Extensions) {
                $extension = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
                if ($extension) {
                    $san = [System.Security.Cryptography.AsnEncodedData]::new($extension.Oid, $extension.RawData)
                    $sanValue = $san.Format($false)
                    $dnsNames = ($sanValue -split ",\s*") -replace "^DNS Name="
                    $dnsNames = $dnsNames -join ";"
                }
            }
            $output += [pscustomobject]@{
                Subject    = $cert.Subject
                Issuer     = $cert.Issuer
                NotBefore  = $cert.NotBefore
                NotAfter   = $cert.NotAfter
                Thumbprint = $cert.Thumbprint
                DnsNames   = $dnsNames
            }
        }

        $output | Export-Csv $csvFile -NoTypeInformation
    
        $csvFileRoot = "$env:USERPROFILE\Desktop\$dateString`_$hostname`_root_certs.csv"
        $outputRoot = @()

        foreach ($cert in $rootCerts) {
            $outputRoot += [pscustomobject]@{
                Subject      = $cert.Subject
                Issuer       = $cert.Issuer
                NotBefore    = $cert.NotBefore
                NotAfter     = $cert.NotAfter
                FriendlyName = $cert.FriendlyName
                Thumbprint   = $cert.Thumbprint
            }
        }

        $outputRoot | Export-Csv $csvFileRoot -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("certificates list exported successfully to: $csvFile, $csvFileRoot", "Export Complete")

    })

#########################################################################################
# Controls: Labels
#########################################################################################

$LabelGetfromURL = New-Object System.Windows.Forms.Label
$LabelGetfromURL.Location = New-Object System.Drawing.Size(10, 864)
$LabelGetfromURL.AutoSize = $true
$LabelGetfromURL.Text = "View certificate data from URL"



#########################################################################################
# Controls: Labels: personal and trusted root store titles
#########################################################################################

$labelPersonal = New-Object System.Windows.Forms.Label
$labelPersonal.Location = New-Object System.Drawing.Size(10, 180)
$labelPersonal.AutoSize = $true
$labelPersonal.Text = "Personal Certificates"

$labelRootStore = New-Object System.Windows.Forms.Label
$labelRootStore.Location = New-Object System.Drawing.Size(10, 520)
$labelRootStore.AutoSize = $true
$labelRootStore.Text = "Trusted Root Store Certificates"


#########################################################################################
# Controls: groupbox :groupbox1: Labels
#########################################################################################
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Location = New-Object System.Drawing.Size(10, 40)
$groupBox1.Size = New-Object System.Drawing.Size(1170, 130)
$form.Controls.Add($groupBox1)

$labelIIS = New-Object System.Windows.Forms.Label
$labelIIS.Location = New-Object System.Drawing.Size(10, 10)
$labelIIS.AutoSize = $true
$labelIIS.Text = "The Below Certificate is bound to the IIS"
$groupBox1.Controls.Add($labelIIS)

$labelIssuer = New-Object System.Windows.Forms.Label
$labelIssuer.Location = New-Object System.Drawing.Size(10, 35)
$labelIssuer.AutoSize = $true
$groupBox1.Controls.Add($labelIssuer)

$labelSubject = New-Object System.Windows.Forms.Label
$labelSubject.Location = New-Object System.Drawing.Size(10, 55)
$labelSubject.AutoSize = $true
$groupBox1.Controls.Add($labelSubject)

$labelThumbprint = New-Object System.Windows.Forms.Label
$labelThumbprint.Location = New-Object System.Drawing.Size(10, 75)
$labelThumbprint.AutoSize = $true
$groupBox1.Controls.Add($labelThumbprint)

$labelExpiration = New-Object System.Windows.Forms.Label
$labelExpiration.Location = New-Object System.Drawing.Size(10, 95)
$labelExpiration.AutoSize = $true
$groupBox1.Controls.Add($labelExpiration)

#########################################################################################
#Controls: Labels check password label
#########################################################################################
$passwordLabel = New-Object System.Windows.Forms.Label
$passwordLabel.Location = New-Object System.Drawing.Size(370, 130)
$passwordLabel.Size = New-Object System.Drawing.Size(420, 30)
$passwordLabel.Text = "** Enter the PFX password below before importing the file **"
$passwordLabel.ForeColor = [System.Drawing.Color]::Red
#########################################################################################
# Controls: dataGrids: get url cert grid
#########################################################################################
$GetUrlGrid = New-Object System.Windows.Forms.DataGridView
$GetUrlGrid.Width = 1170
$GetUrlGrid.Height = 160
$GetUrlGrid.Location = New-Object System.Drawing.Point(10, 900)
$GetUrlGrid.AllowUserToAddRows = $false
$GetUrlGrid.AllowUserToDeleteRows = $false
$GetUrlGrid.AllowUserToAddRows = $false
$GetUrlGrid.ReadOnly = $true
$GetUrlGrid.SelectionMode = "FullRowSelect"
$GetUrlGrid.MultiSelect = $false
$GetUrlGrid.RowHeadersVisible = $false
$GetUrlGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
#########################################################################################
# Controls: dataGrids: gridRoot for trusted root store certs
#########################################################################################
$gridRoot = New-Object System.Windows.Forms.DataGridView
$gridRoot.Location = New-Object System.Drawing.Size(10, 550)
$gridRoot.Size = New-Object System.Drawing.Size(1170, 300)
$gridRoot.AllowUserToAddRows = $false
$gridRoot.ReadOnly = $true
$gridRoot.SelectionMode = "FullRowSelect"
$gridRoot.MultiSelect = $false
$gridRoot.RowHeadersVisible = $false
$gridRoot.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
# Add the certificate properties as columns to the DataGridView
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Expiry Date"
$column.Name = "ExpiryDate"
$gridRoot.Columns.Add($column)

$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Subject"
$column.Name = "Subject"
$gridRoot.Columns.Add($column)

$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "FriendlyName" 
$column.Name = "FriendlyName"
$gridRoot.Columns.Add($column)

$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Issuer"
$column.Name = "Issuer"
$gridRoot.Columns.Add($column)
 
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Thumbprint"
$column.Name = "Thumbprint"
$gridRoot.Columns.Add($column)
#get all root cert from root store with the below values 
$Rootcertificates = Get-ChildItem cert:\LocalMachine\Root | Select-Object Subject, NotAfter, Issuer, FriendlyName, Thumbprint
# Add the certificates as rows to the DataGridView
foreach ($certificate in $Rootcertificates) {
    $row = New-Object System.Windows.Forms.DataGridViewRow
    $row.CreateCells($gridRoot)
    $row.Cells[0].Value = $certificate.NotAfter
    $row.Cells[1].Value = $certificate.Subject
    $row.Cells[2].Value = $certificate.FriendlyName
    $row.Cells[3].Value = $certificate.Issuer
    $row.Cells[4].Value = $certificate.thumbprint
    $now = (Get-Date).ToUniversalTime()
    $expiry = $certificate.NotAfter.ToUniversalTime()
    $timeLeft = $expiry - $now
    if ($timeLeft.TotalDays -lt 90) {
        # Set the expiry date cell's foreground color to red
        $row.Cells[0].Style.ForeColor = [System.Drawing.Color]::Red
    }
    $gridRoot.Rows.Add($row)
}
#########################################################################################
# Controls: dataGrids: datagrid for personal root store certs
#########################################################################################
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Size(10, 210)
$grid.Size = New-Object System.Drawing.Size(1170, 300)
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
# Add the certificate properties as columns to the DataGridView
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "HTTPS Bound"
$column.Name = "IisHttpsBinding"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Subject"
$column.Name = "Subject"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Expiry Date"
$column.Name = "ExpiryDate"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Issuer"
$column.Name = "Issuer"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Has Private Key"
$column.Name = "HasPrivateKey"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Thumbprint"
$column.Name = "Thumbprint"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Self Signed?"
$column.Name = "IsSelfSigned"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "Binding Port"
$column.Name = "BindingPort"
$grid.Columns.Add($column)
$column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$column.HeaderText = "DNS Names"
$column.Name = "DNS Names"
$grid.Columns.Add($column)
# Get the SSL certificates from personal store
$certificates = Get-ChildItem cert:\LocalMachine\my 
# Get the SSL certificate thumbprint for the IIS binding
$bindings = Get-WebBinding -Protocol "https" | Select-Object -ExpandProperty certificatehash 
# Get the SSL certificates and binding ports
$sslCerts = netsh http show sslcert
$certPortMap = @{}
foreach ($sslCert in $sslCerts) {
    if ($sslCert -match '^\s+IP:Port\s+:\s+(.*):(\d+)$') {
        $ip = $Matches[1]
        $port = $Matches[2]
    }
    if ($sslCert -match '^\s+Certificate Hash\s+:\s+(.*)$') {
        $thumbprint = $Matches[1]
        $certPortMap[$thumbprint] = $port
    }
}
# Add the certificates as rows to the DataGridView
foreach ($certificate in $certificates) {
    $row = New-Object System.Windows.Forms.DataGridViewRow
    $row.CreateCells($grid)
    $row.Cells[0].Value = if ($bindings -contains $certificate.thumbprint) { "Yes" } else { "No" }
    $row.Cells[1].Value = $certificate.Subject
    $row.Cells[2].Value = $certificate.NotAfter
    $row.Cells[3].Value = $certificate.Issuer
    $row.Cells[4].Value = $certificate.HasPrivateKey
    $row.Cells[5].Value = $certificate.thumbprint
    $now = (Get-Date).ToUniversalTime()
    $expiry = $certificate.NotAfter.ToUniversalTime()
    $timeLeft = $expiry - $now
    if ($timeLeft.TotalDays -lt 90) {
        # Set the expiry date cell's foreground color to red
        $row.Cells[2].Style.ForeColor = [System.Drawing.Color]::Red
    }
    $row.Cells[6].Value = if ($certificate.Subject -contains $certificate.Issuer) { "Yes" } else { "No" }
    $bindingPort = ""
    if ($certPortMap.ContainsKey($certificate.Thumbprint)) {
        $bindingPort = $certPortMap[$certificate.Thumbprint]
    }
    $row.Cells[7].Value = $bindingPort
    $dnsNames = ""
    if ($certificate.Extensions) {
        $extension = $certificate.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
        if ($extension) {
            $san = [System.Security.Cryptography.AsnEncodedData]::new($extension.Oid, $extension.RawData)
            $sanValue = $san.Format($false)
            $dnsNames = ($sanValue -split ",\s*") -replace "^DNS Name="
            $dnsNames = $dnsNames -join ";"
        }
    }
    $row.Cells[8].Value = $dnsNames
    $grid.Rows.Add($row)
}


#########################################################################################################################################################################################
# Functions
#########################################################################################################################################################################################

########################################################################################
# Function to refresh the Root certs store
###############################################
function RefreshRootDataGridView {
    $gridRoot.Rows.Clear()
    $Rootcertificates = Get-ChildItem cert:\LocalMachine\Root | Select-Object Subject, NotAfter, Issuer, FriendlyName, Thumbprint
    foreach ($certificate in $Rootcertificates) {
        $row = New-Object System.Windows.Forms.DataGridViewRow
        $row.CreateCells($gridRoot)
        $row.Cells[0].Value = $certificate.NotAfter
        $row.Cells[1].Value = $certificate.Subject
        $row.Cells[2].Value = $certificate.FriendlyName
        $row.Cells[3].Value = $certificate.Issuer
        $row.Cells[4].Value = $certificate.thumbprint
        $now = (Get-Date).ToUniversalTime()
        $expiry = $certificate.NotAfter.ToUniversalTime()
        $timeLeft = $expiry - $now
        if ($timeLeft.TotalDays -lt 90) {
            # Set the expiry date cell's foreground color to red
            $row.Cells[0].Style.ForeColor = [System.Drawing.Color]::Red
        }
        $gridRoot.Rows.Add($row)
    }
}
# Call the Refresh-DataGridView function to populate the DataGridView with data
RefreshRootDataGridView

########################################################################################
# Function to refresh the personal certs store
###############################################
function RefreshPersonalDataGridView {
    $grid.Rows.Clear()
    $certificates = Get-ChildItem cert:\LocalMachine\my 
    $bindings = Get-WebBinding -Protocol "https" | Select-Object -ExpandProperty certificatehash 
    $sslCerts = netsh http show sslcert
    $certPortMap = @{}
    foreach ($sslCert in $sslCerts) {
        if ($sslCert -match '^\s+IP:Port\s+:\s+(.*):(\d+)$') {
            $ip = $Matches[1]
            $port = $Matches[2]
        }
        if ($sslCert -match '^\s+Certificate Hash\s+:\s+(.*)$') {
            $thumbprint = $Matches[1]
            $certPortMap[$thumbprint] = $port
        }
    }
    foreach ($certificate in $certificates) {
        $row = New-Object System.Windows.Forms.DataGridViewRow
        $row.CreateCells($grid)
        $row.Cells[0].Value = if ($bindings -contains $certificate.thumbprint) { "Yes" } else { "No" }
        $row.Cells[1].Value = $certificate.Subject
        $row.Cells[2].Value = $certificate.NotAfter
        $row.Cells[3].Value = $certificate.Issuer
        $row.Cells[4].Value = $certificate.HasPrivateKey
        $row.Cells[5].Value = $certificate.thumbprint
        $now = (Get-Date).ToUniversalTime()
        $expiry = $certificate.NotAfter.ToUniversalTime()
        $timeLeft = $expiry - $now
        if ($timeLeft.TotalDays -lt 90) {
            # Set the expiry date cell's foreground color to red
            $row.Cells[2].Style.ForeColor = [System.Drawing.Color]::Red
        }
        $row.Cells[6].Value = if ($certificate.Subject -contains $certificate.Issuer) { "Yes" } else { "No" }
        $bindingPort = ""
        if ($certPortMap.ContainsKey($certificate.Thumbprint)) {
            $bindingPort = $certPortMap[$certificate.Thumbprint]
        }
        $row.Cells[7].Value = $bindingPort
        $dnsNames = ""
        if ($certificate.Extensions) {
            $extension = $certificate.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
            if ($extension) {
                $san = [System.Security.Cryptography.AsnEncodedData]::new($extension.Oid, $extension.RawData)
                $sanValue = $san.Format($false)
                $dnsNames = ($sanValue -split ",\s*") -replace "^DNS Name="
                $dnsNames = $dnsNames -join ";"
            }
        }
        $row.Cells[8].Value = $dnsNames
        $grid.Rows.Add($row)
    }
}
# Call the Refresh-DataGridView function to populate the DataGridView with data
RefreshPersonalDataGridView
########################################################################################
# Function to refresh the IIS bound box
###############################################
function RefreshIISBox {
    $binding = Get-WebBinding -Port 443 | Where-Object { $_.Protocol -eq "https" }
    if ($binding) {
        $certificateThumbprint = $binding.CertificateHash
        $cert = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { $_.Thumbprint -eq $certificateThumbprint }
        if ($cert) {
            $labelIssuer.Text = "Issued By: $($cert.issuer)"
            $labelSubject.Text = "Subject: $($cert.Subject)"
            $labelThumbprint.Text = "Thumbprint: $($cert.Thumbprint)"
            $labelExpiration.Text = "Expiration Date: $($cert.NotAfter)"            
            # Check if the expiration date is within 90 days
            $ninetyDays = (Get-Date).AddDays(90)
            if ($cert.NotAfter -le $ninetyDays) {
                # Set the text color to red
                $labelExpiration.ForeColor = [System.Drawing.Color]::Red
                $labelExpiration.Text = "Expiration Date within 90 days: $($cert.NotAfter)"
            }
        }
        else {
            $labelSubject.Text = ""
            $labelThumbprint.Text = ""
            $labelExpiration.Text = "No certificate found with thumbprint $certificateThumbprint."
        }
    }
    else {
        $labelSubject.Text = ""
        $labelThumbprint.Text = ""
        $labelExpiration.Text = "No HTTPS binding found on port 443."
    }
}
# Call the RefreshIISBox function to populate the labels with data
RefreshIISBox
########################################################################################
# Controls: add form controls
###############################################
$form.Controls.AddRange(($LabelGetfromURL, $textBoxURL, $GetUrlGrid, $GetUrlButton, $labelPersonal, $labelRootStore, $gridRoot, $ExportButton, $InstallRoot, $groupBox1, $button, $grid, $buttonimport, $selectButton, $ConvertButton, $bindButton, $ImportpasswordTextBox))
# Show form
$form.ShowDialog() | Out-Null