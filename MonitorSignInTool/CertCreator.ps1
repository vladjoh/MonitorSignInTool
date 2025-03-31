# Create C:\temp folder if it doesn't exist (skip this if you dont want to create folder with # )
if (-not (Test-Path -Path "C:\temp")) {
    New-Item -ItemType Directory -Path "C:\temp"
}

# Define certificate parameters 
$certSubject = "CN=SignInMonitorCert"
$certStore = "Cert:\LocalMachine\My"
$certPath = "C:\temp\SignInMonitorCert.pfx" #Edit folder paths if you want to use another locations
$cerPublicPath = "C:\temp\SignInMonitorCert.cer" #Edit folder paths if you want to use another locations
$certPassword = ConvertTo-SecureString -String "YOUR STRONG PASSWORD" -Force -AsPlainText #Create strong password and copy paste it in "YOUR STRONG PASSWORD"

# Create self-signed certificate
$cert = New-SelfSignedCertificate `
    -Subject $certSubject `
    -CertStoreLocation $certStore `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(1) #Edit AddYears Value if you want

# Get the thumbprint
$thumbprint = $cert.Thumbprint
Write-Host "Certificate created with thumbprint: $thumbprint"
Write-Host "Use this thumbprint in your monitoring script."

# Export the certificate with private key to PFX file
Export-PfxCertificate -Cert "$certStore\$thumbprint" -FilePath $certPath -Password $certPassword
Write-Host "Certificate with private key exported to: $certPath"

# Export the certificate public key to CER file (for Entra ID)
Export-Certificate -Cert "$certStore\$thumbprint" -FilePath $cerPublicPath -Type CERT
Write-Host "Certificate public key exported to: $cerPublicPath"

Write-Host "`n==== NEXT STEPS ====="
Write-Host "1. Manually upload the certificate file $cerPublicPath to your Entra ID application"
Write-Host "2. Update your monitoring script with this certificate thumbprint: $thumbprint"
Write-Host "3. Run the monitoring script to test certificate authentication"
Write-Host "4. Remember to restrict access to certificates private key and password and your script file"
Write-Host "5. Don't forget about expiration date of the certificate, create reminder in your calendar, follow the same steps to renew it"
Write-Host "6. Create and setup scheduled task on your endpoint to run this script every X minutes/hours to check sign ins on your specific account/ accounts"