$exePath = ".\DbfProtocolsParser\bin\Debug\net8.0\win-x86\publish\DbfProtocolsParser.exe"
$pfxPath = ".\CodeSigningRootCA.pfx"

$cert = Get-PfxCertificate -FilePath $pfxPath -Password (ConvertTo-SecureString -String "password" -AsPlainText -Force)
Set-AuthenticodeSignature -FilePath $exePath -Certificate $cert
