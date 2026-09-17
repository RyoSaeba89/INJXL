# Signe INJXL.exe avec le certificat de signature de code du poste.
#
#   powershell -ExecutionPolicy Bypass -File signer.ps1
#
# A relancer apres CHAQUE recompilation : PyInstaller reecrit l'exe et la
# signature precedente disparait. build_exe.bat l'appelle automatiquement.
#
# POURQUOI LE CERTIFICAT « XLDiff » PAR DEFAUT. C'est celui qui signe deja
# CONDA : sa cle publique est donc deja installee la ou ces outils tournent,
# alors qu'un certificat propre a INJXL demanderait une installation de plus
# avant que la moindre machine lui fasse confiance.
param(
  [string]$Exe = "$PSScriptRoot\dist\INJXL.exe",
  # empreinte du certificat a utiliser ; a defaut, le premier certificat de
  # signature de code trouve dans le magasin personnel
  [string]$Thumbprint = "634E5F62BF771BECE3F3FB818E45CF3081E939D1",
  [string]$TimestampServer = "http://timestamp.digicert.com"
)

if (-not (Test-Path $Exe)) {
  Write-Host "Introuvable : $Exe" -ForegroundColor Red
  Write-Host "Compile d'abord avec build_exe.bat."
  exit 1
}

$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Thumbprint -eq $Thumbprint -and $_.HasPrivateKey }
if (-not $cert) {
  $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
          Where-Object { $_.HasPrivateKey } | Select-Object -First 1
}
if (-not $cert) {
  Write-Host "Aucun certificat de signature de code avec cle privee dans Cert:\CurrentUser\My" -ForegroundColor Red
  Write-Host "Pour en creer un (auto-signe, valable 10 ans) :"
  Write-Host '  New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=INJXL - Prenom Nom, O=Organisation" -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(10)'
  exit 1
}

Write-Host "Certificat : $($cert.Subject)"
Write-Host "Expire le  : $($cert.NotAfter)"

# l'horodatage garde la signature valide apres l'expiration du certificat
$res = Set-AuthenticodeSignature -FilePath $Exe -Certificate $cert `
        -HashAlgorithm SHA256 -TimestampServer $TimestampServer -IncludeChain All

$sig = Get-AuthenticodeSignature $Exe
if (-not $sig.SignerCertificate) {
  Write-Host "ECHEC de la signature : $($res.StatusMessage)" -ForegroundColor Red
  exit 1
}

Write-Host ""
Write-Host "Signe          : $Exe"
Write-Host "Signataire     : $($sig.SignerCertificate.Subject)"
if ($sig.TimeStamperCertificate) {
  Write-Host "Horodate par   : $($sig.TimeStamperCertificate.Subject)"
} else {
  Write-Host "Horodatage     : ABSENT (pas d'acces reseau ?) - la signature expirera avec le certificat" -ForegroundColor Yellow
}
Write-Host "Verification   : $($sig.Status)"

if ($sig.Status -ne "Valid") {
  Write-Host ""
  Write-Host "Le fichier EST signe, mais ce poste ne fait pas confiance au certificat." -ForegroundColor Yellow
  Write-Host "Normal pour un certificat auto-signe : voir README.md, section Signature." -ForegroundColor Yellow
}
