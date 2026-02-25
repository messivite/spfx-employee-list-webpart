# PnP PowerShell ile "Çalışanlar" SharePoint listesi oluşturur
# Kullanım: .\Create-EmployeesList.ps1 -SiteUrl "https://mustafaaksoy.sharepoint.com/sites/SiteAdi"
# Gereksinim: Install-Module PnP.PowerShell

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl
)

Connect-PnPOnline -Url $SiteUrl -Interactive

$listName = "Çalışanlar"

# Liste zaten varsa uyar
$existingList = Get-PnPList -Identity $listName -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host "Liste '$listName' zaten mevcut." -ForegroundColor Yellow
    exit
}

Write-Host "Liste oluşturuluyor: $listName" -ForegroundColor Cyan
New-PnPList -Title $listName -Template GenericList -OnQuickLaunch

# Varsayılan Title kolonunu "Ad" olarak yeniden adlandırmak zor; yeni kolonlar ekliyoruz
# Ad (Text)
Add-PnPField -List $listName -DisplayName "Ad" -InternalName "Ad" -Type Text -AddToDefaultView
# Soyad (Text)
Add-PnPField -List $listName -DisplayName "Soyad" -InternalName "Soyad" -Type Text -AddToDefaultView
# Departman (Choice veya Text)
Add-PnPField -List $listName -DisplayName "Departman" -InternalName "Departman" -Type Text -AddToDefaultView
# E-posta (Text - Email formatı)
Add-PnPField -List $listName -DisplayName "E-posta" -InternalName "Eposta" -Type Text -AddToDefaultView
# Fotoğraf (URL)
Add-PnPField -List $listName -DisplayName "Fotoğraf" -InternalName "Fotograf" -Type URL -AddToDefaultView
# JobTitle (mevcut listeyle uyumluluk için)
Add-PnPField -List $listName -DisplayName "Görev / Unvan" -InternalName "JobTitle" -Type Text -AddToDefaultView
# Person (Kullanıcı seçici)
Add-PnPField -List $listName -DisplayName "Personel" -InternalName "Person" -Type User -AddToDefaultView

Write-Host "Liste '$listName' basariyla olusturuldu." -ForegroundColor Green
Write-Host "Kolonlar: Ad, Soyad, Departman, E-posta, Fotograf, JobTitle, Person" -ForegroundColor Gray
