

# fetch differents links

$download = "Please wait while your download finishes (This could take up to an hour...)"
$complete = "Your download has finished"

$office2021professional = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2021Retail.img"
$office2021professionalplus = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img"
$office2021homestudent = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2021Retail.img"
$office2021homebusiness = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2021Retail.img"

$office2019professional = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/Professional2019Retail.img"
$office2019professionalplus = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img"
$office2019homestudent = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudent2019Retail.img"
$office2019homebusiness = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusiness2019Retail.img"

$office2016professional = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProfessionalRetail.img"
$office2016professionalplus = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlusRetail.img"
$office2016homestudent = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeStudentRetail.img"
$office2016homebusiness = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/HomeBusinessRetail.img"

$office365professionalplus = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/O365ProPlusRetail.img"
$office365business = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/fr-fr/O365BusinessRetail.img"
$office365homepremium = "https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/fr-fr/O365HomePremRetail.img"
# program start 

$validnumbers = @(1 ,2 ,3 ,4)
$validnumbers | ConvertTo-Json
Clear-Host
Write-Host ""
Write-Host "OfficeScript V.1.1"
Write-Host ""
Write-Host "Select your Microsoft Office edition (more will be added later!)"
Write-Host ""
Write-Host "1. Office 365"
Write-Host "2. Office 2021"
Write-Host "3. Office 2019"
Write-Host "4. Office 2016"
$officeedition = Read-Host "Number"

if ($validnumbers -contains $officeedition){
 if($officeedition -eq 1){
    Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 365 edition"
    Write-Host ""
    Write-Host "1. Office 365 Professional Plus"
    Write-Host "2. Office 365 Home Premium"
    Write-Host "3. Office 365 Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:" -ForegroundColor Green
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en" -ForegroundColor Green
    $office365edition = Read-Host "Number"

    if ($validnumbers -contains [int]$office365edition){

    if ($office365edition -eq 1){
    Write-Host ""
    Write-Host $download -ForegroundColor Cyan
    Invoke-WebRequest -Uri $office365professionalplus -OutFile "C:/office365professionalplus.img"
    Write-Host $complete -ForegroundColor Cyan} 

    elseif($office365edition -eq 2){
    Write-Host ""
    Wrte-Host $download -ForegroundColor Cyan
    Invoke-WebRequest -Uri $office365homepremium -OutFile "C:/office365homepremium.img"
    Write-Host $complete -ForegroundColor Cyan}

    elseif($office365edition -eq 3){
    Write-Host ""
    Wrte-Host $download -ForegroundColor Cyan
    Invoke-WebRequest -Uri $office365business -OutFile "C:/office365business.img"
    Write-Host $complete -ForegroundColor Cyan}

    }
    else {
    Write-Host "Wrong Number, Retry Again" -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit}
 }
 elseif ($officeedition -eq 2) {
    Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2021 edition"
    Write-Host ""
    Write-Host "1. Office 2021 Professional"
    Write-Host "2. Office 2021 Professional Plus"
    Write-Host "3. Office 2021 Home Student"
    Write-Host "4. Office 2021 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:" -ForegroundColor Green
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en" -ForegroundColor Green
    $office2021version = Read-Host "Number"
    if ($validnumbers -contains [int]$office2021version){
    if ($office2021version -eq 1) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2021professional -OutFile "C:/office2021professional.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2021version -eq 2) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2021professionalplus -OutFile "C:/office2021professionalplus.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2021version -eq 3) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2021homestudent -OutFile "C:/office2021homestudent.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2021version -eq 4) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2021homebusiness -OutFile "C:/office2021homebusiness.img"
        Write-Host $complete -ForegroundColor Cyan
    }
  }
    else {
    Write-Host "Wrong Number, Retry" -ForegroundColor Red
    Start-Sleep -Seconds 2
    exit }
}
 elseif ($officeedition -eq 3) {
   
    Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2019 edition"
    Write-Host ""
    Write-Host "1. Office 2019 Professional"
    Write-Host "2. Office 2019 Professional Plus"
    Write-Host "3. Office 2019 Home Student"
    Write-Host "4. Office 2019 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:" -ForegroundColor Green
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en" -ForegroundColor Green
    $office2019edition = Read-Host "Number"
    if ($validnumbers -contains [int]$office2019edition){
    if ($office2019edition -eq 1) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2019professional -OutFile "C:/office2019professional.img"
        Write-Host $complete -ForegroundColor Cyan
    } 
    elseif ($office2019edition -eq 2) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2019professionalplus -OutFile "C:/office2019professionalplus.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2019edition -eq 3) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2019homestudent -OutFile "C:/office2019homestudent.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2019edition -eq 4) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2019homebusiness -OutFile "C:/office2019homebusiness.img"
        Write-Host $complete -ForegroundColor Cyan
    }
  }
    else {
    Start-Sleep -Seconds 2
    Write-Host "Wrong Number, Retry Again"
    exit}
}
 elseif ($officeedition -eq 4) {
   
    Clear-Host
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2016 edition"
    Write-Host ""
    Write-Host "1. Office 2016 Professional"
    Write-Host "2. Office 2016 Professional Plus"
    Write-Host "3. Office 2016 Home Student"
    Write-Host "4. Office 2016 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:" -ForegroundColor Green
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en" -ForegroundColor Green
    $office2016edition = Read-Host "Number"
    if ($validnumbers -contains [int]$office2016edition){
    if ($office2016edition -eq 1) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2016professional -OutFile "C:/office2016professional.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2016edition -eq 2) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2016professionalplus -OutFile "C:/office2016professionalplus.img"
        Write-Host $complete -ForegroundColor Cyan
    }
     elseif ($office2016edition -eq 3) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2016homestudent -OutFile "C:/office2016homestudent.img"
        Write-Host $complete -ForegroundColor Cyan
    } 
     elseif ($office2016edition -eq 4) {
        Write-Host ""
        Write-Host $download -ForegroundColor Cyan
        Invoke-WebRequest -Uri $office2016homebusiness -OutFile "C:/office2016homebusiness.img"
        Write-Host $complete -ForegroundColor Cyan
    }
}
    else {
Write-Host "Wrong number, Retry Again" -ForegroundColor Red
Start-Sleep -Seconds 2
exit
}

    
}
}
else {
Write-Host "Wrong number, Retry Again" -ForegroundColor Red
Start-Sleep -Seconds 2
exit
}


# find the downloaded .img file
Clear-Host
$imagepath = "C:\"

$fileitem = Get-ChildItem -Path $imagepath -filter "office*"
$filename = $fileitem.Name

Write-Host ""
Write-host "The file name is: $filename"

# mount the downloaded image


Mount-DiskImage -ImagePath $imagepath\$filename


# find the drive letter in which the image is mounted
$udfvolume = Get-Volume | Where-Object FileSystem -eq "UDF" | Select-Object -First 1 -ExpandProperty DriveLetter

#check system architecture

$setup = ":/Office/Setup64"
$setup32 = ":/Office/Setup32"
Write-Host ""
Write-Host "Filename $filename mounted in $udfvolume volume"
Start-Sleep -Seconds 3
Write-Host ""
Write-Host "Is your computer 64bit or 32bit?"
Write-Host "1. 64bit"
Write-Host "2. 32bit"
$archi = Read-Host("Number")

# execute the correct file depending on the architecture
if ($archi -eq 1)
{$exepath = $udfvolume+$setup
Write-Host "Starting $exepath"}
elseif ($archi -eq 2){$exepath = $udfvolume+$setup32
Write-Host "Starting $exepath"}

Write-Host "Installing $filename"
Start-Process -FilePath $exepath -Wait

Clear-Host

# activate office, and windows if not activated
& ([ScriptBlock]::Create((Invoke-RestMethod https://get.activated.win))) /Ohook 

Clear-Host

# cleanup
Dismount-DiskImage -ImagePath $imagepath/$filename
Start-Sleep -Seconds 5
Clear-Host
Write-Host "Do you want to delete the original image?"
Write-Host "1. Yes"
Write-Host "2. No"
$cleanup = Read-Host "Number"
if ($validnumbers -contains [int]$cleanup){
if ($cleanup = 1){
Remove-Item -Path $imagepath/$filename}
if ($cleanup = 2){
Write-Host "Office .img file saved in $imagepath/$filename"}
Write-Host "Script Finished"
}
