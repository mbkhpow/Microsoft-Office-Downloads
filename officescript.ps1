﻿$download = "Please wait while your download finishes"
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

Clear-Host
Write-Host ""
Write-Host "OfficeScript V.1.0"
Write-Host ""
Write-Host "Select your Microsoft Office edition (more will be added later!)"
Write-Host ""
Write-Host "1. Office 2021"
Write-Host "2. Office 2019"
Write-Host "3. Office 2016"
$officeedition = Read-Host "Number"

if ($officeedition -eq 1) {
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2021 edition"
    Write-Host ""
    Write-Host "1. Office 2021 Professional"
    Write-Host "2. Office 2021 Professional Plus"
    Write-Host "3. Office 2021 Home Student"
    Write-Host "4. Office 2021 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:"
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en"
    $office2021version = Read-Host "Number"

    if ($office2021version -eq 1) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2021professional -OutFile "C:/office2021professional.img"
        Write-Host $complete
    }
     elseif ($office2021version -eq 2) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2021professionalplus -OutFile "C:/office2021professionalplus.img"
        Write-Host $complete
    }
     elseif ($office2021version -eq 3) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2021homestudent -OutFile "C:/office2021homestudent.img"
        Write-Host $complete
    }
     elseif ($office2021version -eq 4) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2021homebusiness -OutFile "C:/office2021homebusiness.img"
        Write-Host $complete
    }
} elseif ($officeedition -eq 2) {
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2019 edition"
    Write-Host ""
    Write-Host "1. Office 2019 Professional"
    Write-Host "2. Office 2019 Professional Plus"
    Write-Host "3. Office 2019 Home Student"
    Write-Host "4. Office 2019 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:"
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en"
    $office2019edition = Read-Host "Number"

    if ($office2019edition -eq 1) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2019professional -OutFile "C:/office2019professional.img"
        Write-Host $complete
    } elseif ($office2019edition -eq 2) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2019professionalplus -OutFile "C:/office2019professionalplus.img"
        Write-Host $complete
    } elseif ($office2019edition -eq 3) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2019homestudent -OutFile "C:/office2019homestudent.img"
        Write-Host $complete
    } elseif ($office2019edition -eq 4) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2019homebusiness -OutFile "C:/office2019homebusiness.img"
        Write-Host $complete
    }
} elseif ($officeedition -eq 3) {
    Write-Host ""
    Write-Host ""
    Write-Host "Select Your Office 2016 edition"
    Write-Host ""
    Write-Host "1. Office 2016 Professional"
    Write-Host "2. Office 2016 Professional Plus"
    Write-Host "3. Office 2016 Home Student"
    Write-Host "4. Office 2016 Home Business"
    Write-Host ""
    Write-Host "For a comparison between these versions go to:"
    Write-Host "https://www.licencedeals.com/blogs/licencedeals-info-corner/microsoft-office-editions-comparison?lang=en"
    $office2016edition = Read-Host "Number"

    if ($office2016edition -eq 1) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2016professional -OutFile "C:/office2016professional.img"
        Write-Host $complete
    }
     elseif ($office2016edition -eq 2) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2016professionalplus -OutFile "C:/office2016professionalplus.img"
        Write-Host $complete
    }
     elseif ($office2016edition -eq 3) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2016homestudent -OutFile "C:/office2016homestudent.img"
        Write-Host $complete
    } 
     elseif ($office2016edition -eq 4) {
        Write-Host ""
        Write-Host $download
        Invoke-WebRequest -Uri $office2016homebusiness -OutFile "C:/office2016homebusiness.img"
        Write-Host $complete
    }
}
$fileitem = Get-ChildItem -Path "C:\" -filter "office*"
$filename = $fileitem.Name
Write-Host ""
Write-host "The file name is: $filename"
$imagepath = "C:\"
Mount-DiskImage -ImagePath $imagepath\$filename
$udfvolume = Get-Volume | Where-Object FileSystem -eq "UDF" | Select-Object -First 1 -ExpandProperty DriveLetter
$setup = ":/Office/Setup64"
Start-Sleep -Seconds 5
$exepath = $udfvolume+$setup
Write-Host $exepath
Start-Process -FilePath $exepath
Dismount-DiskImage -ImagePath $imagepath\$filepath