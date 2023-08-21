<#PSScriptInfo
 
.VERSION 2.1
 
.GUID f95d9be8-dfdc-4c5e-8bc9-e06ce585e830
 
.AUTHOR Alexander Baker
 
.COMPANYNAME
 
.COPYRIGHT
 
.TAGS
 
.LICENSEURI
 
.PROJECTURI https://github.com/alex6851/BulkUpload-ADandEXOUserPhotos/tree/main
 
.ICONURI
 
.EXTERNALMODULEDEPENDENCIES
 
.REQUIREDSCRIPTS
 
.EXTERNALSCRIPTDEPENDENCIES
 
.RELEASENOTES
 
 
#>

<#
 
.DESCRIPTION
 This is for bulk uploading user photos to AD and Exchange Online.
 It will also look for an AD user account based on the Name of the picture.

 It will look for an employeeID in the file name, then it will look for a first/lastname, and then
 finally it will try to create a user's samaccountname based on what it finds in the filename.

 If it still does not find a user then it will give you chance to manually look for the user in AD and then you can put in the username of the user.

 In our situation we had 2 different M365 tenants but 1 on prem active directory. We would use extensionattribute15 to designate which
 tenant the user belonged in. So I use the same attribute to determine which tenant to upload the photo to. 

 for example: extensionattribute15 could equal 'GCCHighTenant' or it could equal 'GlobalTenant'
 
#> 

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$EXOModule = Get-InstalledModule "ExchangeOnlineManagement" -ErrorAction SilentlyContinue

if(!($EXOModule)){
    Install-Module "ExchangeOnlineManagement"
}
Import-Module "ExchangeOnlineManagement"

#This function is for resizing the images
Function Set-ImageSize {
	<#
	.SYNOPSIS
	    Resize image file.

	.DESCRIPTION
	    The Set-ImageSize cmdlet to set new size of image file.
		
	.PARAMETER Image
	    Specifies an image file. 

	.PARAMETER Destination
	    Specifies a destination of resized file. Default is current location (Get-Location).
	
	.PARAMETER WidthPx
	    Specifies a width of image in px. 
		
	.PARAMETER HeightPx
	    Specifies a height of image in px.		
	
	.PARAMETER DPIWidth
	    Specifies a vertical resolution. 
		
	.PARAMETER DPIHeight
	    Specifies a horizontal resolution.	
		
	.PARAMETER Overwrite
	    Specifies a destination exist then overwrite it without prompt.
	.PARAMETER MakePicturesSmaller
	    Will make pictures Smaller but not larger. 	 
		
	.PARAMETER FixedSize
	    Set fixed size and do not try to scale the aspect ratio. 

	.PARAMETER RemoveSource
	    Remove source file after conversion. 
		
	.EXAMPLE
		PS C:\> Get-ChildItem 'P:\test\*.jpg' | Set-ImageSize -Destination "p:\test2" -WidthPx 300 -HeightPx 375 -Verbose
		VERBOSE: Image 'P:\test\00001.jpg' was resize from 236x295 to 300x375 and save in 'p:\test2\00001.jpg'
		VERBOSE: Image 'P:\test\00002.jpg' was resize from 236x295 to 300x375 and save in 'p:\test2\00002.jpg'
		VERBOSE: Image 'P:\test\00003.jpg' was resize from 236x295 to 300x375 and save in 'p:\test2\00003.jpg'
		
	.NOTES
		Author: Michal Gajda
		Blog  : http://commandlinegeeks.com/
	#>
	[CmdletBinding(
		SupportsShouldProcess = $True,
		ConfirmImpact = "Low"
	)]		
	Param
	(
		[parameter(Mandatory = $true,
			ValueFromPipeline = $true,
			ValueFromPipelineByPropertyName = $true)]
		[Alias("Image")]	
		[String[]]$FullName,
		[String]$Destination = $(Get-Location),
		[bool]$Overwrite,
		[Int]$WidthPx,
		[Int]$HeightPx,
		[Int]$DPIWidth,
		[Int]$DPIHeight,
		[bool]$MakePicturesSmaller,
		[Switch]$FixedSize,
		[Switch]$RemoveSource
	)

	Begin {
		[void][reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
		#[void][reflection.assembly]::loadfile( "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Drawing.dll")
	}
	
	Process {

		Foreach ($ImageFile in $FullName) {
			If (Test-Path $ImageFile) {
				$OldImage = new-object System.Drawing.Bitmap $ImageFile
				$OldWidth = $OldImage.Width
				$OldHeight = $OldImage.Height
				
				if ($WidthPx -eq $Null) {
					$WidthPx = $OldWidth
				}
				if ($HeightPx -eq $Null) {
					$HeightPx = $OldHeight
				}
				if ($MakePicturesSmaller) {
					if ($OldHeight -lt $HeightPx) {
						$HeightPx = $OldHeight
					}
					if ($OldWidth -lt $WidthPx) {
						$WidthPx = $OldWidth
					}
				}
				if ($FixedSize) {
					$NewWidth = $WidthPx
					$NewHeight = $HeightPx
				}
				else {
					if ($OldWidth -lt $OldHeight) {
						$NewWidth = $WidthPx
						[int]$NewHeight = [Math]::Round(($NewWidth * $OldHeight) / $OldWidth)
						
						if ($NewHeight -gt $HeightPx) {
							$NewHeight = $HeightPx
							[int]$NewWidth = [Math]::Round(($NewHeight * $OldWidth) / $OldHeight)
						}
					}
					else {
						$NewHeight = $HeightPx
						[int]$NewWidth = [Math]::Round(($NewHeight * $OldWidth) / $OldHeight)
						
						if ($NewWidth -gt $WidthPx) {
							$NewWidth = $WidthPx
							[int]$NewHeight = [Math]::Round(($NewWidth * $OldHeight) / $OldWidth)
						}						
					}
				}

				$ImageProperty = Get-ItemProperty $ImageFile				
				$SaveLocation = Join-Path -Path $Destination -ChildPath ($ImageProperty.Name)

				If (!$Overwrite) {
					If (Test-Path $SaveLocation) {
						$Title = "A file already exists: $SaveLocation"
							
						$ChoiceOverwrite = New-Object System.Management.Automation.Host.ChoiceDescription "&Overwrite"
						$ChoiceCancel = New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel"
						$Options = [System.Management.Automation.Host.ChoiceDescription[]]($ChoiceCancel, $ChoiceOverwrite)		
						If (($host.ui.PromptForChoice($Title, $null, $Options, 1)) -eq 0) {
							Write-Verbose "Image '$ImageFile' exist in destination location - skiped"
							Continue
						} #End If ($host.ui.PromptForChoice($Title, $null, $Options, 1)) -eq 0
					} #End If Test-Path $SaveLocation
				} #End If !$Overwrite	
				
				$NewImage = new-object System.Drawing.Bitmap $NewWidth, $NewHeight

				$Graphics = [System.Drawing.Graphics]::FromImage($NewImage)
				$Graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
				$Graphics.DrawImage($OldImage, 0, 0, $NewWidth, $NewHeight) 

				$ImageFormat = $OldImage.RawFormat
				$OldImage.Dispose()
				if ($DPIWidth -and $DPIHeight) {
					$NewImage.SetResolution($DPIWidth, $DPIHeight)
				} #End If $DPIWidth -and $DPIHeight
				

				$NewImage.Save($SaveLocation, $ImageFormat)
				$NewImage.Dispose()
				Write-Verbose "Image '$ImageFile' was resize from $($OldWidth)x$($OldHeight) to $($NewWidth)x$($NewHeight) and save in '$SaveLocation'"
					
				If ($RemoveSource) {
					Remove-Item $Image -Force
					Write-Verbose "Image source '$ImageFile' was removed"
				} #End If $RemoveSource

			}
		}

	} #End Process
	
	End {}
}
function Connect-XExchange([string] $username, [string] $servername) {

	$session = Get-PSSession | Where-Object { $_.ConfigurationName -match "Microsoft.Exchange" } 
	if (!($session)) {
		if ($env:USERNAME -eq $username) {
			$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$servername.ad.mc.com/PowerShell/" -Authentication Kerberos
		}
		else {
			$cred = Get-Credential -Message "Exchange Admin Password" -UserName $username
			$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$servername.ad.mc.com/PowerShell/" -Authentication Kerberos -Credential $cred
		}
		Import-PSSession -Session $session -DisableNameChecking
	}
	$session | Format-Table -AutoSize
}

function New-SamAccount {
	param (
		$FirstNameChars, $LastName, $Number
	)
	$FirstNameChars = $FirstNameChars -replace " ", ""
	$FirstNameChars = $FirstNameChars -replace "-", ""
	$LastName = $LastName -replace " ", ""
	$LastName = $LastName -replace "-", ""

	$samaccount = $FirstNameChars + $LastName
	if ($samaccount.Length -gt $Number) {
		$samaccount = $samaccount.Substring(0, ($Number))
	}
	$samaccount = $samaccount.ToLower()
	$samaccount = $samaccount.Trim()
	$samaccount
}


# This function is used to upload to AD user account AND the user's mailbox.
# I uploaded to to both the AD user account and the mailbox because a higher quality photo could be stored in the mailbox and
# once you move to Exchange Online the photo has to be set on the mailbox. 
# We also had some custom made programs that used the photo that was stored in AD so we needed the photo in both locations.
function Set-XUserPhoto {
	param (
		[object]$photo
	)
	if (Test-Path $photo.fullname) {
		if ($photo.User -notlike $null) {
	
			if ($Photo.user.ThumbnailPhoto -notlike $null) {
				Write-Host "User already has photo....Removing existing photo for $($photo.User.DisplayName)"
				Remove-UserPhoto $Photo.User.SamaccountName -ClearMailboxPhotoRecord -Confirm:$false
			}
			
			Set-UserPhoto -Identity $photo.user.samaccountname -PictureData ([System.IO.File]::ReadAllBytes($photo.fullName)) -Preview -Confirm:$false 
			Set-UserPhoto $photo.user.samaccountname -Save -Confirm:$false
			if ($?) {
				Write-Host "Successful 'Set-UserPhoto' for: $($photo.user.DisplayName) from: $($photo.Folder).....with photo: $($photo.Name)"
				$photo.SuccessfulEXOUpload = $true
			}
			else {
				Write-Host "Failed to upload photo:$($photo.Name) for user $($photo.user.Samaccountname) from: $($photo.Folder)"	
			}

			$file = [System.IO.File]::ReadAllBytes($photo.SmallerPhoto.FullName)
			Set-aduser $photo.User.SamAccountName Replace @{thumbnailPhoto=$file} -Confirm:$false
			if ($?) {
				Write-Host "Successful 'Set-Aduser' for: $($photo.user.DisplayName) from Folder: $($photo.SmallerPhoto.Directory).....with photo: $($photo.SmallerPhoto.Name)"
				$photo.SuccessfulADUpload = $true
			}
			else {
				"Failed to upload photo to users' on prem AD User Account for user: $($photo.User.DisplayName) with photo: $($photo.SmallerPhoto.Name)"
			}
			
		}
	}
	else {
		Write-Host -ForegroundColor Yellow "Can't find $($photo.user.SamaccountName)'s photo at this location: $($photo.Fullname)"
	}
	$photo
}




    



$Email = Read-Host "What email address do you want to send alerts to?"

while (( $Email -notmatch '@.*\..*')) {
	$Email = Read-Host 'Im sorry I didnt understand....What email address do you want to send alerts to?'
}

#This is where you define the SMTP server info
$smtpServer = "SMTPserverName"
$message = New-Object System.Net.Mail.MailMessage
$message.From = "ADPhotoUploader@mrcy.com"
$message.To.Add("$($Email)")




function Get-PicturesInUpdate {
	param(
		$UPDATEFolders
	)
	[System.Collections.ArrayList]$PicturesInUpdate = @()
	foreach ($folder in $UPDATEFolders ) {

		$pics = Get-ChildItem "$($folder.fullname)" -File

		if ($pics.Count -ne 0) {
            
			if ($pics.count -gt 1) {
				[void]$PicturesInUpdate.AddRange($pics)
			}
			else {
				[void]$PicturesInUpdate.Add($pics)
			}
		}

	} 
	$PicturesInUpdate
}


#This section is where you define where your Sharepoint location is that holds the badge photos
$SharePointDrive = Get-PSDrive | Where-Object { $_.Root -match "Badge Photos" }

#This is where yo you create a PSDrive with the Sharepoint Location
if (!$SharePointDrive) {
	New-PSDrive -Name "Z" -Root "\\sp.SharePointSite.com@SSL\DavWWWRoot\sites\Badge Photos\" -PSProvider FileSystem -Scope Global
}



$Date = Get-Date -Format "MM.dd.yyyy_HH.mm"


$UPDATEFolders = [System.Collections.ArrayList]@()
 foreach ($folder in Get-ChildItem -Directory -Path Z:\) {
	if ($folder.Name -notmatch "ASSESSMENT RECORDS" -and $folder.Name -notmatch "DON'T USE") {
		foreach ($secondfolder in Get-ChildItem -Directory -Path "$($folder.FullName)") {
			if (!(Test-Path "$($folder.FullName)\UPDATE")) {
				Read-Host "$($folder.FullName)\UPDATE does not exist Fix the problem then press any key to continue."
				$F = Get-Item "$($folder.FullName)\UPDATE"
				[void]$UPDATEFolders.add($F)
			}
			if ($secondfolder -match 'UPDATE') {
				[void]$UPDATEFolders.Add($secondfolder)
			} 
		}
	}
}

if(!($UPDATEFolders)){
	Write-Host "Couldn't find UPDATE folders, something has gone wrong!" -ForegroundColor Red
	break
}



if (!(Test-Path $env:USERPROFILE\PICS)) {
	mkdir $env:USERPROFILE\PICS
	mkdir $env:USERPROFILE\PICS\Upload
	mkdir $env:USERPROFILE\PICS\Upload\Exceptions
	mkdir $env:USERPROFILE\PICS\Upload\SmallerPhotosTemp
	mkdir $env:USERPROFILE\PICS\Upload\Not_Found
	mkdir $env:USERPROFILE\PICS\Upload\Wrong_Format
	mkdir $env:USERPROFILE\PICS\REPORTS
}





$pics = Get-ChildItem "$env:USERPROFILE\PICS\" -Include ('*.jpg', '*.jpeg', '*.png', '*.jfif') -Exclude "$env:USERPROFILE\PICS\Upload\Exceptions" -Recurse

if ($pics.count -gt 0) {
	Write-Host -ForegroundColor Yellow "Deleting photos stored locally in the $env:USERPROFILE\PICS folder."
	foreach ($pic in $pics) {
		try {
			Remove-Item $pic.Fullname -Force
		}
		catch {

			if ($_.Exception.Message -match ".*The process cannot access the file.*because it is being used by another process.*") {
				stop-process -name explorer –force
				Start-Sleep -Seconds 10
				Remove-Item $pic.Fullname -Force
			}
		}
			
	}
}





[System.Collections.ArrayList]$PhotoObjects = @()
foreach ($pic in Get-PicturesInUpdate -UPDATEFolders $UPDATEFolders) {
	$folder = $pic.Directory.Parent.Name
	$file = Copy-Item $pic.FullName -Destination "$env:USERPROFILE\PICS\Upload" -PassThru
	Write-Host "Downloaded photo: $($file.Name) from: $folder"
	$image = [System.Drawing.Image]::FromFile($file.FullName)
	$photo = [PSCustomObject]@{
		Name             = $file.Name
		BaseName         = $file.BaseName
		Extension        = $file.Extension
		FullName         = $file.FullName
		OldName          = $file.Name
		OldFullName      = $pic.FullName
		Directory        = $file.Directory
		Folder           = $folder
		User             = $null
		UserFound        = $false
		FirstName        = $null
		LastName         = $null
		SuccessfulEXOUpload = $false
		SuccessfulADUpload = $false
		OldSize          = "$($image.Width)x$($image.Height)"
		NewSize          = $null
		SmallerPhoto     = $null
		
	}
	$image.Dispose()
	[void]$PhotoObjects.Add($photo)
}						 		
Write-Host "------------------------------------"




[System.Collections.ArrayList]$PhotosWithoutUsers = @()



#Get list of User Organizational Units
Write-Host "Retrieving list of all users ....." -ForegroundColor Yellow
try {
	$OUs = Get-ADOrganizationalUnit -SearchBase "OU=Domain Users,DC=ad,DC=mc,DC=com" -SearchScope 1 -Filter *
}
catch {
	if ($_.Exception.Message -match ".*Unable to contact the server. This may be because this server does not exist.*") {
		$OUs = Get-ADOrganizationalUnit -SearchBase "OU=Domain Users,DC=ad,DC=mc,DC=com" -SearchScope 1 -Filter * -Server hq-dc3
	}
}

$properties = @(
	"GivenName",
	"SurName",
	"DistinguishedName",
	"ThumbnailPhoto",
	"SamAccountName",
	"Description",
	"Enabled",
	"CanonicalName",
	"EmployeeID",
	"DisplayName",
	"msExchRecipientTypeDetails",
	"extensionattribute15"
)
#Get list of Users with an OrganizanalUnit that's name is only 3 characters long.
#The users that I was trying to target where in these OUs.
$userList = foreach ($OU in $OUs) {
	if ($OU.Name.Length -eq 3) {
		Get-ADUser -Filter * -SearchBase $OU.DistinguishedName -properties $properties | Select $properties
	}
}



Write-Host "------------------------------------"
Start-Sleep -Seconds 3

Write-Host "Looking for User accounts to go with photos...."
Write-Host ""

# Here is where I match AD user accounts with photos, whith first doing looking for an EmployeeID in the Name of the photo.
# Then matching that employeeID with an user in AD.
# If the photo did not have the employeeId in the name, then I would try to find a match based on the First and Last Name of the user.
# Then finally if the First and LastName search fails then I would try to find by SamAccountName.

for ($i = 0; $i -le ($PhotoObjects.count - 1)) {

	
	$photo = $PhotoObjects[$i]
	if ($photo.user -like $null) {
		$employee = $null
				#Employee Ids are 5 digits long
		[regex]$regex = "\d\d\d\d\d"
		$matches = $regex.Matches("$($photo.BaseName)")
		
		if ($matches.Success) {
			foreach ($user in $userList) {
				if ($user.EmployeeID -eq $matches.Value) {
					$employee = $user
					break
				}
			}	
		}	
		if ($employee) {
			$DN = $employee.DistinguishedName -split "="
			$OU = ($DN[2] -split ",")[0]
			Write-Host -ForegroundColor Green "User Account found by Employee ID"
			Write-Host -ForegroundColor Green "User Account: $($employee.DisplayName)" 
			Write-Host -ForegroundColor Green "Photo Name: $($photo.Name)"
			Write-Host -ForegroundColor Green "User OU: $OU"
			Write-Host -ForegroundColor Green "Photo Folder: $($photo.Folder)"
			Write-Host "-----------------------------------"	
			$photo.user = $employee
			$photo.UserFound = $true
			$i++
		}
		else {
			$Name = $photo.BaseName -replace "^.*_", ""
			$Name = $Name -replace "^\d*\W*", ""
			if ($Name -match ",") {
				$names = $Name -Split (",")
				$photo.LastName = ($Names[0]).Trim()
				$photo.FirstName = ($Names[1] -replace "\W*$").Trim()
				if($photo.FirstName -match "\s{1,}"){
					$photo.FirstName = (($photo.FirstName -split " ")[0]).Trim()
				}
			}
			else {
				$names = $Name -Split (" ")
				$photo.LastName = ($Names[1] -replace "\W*$").Trim()
				$photo.FirstName = ($Names[0]).Trim()
			}
			

			
		
			# if ($names.count -lt 2) {
			# 	$names = $Name -cSplit "(?=[A-Z])"
			# 	$names = $names.Where( { $_ -notlike $null })
			# }
					
			if ($names.count -ge 2) {
				# $names[1] = $names[1] -replace "[\._].*", ""
				# $photo.FirstName = $names[1].Trim() -replace "[\W*_*]", ""
				# $photo.LastName = $names[0].Trim() -replace "[\W*_*]", ""

						
				$FoundUsers = [System.Collections.ArrayList]@()
				foreach ($user in $userList) {
					if ($user.GivenName -imatch $photo.FirstName -and $user.Surname -imatch $photo.LastName) {
						[void]$FoundUsers.Add($user)
						break
					}
				}
				if ($foundusers.count -eq 1) {
					$DN = $foundusers[0].DistinguishedName -split "="
					$OU = ($DN[2] -split ",")[0]
					Write-Host -ForegroundColor Blue "User Account found by LastName, FirstName"
					Write-Host -ForegroundColor Blue "User Account: $($foundusers[0].DisplayName)" 
					Write-Host -ForegroundColor Blue "Photo Name: $($photo.Name)"
					Write-Host -ForegroundColor Blue "User OU: $OU"
					Write-Host -ForegroundColor Blue "Photo Folder: $($photo.Folder)"
					Write-Host "-----------------------------------"			
					$photo.user = $foundusers[0]
					$photo.UserFound = $true
					$i++
				}			
				else {
					
					$PossibleSamAccount = New-SamAccount -FirstNameChars $photo.FirstName[0] -LastName $photo.LastName -Number 8
					$foundusers = [System.Collections.ArrayList]@()
					foreach ($user in $userList) {
						if ($user.SamAccountName -like $PossibleSamAccount) {
							[void]$foundusers.add($user)
							break
						}
					}
					if($foundusers.count -ne 1){
						$PossibleSamAccount = New-SamAccount -FirstNameChars $photo.FirstName[0..1] -LastName $photo.LastName -Number 8
						$foundusers = [System.Collections.ArrayList]@()
						foreach ($user in $userList) {
							if ($user.SamAccountName -like $PossibleSamAccount) {
								[void]$foundusers.add($user)
								break
							}
					}

					}
					#If couldn't find by EmployeeId or First and last name then will try to find by SamAccountName.
					if ($foundusers.count -eq 1) {
						$DN = $foundusers[0].DistinguishedName -split "="
						$OU = ($DN[2] -split ",")[0]
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "User found by SamAccountName"
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "User SamAccountName: $PossibleSamAccount"
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "User Account: $($foundusers[0].DisplayName)"
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "Photo Name: $($photo.Name)"
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "User OU: $OU"
						Write-Host -ForegroundColor Red -BackgroundColor Yellow "Photo Folder: $($photo.Folder)"
						Write-Host "-----------------------------------"
						
						$message.Subject = "Uploader Script Needs Confirmation"
						$message.IsBodyHTML = $true
						$message.Body = "<b>AD Photo Uploader Script Needs Confirmation</b>"
						$SMTP = New-Object System.Net.Mail.SmtpClient($smtpServer)
						$smtp.Send($message)

						
						[console]::beep(1000, 700); [console]::beep(1000, 700);
						$UserInput = $null
						$UserInput = Read-Host 'Does the SamAccount work for this photo? Y\N'
						Write-Host "-----------------------------------"

						while (( $UserInput -notmatch 'Y') -and ($UserInput -notmatch 'N')) {
							$UserInput = Read-Host 'Im sorry I didnt understand....Does the SamAccount work for this photo? Y\N'
						}
						if ($UserInput -imatch 'Y') {
							$photo.user = $foundusers[0]
							$photo.UserFound = $true
							$i++
						}
						else {
							$photo.FullName = (Move-Item "$env:USERPROFILE\PICS\Upload\$($Photo.Name)" -Destination $env:USERPROFILE\PICS\Upload\Not_Found -PassThru -Force).FullName
							Write-Host "-----------------------------------"
							Write-Host -ForegroundColor Red "No Users Found for photo: $($Photo.Name)  from: $($photo.Folder) with Name: $($photo.firstName) $($photo.LastName)"
							Write-Host "-----------------------------------"
							[void]$PhotosWithoutUsers.add($photo)
							[void]$PhotoObjects.Remove($photo)	
						}			
						
					}
					else {
						$photo.FullName = (Move-Item "$env:USERPROFILE\PICS\Upload\$($Photo.Name)" -Destination $env:USERPROFILE\PICS\Upload\Not_Found -PassThru -Force).FullName
						Write-Host "-----------------------------------"
						Write-Host -ForegroundColor Red "No Users Found for photo: $($Photo.Name)  from: $($photo.Folder)"
						Write-Host "-----------------------------------"
						[void]$PhotosWithoutUsers.add($photo)
						[void]$PhotoObjects.Remove($photo)	
					}	
				}
			}
			else {
				Move-Item "$env:USERPROFILE\PICS\Upload\$($Photo.Name)" -Destination $env:USERPROFILE\PICS\Upload\Wrong_Format -Force
				$photo.FullName = "$env:USERPROFILE\PICS\Upload\Wrong_Format\$($Photo.Name)"
				Write-Host "$($Photo.Name) is in the Wrong Format."
				[void]$PhotosWithoutUsers.add($photo)
				[void]$PhotoObjects.Remove($photo)
			}
					

		}
	}		
	else {
		$i++
	}
		
}

Write-Host "------------------------"
	Write-Warning "Be sure the user accounts found match the photos above."
	

	$UserInput = Read-Host "Do any of user accounts found NOT match the photos above Y/N?"
while (( $UserInput -inotmatch 'Y|N')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Do any of user accounts found NOT match the photos above Y/N?'
}

	if ($UserInput -imatch 'Y') {
		do {
			$XInput = Read-Host "Which photo(s) is matched with the wrong account?"

		for ($i = 0; $i -le ($PhotoObjects.count - 1)) {
			$PhotoObject = $PhotoObjects[$i]
			if("$Xinput" -eq $photoObject.Name){
				$PhotoObject.User = $null
				Write-Host "Removing $($user.DisplayName) from $($PhotoObject.Name)" -ForegroundColor Yellow
				[void]$PhotosWithoutUsers.add($PhotoObject)
				[void]$PhotoObjects.Remove($PhotoObject)	
				break
			}
			else {
				$i++
			}
		}
		$RunAgain = Read-Host "Is there another user account matched to a wrong photo? Y/N"
		} until (
			$RunAgain -imatch "N"
		)
		
	}

	

	

if ($PhotosWithoutUsers.Count -gt 0) {
	Write-Host "------------------------"
	Write-Warning "Be sure the user accounts found match the photos above."
	$UserInput = Read-Host 'Ready to find user accounts for Photos by SamAccountName? Y\N'

	
	for ($i = 0; $i -le ($PhotosWithoutUsers.count - 1)) {
		$photo = $PhotosWithoutUsers[$i]


		$SamAccountName = Read-Host "What is the SamAccountName for $($Photo.Name) of $($photo.Folder) ?"
		if ($SamAccountName -notlike $null) {
			$user = $null
			foreach ($member in $UserList) {
				if ($member.SamAccountName -like $SamAccountName) {
					$user = $member
					break
				}
			}
			while (!($user)) {
				$SamAccountName = Read-Host "We couldn't find a SamAccountName for $($Photo.Name).....Try again?"
				if ($SamAccountName -like $null -or $SamAccountName -eq "N") {
					break
				}
				$user = $null
				foreach ($member in $UserList) {
					if ($member.SamAccountName -like $SamAccountName) {
						$user = $member
						break
					}
				}	
			}
			if ($user) {
				$photo.user = $user
				$photo.UserFound = $true
				$photo.FullName = (Move-Item $photo.FullName -Destination $env:USERPROFILE\PICS\Upload\ -PassThru).FullName
				[void]$PhotoObjects.add($photo)
				[void]$PhotosWithoutUsers.RemoveAt($i)
			}			
		}
		else {
			$i++
		}


	}
}

# This section pauses the script while you check the orienation of each photo through a program called GIMP.
# The photos just have to be in SQUARE orientation (Height and Width the same).
Write-Host ""
$UserInput = Read-Host "Make sure Photos are not SIDEWAYS using the program called GIMP.......Ready to Continue Y/N?"
while (( $UserInput -notmatch 'Y')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Ready to Continue Y/N?'
}
Write-Host ""
$UserInput = Read-Host 'Ready to resize photos? Y/N'

while (( $UserInput -notmatch 'Y')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Ready to resize photos? Y/N'
}

#Now this section will actually resize the photos.
foreach ($photo in $PhotoObjects) {
	if (Test-Path $photo.fullname) {
		try {
			Set-ImageSize -Image $photo.FullName -WidthPx 640 -HeightPx 640 -Verbose -MakePicturesSmaller $true -Destination $photo.Directory -Overwrite $true
		}
		catch {
			stop-process -name explorer –force
			Start-Sleep -Seconds 10
			Set-ImageSize -Image $photo.FullName -WidthPx 640 -HeightPx 640 -Verbose -MakePicturesSmaller $true -Destination $photo.Directory -Overwrite $true
		}
		$image = [System.Drawing.Image]::FromFile($photo.FullName)

		if ($photo.OldSize -ne "$($image.Width)x$($image.Height)") {
			$photo.NewSize = "$($image.Width)x$($image.Height)"
		}
		else {
			$photo.NewSize = $photo.OldSize
		}
		$image.Dispose()
		$photo.SmallerPhoto = Copy-item $photo.fullName -Destination $env:USERPROFILE\PICS\Upload\SmallerPhotosTemp\ -Force -PassThru
		if($photo.SmallerPhoto){
			try {
				Set-ImageSize -Image $photo.SmallerPhoto.FullName -WidthPx 96 -HeightPx 96 -Verbose -MakePicturesSmaller $true -Destination $photo.SmallerPhoto.Directory -Overwrite $true
			}
			catch {
				stop-process -name explorer –force
				Start-Sleep -Seconds 10
				Set-ImageSize -Image $photo.SmallerPhoto.FullName -WidthPx 96 -HeightPx 96 -Verbose -MakePicturesSmaller $true -Destination $photo.SmallerPhoto.Directory -Overwrite $true
			}
			
		}
		
	}
	else {
		Write-Host -ForegroundColor Yellow "Can't find $($photo.user.SamaccountName)'s photo at this location: $($photo.Fullname)"
	}

}


$UserInput = Read-Host 'Ready to Upload photos? Y/N'

while (( $UserInput -notmatch 'Y')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Ready to Upload photos? Y/N'
}

#This section would actually would check attribute15 to see which tenant the user was a member of

$GCCHMailboxes = [System.Collections.ArrayList]@()
$CommercialMailboxes = [System.Collections.ArrayList]@()
$OnPremMailboxes = [System.Collections.ArrayList]@()
foreach ($photo in $photoObjects) {
	if($photo.User.msExchRecipientTypeDetails -ne 1)
	{
		if ($photo.User.ExtensionAttribute15 -match "GCCHigh") {
			[void]$GCCHMailboxes.Add($photo)
		}
		if ($photo.User.ExtensionAttribute15 -match "Commercial") {
			[void]$CommercialMailboxes.Add($photo)
		}
	}
	else {
		[void]$OnPremMailboxes.Add($photo)
	}
}

#If there any On prem mailboxes then this is where I connect to on prem exchange and upload photos to those users.
if($OnPremMailboxes.count -gt 0) {
	if (!($ExchSesssion)) {
		$username = Read-Host "Admin username for On-Prem Exchange?"
		$servername = Read-Host "On-Prem Exchange Server Name?"
		Connect-XExchange -username $username -servername $servername 
	}

	foreach ($photo in $OnPremMailboxes) {

		$photo = Set-XUserPhoto -photo $photo
	}
	Get-PSSession | Where-Object {$_.ConfigurationName -match "Microsoft.Exchange" -and $_.ComputerName -match "ad\.mc\.com"} | Remove-PSSession
}


if($GCCHMailboxes.count -gt 0 -or $CommercialMailboxes.count -gt 0){
	if($GCCHMailboxes.count -gt 0)
	{
		Write-Host "Need to Connect to GCCHigh Exchange Online to upload photos for those users" -ForegroundColor Yellow
		Write-Host ""
		Write-Host "There will be another window that comes up requesting GCCHigh Exchange Online Admin credentials." -ForegroundColor Yellow

		Connect-ExchangeOnline -ExchangeEnvironmentName "O365USGovGCCHigh" -ShowBanner:$false

		foreach ($photo in $GCCHMailboxes) {

			$photo = Set-XUserPhoto -photo $photo
		}
		Disconnect-ExchangeOnline -Confirm:$false
	}
	if($CommercialMailboxes.Count -gt 0)
	{
		Write-Host "Need to Connect to Commercial Exchange Online to upload photos for those users" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "There will be another window that comes up requesting Commercial Exchange Online Admin credentials." -ForegroundColor Yellow

		Connect-ExchangeOnline -ShowBanner:$false

		foreach ($photo in $CommercialMailboxes) {

			$photo = Set-XUserPhoto -photo $photo
		}
		Disconnect-ExchangeOnline -Confirm:$false
	}
}




Write-Host "Creating Report of users who still do not have photos...." -ForegroundColor Yellow
Get-ADUser -searchBase "OU=Domain Users,DC=ad,DC=mc,DC=com"  -filter { ThumbnailPhoto -notlike "*" -AND EmployeeID -like "*" -AND Enabled -eq $true } -properties ThumbnailPhoto, SamAccountName, Description, Enabled, CanonicalName, EmployeeID `
| Select-Object UserPrincipalName, SamAccountName, Description, Enabled, CanonicalName | Export-Csv "$env:USERPROFILE\PICS\REPORTS\AfterUploadReport$date.csv"


$message.Subject = "Uploader Script Is done uploading photos"
$message.IsBodyHTML = $true
$message.Body = "<b>AD Photo Uploader Script is done uploading photos</b>"
$SMTP = New-Object System.Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
[console]::beep(1000, 700); [console]::beep(1000, 700);
Write-Host ""

Write-Host ""
$UserInput = Read-Host "Do you want to rename photos in the PICS folder? Y/N"   

while (( $UserInput -notmatch 'Y') -and ($UserInput -notmatch 'N')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Do you want to rename photos in the PICS folder? Y/N'
}
Write-Host ""
Write-Host "-------------------------------------------"


if ($UserInput -match "Y") {
	foreach ($photo in $PhotoObjects) {

		if ($photo.SuccessfulEXOUpload -eq $true -and $photo.SuccessfulADUpload) {
			$newname = $photo.user.Surname + ", " + $photo.user.GivenName + "$($photo.Extension)"


			

			$newphoto = Rename-Item -NewName "$newname" -Path $photo.FullName -PassThru -Force

			if (!($?)) {
				if ($error[0] -match ".*The process cannot access the file.*because it is being used by another process.*") {
					Write-Host "Restarting Windows Explorer... and then waiting some time" -ForegroundColor Yellow 
					stop-process -name explorer –force
					Start-Sleep -Seconds 13
					$newphoto = Rename-Item -NewName "$newname" -Path $photo.FullName -PassThru -Force	
				}
			}
			
			if ($newphoto) {
				$photo.Name = $newphoto.name
				$photo.FullName = $newphoto.fullname	
				Write-Host "Renamed $($photo.oldName) to....$($photo.name)"

			}
			else {
				Write-Host "Failed to Rename $($photo.Name)"
			}
			
	
			
		}
	}
}





$UserInput = Read-Host "Are you ready to upload these photos to Sharepoint? Y/N"   

while (( $UserInput -notmatch 'Y') -and ($UserInput -notmatch 'N')) {
	$UserInput = Read-Host 'Im sorry I didnt understand....Are you ready to upload these photos to Sharepoint? Y/N'
}
Write-Host ""
Write-Host "-------------------------------------------"
if ($UserInput -match "Y") {
	foreach ($photo in $PhotoObjects) {
		if ($photo.SuccessfulEXOUpload -and $photo.SuccessfulADUpload) {
			$movedfile = $null
			try {
				$movedfile = Move-Item -Path $photo.FullName -Destination "Z:\$($photo.Folder)\" -Force -PassThru
			}
			catch {
				if ($_.Exception.Message -match ".*The process cannot access the file.*because it is being used by another process.*") {
					stop-process -name explorer –force
					Start-Sleep -Seconds 10
					$movedfile = Move-Item -Path $photo.FullName -Destination "Z:\$($photo.Folder)\" -Force -PassThru
				}
			}		
			
			if ($movedfile) {					
				Write-Host "Moved $env:USERPROFILE\PICS\$($photo.name) to.....Z:\$($photo.Folder)\"
				Remove-Item "Z:\$($photo.Folder)\UPDATE\$($photo.OldName)" -Force
				Write-Host "Removed from Update Folder Z:\$($photo.Folder)\UPDATE\$($photo.OldName)"

				try {
					Remove-Item "$($photo.SmallerPhoto.FullName)" -Force
				}
				catch {
					if ($_.Exception.Message -match ".*The process cannot access the file.*because it is being used by another process.*") {
						stop-process -name explorer –force
						Start-Sleep -Seconds 10
						Remove-Item "$($photo.SmallerPhoto.FullName)" -Force
					}
				}
				
			}
			else {
				Write-Host "-------------------------------------------"
				Write-Host "FAILED to Move $env:USERPROFILE\PICS\$($photo.name)"
				Write-Host "-------------------------------------------"
			}
		}
	}
}

if ($PhotosWithoutUsers.count -gt 0) {
	Write-Host -ForegroundColor Yellow "The Folowing Photos were not uploaded because there was no user account found:"
	Write-Host "------------------------------------------"
	
	foreach ($photo in $PhotosWithoutUsers) {
		Write-Host "$($photo.Name) from: $($photo.Folder)" 
	}
	Write-Host "------------------------------------------"
	$UserInput = Read-Host "Do you want to upload the photos that dont have user accounts to Sharepoint? Y/N"   
	
	while (( $UserInput -notmatch 'Y') -and ($UserInput -notmatch 'N')) {
		$UserInput = Read-Host 'Im sorry I didnt understand....Do you want to upload the photos that dont have user accounts to Sharepoint? Y/N'
	}


	
	if ($UserInput -match 'Y') {
		foreach ($photo in $PhotosWithoutUsers) {
			if (!($photo.SuccessfulEXOUpload) -and !($photo.SuccessfulADUpload)) {
				$movedfile = $null
				try {
					$movedfile = Move-Item -Path $photo.FullName -Destination "Z:\$($photo.Folder)\" -Force -PassThru
				}
				catch {
					if ($_.Exception.Message -match ".*The process cannot access the file.*because it is being used by another process.*") {
						stop-process -name explorer –force
						Start-Sleep -Seconds 10
						$movedfile = Move-Item -Path $photo.FullName -Destination "Z:\$($photo.Folder)\" -Force -PassThru
					}
				}
				if ($movedfile) {	
					Write-Host "User NOT FOUND for....$($photo.Name)"				
					Write-Host "Moved $($photo.Name) to.....Z:\$($photo.Folder)\"
					Remove-Item "Z:\$($photo.Folder)\UPDATE\$($photo.OldName)" -Force
					Write-Host "Removed from Update Folder Z:\$($photo.Folder)\UPDATE\$($photo.OldName)"
					
				}
				elseif (!($movedfile)) {
					Write-Host "-------------------------------------------"
					Write-Host "FAILED to Move $env:USERPROFILE\PICS\$($photo.name)"
					Write-Host "-------------------------------------------"
				}
			}
		}
			
	}
	
}
Write-Host "-------------------------------------------"
$PhotoObjects | Select OldName, Folder, SuccessfulEXOUpload,SuccessfulADUpload, OldSize, NewSize, UserFound | FT
$Report = "$env:USERPROFILE\PICS\REPORTS\PhotoUploadReport_$Date.csv"

if (Test-Path $Report) {
	Remove-Item $Report -Force
}
$PhotoObjects | Select OldName, Folder, SuccessfulEXOUpload, SuccessfulADUpload, OldSize, NewSize, UserFound | Export-csv $Report -NoTypeInformation

Remove-PSDrive -Name Z