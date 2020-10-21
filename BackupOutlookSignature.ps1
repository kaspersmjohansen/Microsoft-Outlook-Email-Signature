# ******************************************************************************************
# This script is meant to backup users New Mail and Reply Mail signatures to a network share
# The script can be implemented as either a logon script or a logoff script.
# Make sure that the user has write and modify permissions to the network share
# 
# Author: Kasper Johansen, http://virtualwarlock.net
# 
# ******************************************************************************************

$AppData = $env:AppData
$Homedrive = $env:homedrive
$Homepath = $env:homepath
$Homedir = "$Homedrive$Homepath"
$BCKFolder = ""
$LogDir = ""
$OfficeVersion =

if ($OfficeVersion -eq 2013) {
	$OfficeNumVer = "15.0"
}
elseif ($OfficeVersion -eq 2010) {
	$OfficeNumVer = "14.0"
}
elseif ($OfficeVersion -eq 2007) {
	$OfficeNumVer = "12.0"
}

$OfficeLang = (get-itemproperty -Path "HKCU:\Software\Microsoft\Office\$OfficeNumVer\Common\LanguageResources" -Name UILanguage).UILanguage
 
if ($OfficeLang -eq 1033) {
	$SigSource = "$AppData\Microsoft\Signatures"
}
elseif ($OfficeLang -eq 1030) {
	$SigSource = "$AppData\Microsoft\Signaturer"
}

if (Get-Process outlook -ErrorAction silentlycontinue) {
	Stop-Process -Name "Outlook" -ErrorAction silentlycontinue
}
else {
	Write-Host "Outlook is not running"
}

Start-Sleep -s 5

#Extract current signature settings in new messages and reply messages
$MSWord = New-Object -com word.application
$EmailOptions = $MSWord.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$NewMsgSig = $EmailSignature.NewMessageSignature
$ReplyMsgSig = $EmailSignature.ReplyMessageSignature
Write-Host "NewMsgSignature name is $NewMsgSig"
Write-Host "ReplyMsgSignature name is $ReplyMsgSig"
$MSWord.Quit()

#Create folder structure in backup location. The PrimSig folder contains the New Mail signature and the ReplySig folder contains the Reply Mail signature. 
#Other signatures are copied to the Signatures folder in the backup location
#A log file is also created
If (!(Test-Path -Path "$BCKFolder\Outlook\Signatures")){
New-Item -ItemType Directory -Force "$BCKFolder"
New-Item -ItemType Directory -Force "$BCKFolder\PrimSig"
New-Item -ItemType Directory -Force "$BCKFolder\ReplySig"
New-Item -ItemType Directory -Force "$BCKFolder\Other"
}

if (($NewMsgSig -eq "(ingen)") -or ($NewMsgSig -eq "(none)")) {
Write-Host "No new mail signature set"
}
elseif (($NewMsgSig) -and ($OfficeUI -eq 1033)){
Copy-Item "$SigSource\$NewMsgSig-files" "$BCKFolder\PrimSig\" -Recurse -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
Copy-Item "$SigSource\$NewMsgSig.*" "$BCKFolder\PrimSig\" -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
}
elseif (($NewMsgSig) -and ($OfficeUI -eq 1030)){
Copy-Item "$SigSource\$NewMsgSig-filer" "$BCKFolder\PrimSig\" -Recurse -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
Copy-Item "$SigSource\$NewMsgSig.*" "$BCKFolder\PrimSig\" -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
}

if (($ReplyMsgSig -eq "(ingen)") -or ($ReplyMsgSig -eq "(none)")) {
Write-Host "No reply signature set"
}
elseif (($ReplyMsgSig) -and ($OfficeUI -eq 1033)){
Copy-Item "$SigSource\$ReplyMsgSig-files" "$BCKFolder\ReplySig\" -Recurse -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
Copy-Item "$SigSource\$ReplyMsgSig.*" "$BCKFolder\ReplySig\" -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
}
elseif (($ReplyMsgSig) -and ($OfficeUI -eq 1030)){
Copy-Item "$SigSource\$ReplyMsgSig-filer" "$BCKFolder\ReplySig\" -Recurse -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
Copy-Item "$SigSource\$ReplyMsgSig.*" "$BCKFolder\ReplySig\" -Force -PassThru | Out-File -Append "$LogDir\SigLog.txt"
}

$Exclude=@("$NewMsgSig.*","$NewMsgSig-Filer","$NewMsgSig-Files","$ReplyMsgSig.*","$ReplyMsgSig-filer","$ReplyMsgSig-files")
Copy-Item "$SigSource\*" "$BCKFolder\Other\" -Recurse -Force -Exclude $Exclude -PassThru | Out-File -Append "$LogDir\SigLog.txt"