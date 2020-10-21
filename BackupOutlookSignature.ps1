# ******************************************************************************************
# This script is meant to backup users New Mail and Reply Mail signatures to a network share
# The script can be implemented as either a logon script or a logoff script.
# Make sure that the user has write and modify permissions to the network share
# 
# Author: Kasper Johansen, http://virtualwarlock.net
# 
# ******************************************************************************************

$BCKFolder = "C:\Working"
$LogDir = $BCKFolder
$SigSource = "$env:AppData\Microsoft\Signatures"

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
If (!(Test-Path -Path "$BCKFolder\Outlook")){
New-Item -ItemType Directory -Force "$BCKFolder\Outlook"
New-Item -ItemType Directory -Force "$BCKFolder\Outlook\PrimSig"
New-Item -ItemType Directory -Force "$BCKFolder\Outlook\ReplySig"
New-Item -ItemType Directory -Force "$BCKFolder\Outlook\Other"
}

$PrimSigPath = "$BCKFolder\Outlook\PrimSig"
$ReplySigPath = "$BCKFolder\Outlook\ReplySig"
$OtherSigPath = "$BCKFolder\Outlook\Other"

if ($NewMsgSig -eq "(none)") {
Write-Host "No new mail signature set"
}
elseif ($NewMsgSig){
Copy-Item -Path "$SigSource\$NewMsgSig-files" -Destination "$PrimSigPath" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -Path "$SigSource\$NewMsgSig-filer" -Destination "$PrimSigPath" -Recurse -Force
Copy-Item -Path "$SigSource\$NewMsgSig.*" -Destination "$PrimSigPath" -Force
}


if ($ReplyMsgSig -eq "(none)") {
Write-Host "No reply signature set"
}
elseif ($ReplyMsgSig){
Copy-Item -Path "$SigSource\$ReplyMsgSig-files" -Destination "$ReplySigPath" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item -Path "$SigSource\$NewMsgSig-filer" -Destination "$PrimSigPath" -Recurse -Force
Copy-Item -Path "$SigSource\$ReplyMsgSig.*" -Destination "$ReplySigPath" -Force 
}

$Exclude=@("$NewMsgSig.*","$NewMsgSig-Filer","$NewMsgSig-Files","$ReplyMsgSig.*","$ReplyMsgSig-filer","$ReplyMsgSig-files")
Copy-Item -Path "$SigSource\*" -Destination $OtherSigPath -Recurse -Force -Exclude $Exclude