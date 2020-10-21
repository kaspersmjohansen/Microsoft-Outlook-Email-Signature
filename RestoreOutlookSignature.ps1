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

$PrimSigPath = "$BCKFolder\Outlook\PrimSig"
$ReplySigPath = "$BCKFolder\Outlook\ReplySig"
$OtherSigPath = "$BCKFolder\Outlook\Other"

$PrimSig = Get-ChildItem "$PrimSigPath\*.*" | Foreach-Object {$_.BaseName} | Select-Object -Last 1
$ReplySig = Get-ChildItem "$ReplySigPath\*.*" | Foreach-Object {$_.BaseName} | Select-Object -Last 1

Copy-Item "$PrimSigPath\*" -Recurse "$SigSource"
Copy-Item "$ReplySigPath\*" -Recurse "$SigSource"

#Set Primary New messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.NewMessageSignature="$PrimSig"
$MSWord.Quit() 
 
#Set Reply signature as default for Reply/Forward messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.ReplyMessageSignature="$ReplySig" 
$MSWord.Quit() 

# Start-Sleep 10
# Start-Process "C:\Program Files\Microsoft Office\Office$OfficeNumVer\Outlook.exe"