# ******************************************************************************************
# This script is meant to backup users New Mail and Reply Mail signatures to a network share
# The script can be implemented as either a logon script or a logoff script.
# Make sure that the user has write and modify permissions to the network share
# 
# Author: Kasper Johansen, http://virtualwarlock.net
# 
# ******************************************************************************************

$Homedrive = $env:homedrive
$Homepath = $env:homepath
$Homedir = "$Homedrive$Homepath"
$PrimSigPath = ""
$ReplySigPath = ""
$PrimSig = Get-ChildItem "$PrimSigPath\*.*" | Foreach-Object {$_.BaseName} | Select-Object -Last 1
$ReplySig = Get-ChildItem "$ReplySigPath\*.*" | Foreach-Object {$_.BaseName} | Select-Object -Last 1
$AppData = $env:AppData
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
	$SigFolder = "$AppData\Microsoft\Signatures"
}
elseif ($OfficeLang -eq 1030) {
	$SigFolder = "$AppData\Microsoft\Signaturer"
}

Copy-Item "$PrimSigPath\*" -Recurse "$SigFolder"
Copy-Item "$ReplySigPath\*" -Recurse "$SigFolder"

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
Start-Sleep 10
#Start-Process "C:\Program Files\Microsoft Office\Office$OfficeNumVer\Outlook.exe"