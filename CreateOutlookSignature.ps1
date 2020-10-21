# Use https://www.base64-image.de to create the Logo Base64 code.

Clear-Host
Write-Verbose "Setting Arguments" -Verbose
$StartDTM = (Get-Date)

$Path = "$env:APPDATA\Microsoft\Signatures"
##$User = Get-ADUser -Identity $env:USERNAME -Properties *
$CurrentUser = New-Object DirectoryServices.DirectorySearcher "samaccountname=$env:USERNAME"
$UserADProperties = $CurrentUser.FindOne() | select -ExpandProperty Properties
$SignatureName = "Signature-$env:username"
$Out = "$Path\$SignatureName" + ".htm"

If (!(Test-Path -Path $Path))
{
    New-Item -ItemType directory -Path $Path | Out-Null
}

Push-Location $Path

$full_name = “$($UserADProperties.givenname) $($UserADProperties.sn)”
$account_name = "$($UserADProperties.samaccountname)"
$job_title = "$($UserADProperties.title)"
$location = "$($UserADProperties.office)"
$address = "$($UserADProperties.streetaddress)"
$city = "$($UserADProperties.l)"
#$state = "$($UserADProperties.state)"
#$country = "$($UserADProperties.co)"
$zip = "$($UserADProperties.postalcode)"
$dept = "$($UserADProperties.department)"
$comp = "$($UserADProperties.company)"
$email = "$($UserADProperties.mail)"
$phone =  "$($UserADProperties.mobile)"
#$fax =  "$($UserADProperties.fax)"
$upn = "$($UserADProperties.userprincipalname)"
$www =  "www.virtualwarlock.net"
$url = "https://" + "$www"
$regards = "Med venlig hilsen / Best regards"

$logo = "data:image/png;base64,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"

"<span style=`"font-family: calibri;`">" 
"<font color=black>" + $regards + "<br />",
"<strong>" + $full_name + "</strong><br />", $job_title + "<br><br />", 
"<font color=black>" + "M " + "</font color=black>" + $phone + "<br />", 
"<font color=black>" + "<a href='$email'>$email</a>" + "</font color=black>" + "<br />", 
"</span><br />",
"<img alt=`"corporate logo`" border=`"0`" src=`"$logo`"`" />" | Out-File $Out

$Profiles = (Get-ChildItem HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles).PSChildName
$OutLookProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\" + $Profiles.Trim() + "\9375CFF0413111d3B88A00104B2A6676\00000002"

Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "New Signature" -value $SignatureName -Propertytype string -Force 
Get-Item -Path $OutlookProfilePath | New-Itemproperty -Name "Reply-Forward Signature" -value $SignatureName -Propertytype string -Force

<#
Write-Verbose "Import Microsoft Online PowerShell Module"
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
if (!(Test-Path -Path "C:\Program Files\PackageManagement\ProviderAssemblies\nuget")) {Find-PackageProvider -Name 'Nuget' -ForceBootstrap -IncludeDependencies}
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser | Import-Module -Name ExchangeOnlineManagement | Out-Null
Connect-ExchangeOnline -UserPrincipalName $upn
$Mailbox = Get-EXOMailbox


$HeadingLine = "<HTML><HEAD><TITLE>Signature</TITLE><BODY><BR><table style=`"FONT-SIZE: 10pt; COLOR: black; FONT-FAMILY: `'Arial`' `"> <tr>"
$PersonLine = "<td padding='0'><B>" + $M.DisplayName + " </B>" + "<BR>" + $Job_Title + "<BR><BR>"
$PhoneLine = "<strong>" + "M| " + "</strong>" + $phone + "<BR>"
$EmailLink = "<strong>" + "E| " + "</strong>" + '<a href=mailto:"' + $($M.PrimarySmtpAddress) + '">' + $($M.PrimarySmtpAddress) + "<BR>"
$EndLine = "</td></tr></table><BR><BR></BODY></HTML>"
$SignatureHTML = $HeadingLine + $PersonLine + $PhoneLine + $EmailLink + $EndLine


<#
Write-Verbose "Setting Signature for OWA" -Verbose
Set-MailboxMessageConfiguration -Identity $Mailbox.Identity -SignatureHTML $SignatureHTML -AutoAddSignature $True -AutoAddSignatureOnReply $False
#>

Pop-Location