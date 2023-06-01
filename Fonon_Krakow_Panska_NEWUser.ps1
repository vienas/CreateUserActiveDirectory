Clear-Host
Write-Host -ForegroundColor DarkGray  "FONON KRAKOW PANSKA"

$Credentials = Get-StoredCredential -Target wasko.pl
$Username = $Credentials.UserName
$Password = $Credentials.Password
$Server = "wasko.pl"
$Credentials = New-Object System.Management.Automation.PSCredential $Username,$Password

$pathOU = "OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
(Get-ADOrganizationalUnit -Filter * -Credential $Credentials -Server $Server -SearchBase $pathOU | select DistinguishedName).DistinguishedName


$i = 0
do {$i++; $star="*"
        $section = $(Write-Host "ORGANIZATION UNIT: " -ForegroundColor Cyan -NoNewLine; Read-Host);
                $o=$section+$star;
                $path = @(Get-ADOrganizationalUnit -Filter {Name -like $o} -Credential $Credentials -Server $Server -SearchBase $pathOU | Select-Object distinguishedname)
        if ($o.Contains([string]$path.distinguishedname) = $false) { Write-Warning "Błędne dane wejściowe"}
            else {break}
}while ($i -le 2 )

if ($o.Contains([string]$path.distinguishedname) = $false) { Write-Host -ForegroundColor DarkRed "Kończenie pracy.."; [console]::beep(700,1000) ; Start-Sleep -Seconds 2; break}

[int]$numer = 0
    $baza =@{}
foreach ($_ in $path ) {
        $numer++
        Write-Host -ForegroundColor DarkCyan $numer. $_.distinguishedname
        $baza[$numer]= $_.distinguishedname
}


############ choise OU ##############L

$i=0
do { $i++
    
     $it = ($(Write-Host "WYBIERZ OU: " -ForegroundColor Cyan -NoNewLine; Read-Host))
     if ( $baza.Keys -eq $it) { break}
         elseif ( $baza.Keys -ne $it) {Write-Warning "Błędne dane wejściowe"}
}
while (($i -le 2) )
if ($it -notin $baza.Keys) { Write-Host -ForegroundColor DarkRed "Kończenie pracy.."; [console]::beep(700,1000) ; Start-Sleep -Seconds 2; break}

$mainpath = $baza.Item([int]$it)


############### declaring variables and input data from console

cls
write-host -ForegroundColor DarkGray "Organization Unit: $mainpath"

$Date = Get-Date
$Firstname = Read-Host "Imię użytkownika"
$Lastname = Read-Host "Nazwisko użytkownika"
$Title = Read-Host "Stanowisko"
$Department=Read-Host "Dział"
$AccountExpirationDate = Read-Host "Data wygaśnięcia konta: N [Never] lub [dd-mm-rr]"
$Add0=Read-Host "Dodać do grupy - HCPAW-RegularUsers ? T [Tak]"
$Add1=Read-Host "Dodać do grupy - O365SyncUser ? T [Tak]"
$Add2=Read-Host "Dodać do grupy - ALFA_ProjektyFONON_100016_NOKIA-RW ? T [Tak]"
$Add3=Read-Host "Dodać do grupy - ALFA_ProjektyFONON_100026_GSMR-RW ? [Tak]"
$Add4=Read-Host "Dodać do grupy - ALFA_ProjektyFONON_Pomoce_naukowe-RW ? T [Tak]"
$Add5=Read-Host "Dodać do grupy - ALFA_ProjektyFONON_Niemcy_projektowanie-RW ? T [Tak]"
$Firstname_user = $Firstname[0]
$Lastname_user = $Lastname[0..6] -join ''
$SamAccountUser = ($Firstname_user + $Lastname_user).ToLower()
$SamAccountUser_email = ($Firstname_user + "." + $Lastname).ToLower()
$hash = @{'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ó'='o'; 'ś'='s'; 'ń'='n'; 'ż'='z'; 'ź'='z'}
Foreach ($key in $hash.keys) {
    $SamAccountUser = $SamAccountUser.Replace($key, $hash.$key)
    $SamAccountUser_email = $SamAccountUser_email.Replace($key, $hash.$key)

  }


############ find manager ##############

$i=0

do {$i++; $star="*"; $searchMG2 = "";
        $searchMG = $(Write-Host "MENADŻER - SZUKAJ: " -ForegroundColor Cyan -NoNewLine; Read-Host);
                $searchMG2 = $star+$searchMG+$star;
        if ([string]::IsNullOrWhiteSpace($searchMG)) { Write-Warning "Błędne dane wejściowe" }
                
                    else {$pathMG = @(Get-ADUser -Filter { Displayname -like $searchMG2 } -Properties * -Credential $Credentials -Server $Server -SearchBase $pathOU  | Select-Object Surname, Givename, EmailAddress, SamAccountName, Displayname, Title)
        if ($searchMG2.Contains([string]$pathMG.displayname) = $false) { Write-Warning "Błędne dane wejściowe" }
            else {break}}
}while ($i -le 2 )

if ($searchMG2.Contains([string]$pathMG.displayname) = $false) { Write-Host -ForegroundColor DarkRed "Kończenie pracy.."; [console]::beep(700,1000) ; Start-Sleep -Seconds 2; break}

[int]$numermg = 0
$bazamg =@{}
foreach ($_ in $pathMG ) { $numermg++
    Write-Host -ForegroundColor DarkCyan $numermg. $_.Displayname "," $_.Title;
    $bazamg[$numermg]= $_.Surname, $_.Givename, $_.EmailAddress, $_.SamAccountName, $_.Displayname }

$i=0
do { $i++
     $itmg = $(Write-Host "WYBIERZ MENADŻERA: " -ForegroundColor Cyan -NoNewLine; Read-Host)
     if ( $bazamg.Keys -eq $itmg) { break}
     else {Write-Warning "Błędne dane wejściowe"}
}
while ($i -le 2)
if ($itmg -notin $bazamg.Keys) { Write-Host -ForegroundColor DarkRed "Kończenie pracy.."; [console]::beep(700,1000) ; Start-Sleep -Seconds 2; break}

############ manager single variable ##############

$Manager = $bazamg.Item([int]$itmg)
#$Manager[0] - Surname
#$Manager[1] - Givename
#$Manager[2] - EmailAddress
#$Manager[3] - SamAccountName
#$Manager[4] - Displayname

############ password generator ##############

function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
    $charSet = 'abcdefghi+*=:^%!&?.vwxyzABCDEFG$$$$$$abcdefghHIJKLMfghijkNOPQRS$$$TUVWXYZ01234567890123456789+=+=+=+=+=+'.ToCharArray()
    #$charSet = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
 
    $rng.GetBytes($bytes)
 
    $result = New-Object char[]($length)
 
    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }
 
    return (-join $result)
}
$Pass = Get-RandomPassword 15

$licz=0

if ([string]::IsNullOrWhiteSpace($Firstname) -or ([string]::IsNullOrWhiteSpace($Lastname))) {
    Write-Host -ForegroundColor Red "Pola Imię użytkownika, Nazwisko użytkownika nie mogą być puste"
    $licz++}
if ([string]::IsNullOrWhiteSpace($Title) -or ([string]::IsNullOrWhiteSpace($Department))) {
    Write-Host -ForegroundColor Red "Pola Stanowisko, Dział nie mogą być puste"
    $licz++}
if ([string]::IsNullOrWhiteSpace($AccountExpirationDate))  {
    Write-Host -ForegroundColor Red "Pole Wygaśnięcie konta nie może być puste"
    $licz++}


$check_enabled = (Get-ADUser -Filter {enabled -eq $true} -Properties SamAccountName -Credential $Credentials -Server $Server )
$baza = $check_enabled.samaccountname

if ($baza -eq $SamAccountUser) {
    Write-host -ForegroundColor Red "Obiekt - $SamAccountUser - jest już w Active Directory !" ; [console]::beep(1200,300);[console]::beep(1200,300);[console]::beep(1200,400);[console]::beep(700,400)
    $licz++}


if ($licz -cge 1) {
$next = $(Write-Host "Czy uruchomić ponownie? T [Tak] " -ForegroundColor DarkGreen -NoNewline; Read-Host)
if ( $next -eq "t" ) { C:\Scripts\ADDNewUSER.ps1 }
else {break} 
}

#Added one day  to "AccountExpirationDate" variable and hide error after input incorectyly data

  try { if ($AccountExpirationDate -ne "n") {$ExpirationDateOneDay = [DateTime]::Parse($AccountExpirationDate).AddDays(1)  } }
  
  catch [System.Management.Automation.MethodInvocationException],[System.Management.Automation.CommandNotFoundException] { Write-Warning -Message "Pole Wygaśnięcie konta zawiera niepoprawny format"  

$next2 =  $(Write-Host "Czy uruchomić ponownie? T [Tak] " -ForegroundColor DarkGreen -NoNewline; Read-Host)
if ( $next1 -eq "t" ) {  C:\Scripts\ADDNewUSER.ps1 }
else {break}
  }
    
 New-ADUser `
            -SamAccountName $SamAccountUser `
            -UserPrincipalName $SamAccountUser@wasko.pl `
            -Credential $Credentials `
            -Server $Server `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Displayname "$Firstname $Lastname" `
            -Description "$Department, $Firstname $Lastname" `
            -Department $Department `
            -Path $mainpath `
            -EmailAddress $SamAccountUser_email@fonon.com.pl `
            -Title $Title `
            -Manager $Manager[3] `
            -ChangePasswordAtLogon $true `
            -StreetAddress "Pańska 23" `
            -City "Kraków" `
            -Company "Fonon Sp. z o.o." `
            -Country "pl" `
            -PostalCode "30-565" `
            -AccountExpirationDate $ExpirationDateOneDay `
            -Enabled $true `
            -AccountPassword (ConvertTo-SecureString "$Pass" -AsPlainText -Force)
          
     if ($AccountExpirationDate -eq "n") {Clear-ADAccountExpiration -Identity $SamAccountUser -Credential $Credentials -Server $Server}
           
           Write-Output ""
           
        if ($Add0 -eq "t")
       {
       Add-ADGroupMember -Identity "HCPAW-RegularUsers" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }
        if ($Add1 -eq "t")
       {
       Add-ADGroupMember -Identity "O365SyncUser" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }
        if ($Add2 -eq "t")
       {
       Add-ADGroupMember -Identity "ALFA_ProjektyFONON_100016_NOKIA-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }
        if ($Add3 -eq "t")
       {
       Add-ADGroupMember -Identity "ALFA_ProjektyFONON_100026_GSMR-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }
        if ($Add4 -eq "t")
       {
       Add-ADGroupMember -Identity "ALFA_ProjektyFONON_Pomoce_naukowe-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }
        if ($Add5 -eq "t")
       {
       Add-ADGroupMember -Identity "ALFA_ProjektyFONON_Niemcy_projektowanie-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
       }

       
    Write-Output ""
    Write-host -ForegroundColor Cyan "Dodano użytkownika '$SamAccountUser'"
    Write-host -ForegroundColor Yellow "Dane zostały zapisane: C:\NewUsersList.txt"
    Write-Output "$Date" | Out-File  C:\NewUsersList.txt -append
    Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  C:\NewUsersList.txt -append
    Write-Output "Użytkownik: DOMENA-WASKO\$SamAccountUser" | Out-File  C:\NewUsersList.txt -append
    Write-Output "E-mail: $SamAccountUser_email@fonon.com.pl" | Out-File  C:\NewUsersList.txt -append
    Write-Output "Hasło do zmiany: $Pass" | Out-File  C:\NewUsersList.txt -append
    Write-Output "E-mail Menadżera: $($Manager[2])" | Out-File  C:\NewUsersList.txt -append
    Write-Output "ORGANIZATION UNIT: $mainpath" | Out-File  C:\NewUsersList.txt -append
    Write-Output "============================================" | Out-File  C:\NewUsersList.txt -append



#Create new account email on mx.coig.pl

$mx_coig = $(Write-Host "Smarter-mail: Utworzyć konto $SamAccountUser_email ? T [Tak]: " -ForegroundColor DarkYellow -NoNewline; Read-Host);

if ($mx_coig -eq "t") {

$Cred = Get-StoredCredential -Target mxfonon
$password = $Cred.GetNetworkCredential().Password
$username = $Cred.GetNetworkCredential().UserName

$authBody = @{
username = $Username
password = $password }

$response = Invoke-RestMethod -Uri https://mx.coig.pl/api/v1/auth/authenticate-user -Method Post -Body $authBody

$header = @{
'Authorization' = "Bearer " + $response.accessToken
'Content'='application/json' }
 
$SamAccount = "$SamAccountUser_email"
$Password = "$Pass"
$mx = "$SamAccountUser_email@owa.wasko.pl"
$FullName = "$Firstname $Lastname"

$user = @('{
"userData": {
		"userName": "' + $SamAccount + '",
        "fullName": "' + $FullName + '",
		"password": "' + $Password + '",

    "securityFlags": {
		"authType": 0,
		"authenticatingWindowsDomain": null,
		"isDomainAdmin": false,
        "isDisabled": false
    },
	 "isPasswordExpired": false
    },

"forwardList": {
     "forwardList":
       [
        "' + $mx + '",
       ],

     "keepRecipients": false,
     "deleteOnForward": true
    },

"userMailSettings": {
     "canReceiveMail": true,
     "enableMailForwarding": true
    },
}')

Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/user-put -Method Post -Headers $header -body $user -ContentType "application/json"
}



#File to sending via mail
$path_notepad = "C:\$Firstname $Lastname - Poświadczenia.txt"

    Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  "$path_notepad" -Append
    Write-Output "Użytkownik: DOMENA-WASKO\$SamAccountUser" | Out-File  "$path_notepad" -Append
    Write-Output "E-mail: $SamAccountUser_email@fonon.com.pl" | Out-File  "$path_notepad" -Append
    Write-Output "Hasło do zmiany: $Pass" | Out-File  "$path_notepad" -Append


#variable for it account to send email like CC

$itusername = ($username.split("@")[0])

#Sending email to manager with the user credentials

$Sendmail = Read-Host "Wysłać poświadczenia do $($Manager[2]) ? T [Tak]"
if ($sendmail -eq "t") {
$outlook = new-object -comobject outlook.application
$email = $outlook.CreateItem(0)
$email.Recipients.Add([string]$Manager[2]) > $null
$recip = $email.Recipients.Add("$itusername@wasko.pl")
$recip.Type = 3 #CC = 2 #
$email.Subject = "Nowy użytkownik $Firstname $Lastname"
$email.Body = "W załączniku poświadczenia dla użytkownika $Firstname $Lastname" 
$email.Attachments.add($path_notepad) > $null
$email.Send()
}


$Sendmail2 = Read-Host "Wysłać poświadczenia do innej osoby ? T [Tak]"
if ($sendmail2 -eq "t") {
$NextUserMail = Read-Host "Podaj adres email"
$outlook = new-object -comobject outlook.application
$email = $outlook.CreateItem(0)
$email.To = "$NextUserMail"
$email.Subject = "Nowy użytkownik $Firstname $Lastname"
$email.Body = "W załączniku poświadczenia dla użytkownika $Firstname $Lastname" 
$email.Attachments.add($path_notepad) > $null
$email.Send()
}

Remove-Item -Path "$path_notepad"

$next1 =  $(Write-Host "Czy uruchomić ponownie? T [Tak] " -ForegroundColor DarkGreen -NoNewline; Read-Host)
if ( $next1 -eq "t" ) {  C:\Scripts\ADDNewUSER.ps1 }
else {break}