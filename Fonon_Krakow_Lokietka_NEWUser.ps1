Cls
Write-Output "FONON KRAKOW ŁOKIETKA"
$mainpath=", OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
$path=Read-host "(OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl  >>>  np:'OU=ZPDO, OU=POPC2')"
$Firstname = Read-Host "Imię użytkownika"
$Lastname = Read-Host "Nazwisko użytkownika"
$Firstname_manager = Read-Host "Imię menadżera"
$Lastname_manager = Read-Host "Nazwisko menadżera"
$Title = Read-Host "Stanowisko"
$Department = Read-Host "Dział"
$AccountExpirationDate=Read-Host "Data wygaśnięcia konta: (dd-mm-rrrr + 1 dzień)"
$Firstname_user = $Firstname[0]
$Firstname_manager1 = $Firstname_manager[0]
$SamAccountUser = ($Firstname_user + "." + $Lastname).ToLower()
$SamAccountUser_email = ($Firstname_user + "." + $Lastname).ToLower()
$Manager = ($Firstname_manager1 + "." + $Lastname_manager).ToLower()
$Manager_email = ($Firstname_manager1 + "." + $Lastname_manager).ToLower()

$Add0=Read-Host "Dodać do grupy:  HCPAW-RegularUsers: Wybierz T [tak] lub ENTER [dalej]"
$Add1=Read-Host "Dodać do grupy:  O365SyncUser: Wybierz T [tak] lub ENTER [dalej]"
$Add2=Read-Host "Dodać do grupy:  ALFA_ProjektyFONON_100016_NOKIA-RW: Wybierz T [tak] lub ENTER [dalej]"
$Add3=Read-Host "Dodać do grupy:  ALFA_ProjektyFONON_100026_GSMR-RW: Wybierz T [tak] lub ENTER [dalej]"
$Add4=Read-Host "Dodać do grupy:  ALFA_ProjektyFONON_Pomoce_naukowe-RW: Wybierz T [tak] lub ENTER [dalej]"
$Add5=Read-Host "Dodać do grupy:  ALFA_ProjektyFONON_Niemcy_projektowanie-RW: Wybierz T [tak] lub ENTER [dalej]"

$hash = @{'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ń'='n'; 'ó'='o'; 'ś'='s'; 'ż'='z'; 'ź'='z'}
Foreach ($key in $hash.keys) {
    $SamAccountUser = $SamAccountUser.Replace($key, $hash.$key)
    $SamAccountUser_email = $SamAccountUser_email.Replace($key, $hash.$key)
    $Manager = $Manager.Replace($key, $hash.$key)
    $Manager_email = $Manager_email.Replace($key, $hash.$key)
 }


function Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
    $charSet = 'abcdefghi+*=:^%!&?.vwxyzABCDEFG$$$$abcdefghHIJKLMfghijkNOPQRS$$$TUVWXYZ01234567890123456789+*=:^%!&?.'.ToCharArray()
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

$Pass = Get-RandomPassword 12

if ([string]::IsNullOrWhiteSpace($SamAccountUser)) {
    Write-Host -ForegroundColor Red "Pola 'Imię użytkownika, Nazwisko użytkownika' nie mogą być puste"
    }
  
    elseif (Get-ADUser -Filter {SamAccountName -eq $SamAccountUser}) {
               Write-host -ForegroundColor Yellow "Użytkownik '$SamAccountUser' jest w AD"
       }

       else
       {
 New-ADUser `
            -SamAccountName $SamAccountUser `
            -UserPrincipalName $SamAccountUser@wasko.pl `
            -Name "$Firstname $Lastname" `
            -GivenName $Firstname `
            -Surname $Lastname `
            -Displayname "$Firstname $Lastname" `
            -Description "$Department, $Firstname $Lastname" `
            -Department $Department `
            -Path "$path $mainpath" `
            -EmailAddress $SamAccountUser_email@fonon.com.pl `
            -Title $Title `
            -Manager $Manager `
            -ChangePasswordAtLogon $true `
            -StreetAddress "Łokietka 79" `
            -City "Kraków" `
            -Company "Fonon Sp. z o.o." `
            -Country "pl" `
            -PostalCode "31-280" `
            -AccountExpirationDate $AccountExpirationDate `
            -Enabled $true `
            -AccountPassword (ConvertTo-SecureString "$Pass" -AsPlainText -Force) |
         
 Add-ADGroupMember -Identity "O365SyncUser" -Members $UserName

    Write-Output ""
    Write-host -ForegroundColor Cyan "Dodano użytkownika '$SamAccountUser'"
    Write-host -ForegroundColor Yellow "Dane zostały zapisane: C:\Scripts\New_Users.txt"
    Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  C:\Scripts\New_Users.txt -append
    Write-Output "Użytkownik: DOMENA-WASKO\$SamAccountUser" | Out-File  C:\Scripts\New_Users.txt -append
    Write-Output "E-mail: $SamAccountUser_email@fonon.com.pl" | Out-File  C:\Scripts\New_Users.txt -append
    Write-Output "Hasło do zmiany: $Pass" | Out-File  C:\Scripts\New_Users.txt -append
    Write-Output "E-mail Menadżera: $Manager_email@fonon.com.pl" | Out-File  C:\Scripts\New_Users.txt -append
    Write-Output "" | Out-File  C:\Scripts\New_Users.txt -append
       
       }

$next = Read-Host -Prompt "Dodaj użytkownika 't' lub zakończ ENTER"
        if ($next -eq "t") {
        C:\Scripts\ADDNewUSER.ps1
        }