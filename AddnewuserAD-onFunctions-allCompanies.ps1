function GLOBAL:Get-DeclaringVariables {

    $GLOBAL:Date = Get-Date
    $GLOBAL:Manager = ""
    $GLOBAL:pathOU = ""
    $GLOBAL:Firstname = Read-Host "Imię użytkownika"
    $GLOBAL:Lastname = Read-Host "Nazwisko użytkownika"
    $GLOBAL:Title = Read-Host "Stanowisko"
    $GLOBAL:Department = Read-Host "Dział"
    $GLOBAL:AccountExpirationDate = Read-Host "Wygaśnięcie konta: N [Never] lub [dd-mm-rr]"
}

function GLOBAL:Set-ChangePolishSign {
    $hash = @{'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ń'='n'; 'ó'='o'; 'ś'='s'; 'ż'='z'; 'ź'='z'}
    foreach ($key in $hash.Keys) {
    $GLOBAL:SamAccountUser = $SamAccountUser.Replace($key, $hash[$key])
    $GLOBAL:SamAccountUserEmail = $SamAccountUserEmail.Replace($key, $hash[$key])
    }
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountWasko {
    
    $GLOBAL:HCPAWRegularUsers = Read-Host "Dodać do grupy - HCPAW-RegularUsers? T [Tak]"
    $GLOBAL:O365SyncUser = Read-Host "Dodać do grupy - O365SyncUser? T [Tak]"
    $GLOBAL:SamAccountUser = ($Firstname[0] + "." + $Lastname).ToLower()
    $GLOBAL:SamAccountUserEmail = ($Firstname[0] + "." + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountEnte {

    $LastnameChar8 = $Lastname[0..6] -join ''.ToLower()
    $GLOBAL:SamAccountUser = ($Firstname[0] + $LastnameChar8)
    $GLOBAL:SamAccountUserEmail = ($Firstname[0] + "." + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountFonon {
    
    $GLOBAL:HCPAWRegularUsers = Read-Host "Dodać do grupy - HCPAW-RegularUsers? T [Tak]"
    $GLOBAL:O365SyncUser = Read-Host "Dodać do grupy - O365SyncUser? T [Tak]"
    $GLOBAL:100016 = Read-Host "Dodać do grupy - ALFA_ProjektyFONON_100016_NOKIA-RW ? T [Tak]"
    $GLOBAL:100026 = Read-Host "Dodać do grupy - ALFA_ProjektyFONON_100026_GSMR-RW ? [Tak]"
    $GLOBAL:PomoceNaukoweRW = Read-Host "Dodać do grupy - ALFA_ProjektyFONON_Pomoce_naukowe-RW ? T [Tak]"
    $GLOBAL:NiemcyProjektowanieRW = Read-Host "Dodać do grupy - ALFA_ProjektyFONON_Niemcy_projektowanie-RW ? T [Tak]"
    $GLOBAL:SamAccountUser = ($Firstname[0] + "." + $Lastname).ToLower()
    $GLOBAL:SamAccountUserEmail = ($Firstname[0] + "." + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountCoig {
    $azureADUsers = Read-Host "Dodać do grupy - AzureADUsers ? T [Tak]"
    $vpnCoig = Read-Host "Dodać do grupy - VPN-COIG ? T [Tak]"
    $SamAccountUser = ($Firstname[0] + $Lastname).ToLower()
    $SamAccountUserEmail = ($Firstname + "." + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountD2s {
    
    $D2SUsers = Read-Host "Dodać do grupy - D2S Users ? T [Tak]"
    $HCPAWRegularUsers = Read-Host "Dodać do grupy - HCPAW-RegularUsers? T [Tak]"
    $O365SyncUser = Read-Host "Dodać do grupy - O365SyncUser? T [Tak]"
    $WsparcieProjektowD2s = Read-Host "Dodać do grupy - D2S_Dzial_Wsparcia_Projektów ? T [Tak]"
    $LastnameChar8 = $Lastname[0..6] -join ''.ToLower()
    $GLOBAL:SamAccountUser = ($Firstname + $LastnameChar8)
    $GLOBAL:SamAccountUserEmail = ($Firstname + "." + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountGabos {
    
    $ADAzureSync = Read-Host "Dodać do grupy - ADAzureSync ? T [Tak]"
    $HCPAWUsers = Read-Host "Dodać do grupy - HCPAW Users ? T [Tak]"
    $GLOBAL:SamAccountUser = ($Firstname[0] + $Lastname).ToLower
    $GLOBAL:SamAccountUserEmail = ($Firstname[0] + $Lastname).ToLower()
}

function GLOBAL:Get-DeclaringVariablesGroupAndAccountFk {
    
    $hcp = Read-Host "Dodać do grupy - HCP ? T [Tak]"
    $vpnfk = Read-Host "Dodać do grupy - VPN_FK ? T [Tak]"
    $accountan = Read-Host "Dodać do grupy - Księgowość ? T [Tak]"
    $LastnameChar8 = $Lastname[0..6] -join ''.ToLower()
    $GLOBAL:SamAccountUser = ($Firstname[0] + $LastnameChar8)
    $GLOBAL:SamAccountUserEmail = ($Firstname[0] + "." + $Lastname).ToLower()
}

function Get-AddUserToGroupMember {
    if ($AccountExpirationDate -eq "n") {
        Clear-ADAccountExpiration -Identity $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($HCPAWRegularUsers -eq "t") {
        Add-ADGroupMember -Identity "HCPAWRegularUsers" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($O365SyncUser -eq "t") {
        Add-ADGroupMember -Identity "O365SyncUser" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($100016 -eq "t") {
        Add-ADGroupMember -Identity "ALFA_ProjektyFONON_100016_NOKIA-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($100026 -eq "t") {
        Add-ADGroupMember -Identity "ALFA_ProjektyFONON_100026_GSMR-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($PomoceNaukowe -eq "t") {
        Add-ADGroupMember -Identity "ALFA_ProjektyFONON_Pomoce_naukowe-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($NiemcyProjektowanieRW -eq "t") {
        Add-ADGroupMember -Identity "ALFA_ProjektyFONON_Niemcy_projektowanie-RW" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($azureADUsers -eq "t") {
        Add-ADGroupMember -Identity "AzureADUsers" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($vpnCoig -eq "t") {
        Add-ADGroupMember -Identity "VPN-COIG" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($D2SUsers -eq "t") {
        Add-ADGroupMember -Identity "D2S Users" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($WsparcieProjektowD2s -eq "t") {
        Add-ADGroupMember -Identity "D2S_Dzial_Wsparcia_Projektów" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($hcp -eq "t") {
        Add-ADGroupMember -Identity "HCP" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($vpnfk -eq "t") {
        Add-ADGroupMember -Identity "VPN_FK" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
    if ($accountan -eq "t") {
        Add-ADGroupMember -Identity "Księgowość" -Members $SamAccountUser -Credential $Credentials -Server $Server
    }
}

function GLOBAL:Get-RandomPassword {
    param (
        [Parameter(Mandatory)]
        [int] $length
    )
    #$charSet = 'abcdefghi+*=:^%!&?.vwxyzABCDEFG$$$$$$abcdefghHIJKLMfghijkNOPQRS$$$TUVWXYZ01234567890123456789+=+=+=+=+=+'.ToCharArray()
    $charSet = '!@#$%^&*()qwertasdfgzxcvb0123456789ASDFGZXCVBQWERT0123456789+=_+=_+=_+=_+=_+=_'.ToCharArray()
    $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
    $bytes = New-Object byte[]($length)
 
    $rng.GetBytes($bytes)

    $result = New-Object char[]($length)

    for ($i = 0 ; $i -lt $length ; $i++) {
        $result[$i] = $charSet[$bytes[$i]%$charSet.Length]
    }

    return (-join $result)
}


function Set-CredentialAD {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Target,
        [Parameter(Mandatory=$true)]
        [string]$Server,
        [Parameter(Mandatory=$true)]
        [string]$pathOU
    )
    
    Process {
        
        $GLOBAL:Credentials = Get-StoredCredential -Target "$Target"
        $GLOBAL:Username = $Credentials.UserName
        $GLOBAL:Password = $Credentials.Password
        $GLOBAL:Server = "$Server"
        $GLOBAL:Credentials = New-Object System.Management.Automation.PSCredential $Username, $Password
        $GLOBAL:pathOU = "$pathOU"
        Get-ADOrganizationalUnit -Filter * -Credential $Credentials -Server $Server -SearchBase $pathOU | Select-Object -ExpandProperty DistinguishedName
    }
}


function GLOBAL:Get-OrganizationUnit {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$false)]
        [string]$Server,
        [Parameter(Mandatory=$false)]
        [string]$pathOU
    )
    
    Process {$i = 0
    do {
        $i++
        $star = "*"
        $section = $(Write-Host "ORGANIZATION UNIT: " -ForegroundColor Cyan -NoNewLine; Read-Host);
        $organizationUnit = $section + $star
        $paths = @(Get-ADOrganizationalUnit -Filter {Name -like $organizationUnit} -Credential $Credentials -Server $server -SearchBase $pathOU | Select-Object -ExpandProperty distinguishedname)
        
        if (-not $paths) {
            Write-Warning "Błędne dane wejściowe"
        }
        else {
        break
        
        }
    } while ($i -le 2)

    if (-not $paths) {
        Write-Host -ForegroundColor DarkRed "Kończenie pracy.."
        [console]::beep(700, 1000)
        Start-Sleep -Seconds 2
        break
    }
    
    $number = 0
    $numberOfTables = @{}
    foreach ($path in $paths) {
        $number++
        Write-Host -ForegroundColor DarkCyan "$number. $path"
        $numberOfTables[$number] = $path
    }

        $i = 0
    do {
        $i++
        try {
            $organizationSelection = $(Write-Host "WYBIERZ OU: " -ForegroundColor Cyan -NoNewLine; Read-Host)
            if ($numberOfTables.ContainsKey([int]$organizationSelection)) {
                break
        
            }
            elseif ($numberOfTables.ContainsKey([int]$organizationSelection) -eq $false) {
                Write-Warning "Błędne dane wejściowe"
            }
        }
        catch {
            Write-Warning "Błędne dane wejściowe"
        }
    } while ($i -le 2)


    if ($organizationSelection -notin $numberOfTables.Keys) {
        Write-Host -ForegroundColor DarkRed "Kończenie pracy.."
        [console]::beep(700, 1000)
        Start-Sleep -Seconds 2
        break
    }

    $GLOBAL:mainPath = $numberOfTables.Item([int]$organizationSelection)
    Clear-Host
    Write-Host -ForegroundColor DarkGray "Organization Unit: $mainPath"

    }
}
    

function Get-CheckInputData {

    $i = 0
    if ([string]::IsNullOrWhiteSpace($Firstname) -or [string]::IsNullOrWhiteSpace($Lastname)) {
        Write-Host -ForegroundColor Red "Pola Imię użytkownika, Nazwisko użytkownika nie mogą być puste"
        $i++
    }

    if ([string]::IsNullOrWhiteSpace($Title) -or [string]::IsNullOrWhiteSpace($Department)) {
        Write-Host -ForegroundColor Red "Pola Stanowisko, Dział nie mogą być puste"
        $i++
    }

    if ([string]::IsNullOrWhiteSpace($AccountExpirationDate)) {
        Write-Host -ForegroundColor Red "Pole Wygaśnięcie konta nie może być puste"
        $i++
    }

    $checkIfEnabled = Get-ADUser -Filter {enabled -eq $true} -Properties SamAccountName -Credential $Credentials -Server $Server
    $checkIfEnabledData = $checkIfEnabled.samaccountname

    if ($checkIfEnabledData -eq $SamAccountUser) {
        Write-Host -ForegroundColor Red "Obiekt - $SamAccountUser - jest już w Active Directory!"
        [console]::beep(1200, 300)
        [console]::beep(1200, 300)
        [console]::beep(1200, 400)
        [console]::beep(700, 400)
        $i++
    }
        $onceMoreTurnOn = ""
    if ($i -ge 1) {
        $onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
        if ($onceMoreTurnOn -eq "t") {
            C:\Scripts\ADDNewUSER.ps1
        }
        else {
            break
        }
    }

    try {
        if ($AccountExpirationDate -ne "n") {
            $ExpirationDateOneDay = [DateTime]::Parse($AccountExpirationDate).AddDays(1)
        }
    }
    catch [System.Management.Automation.MethodInvocationException],[System.Management.Automation.CommandNotFoundException] {
        Write-Warning -Message "Pole Wygaśnięcie konta zawiera niepoprawny format"
        $onceMoreTurnOn = ""
        $onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
        if ($onceMoreTurnOn -eq "t") {
            C:\Scripts\ADDNewUSER.ps1
        }
        else {
            C:\Scripts\ADDNewUSER.ps1
        $onceMoreTurnOn = ""
        $onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
        if ($onceMoreTurnOn -eq "t") {
            C:\Scripts\ADDNewUSER.ps1
        }
        }
    }
}

function GLOBAL:Get-AddNewUser {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$StreetAddress,
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$PostalCode,
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Department,
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Description,
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$City,
        [Parameter(Mandatory=$false)]
        [string]$Company,
        [Parameter(Mandatory=$false)]
        [string]$mainPath,
        [Parameter(Mandatory=$true)]
        [string]$DomainMail,
        [Parameter(Mandatory=$true)]
        [string]$Server
    )
    
    Process {
       New-ADUser `
        -SamAccountName $SamAccountUser `
        -UserPrincipalName "$SamAccountUser@$Server" `
        -Credential $Credentials `
        -Server $Server `
        -Name "$Firstname $Lastname" `
        -GivenName $Firstname `
        -Surname $Lastname `
        -Displayname "$Firstname $Lastname" `
        -Description "$Description" `
        -Department $Department `
        -Path $mainPath `
        -EmailAddress "$SamAccountUserEmail@$DomainMail" `
        -Title $Title `
        -Manager $Manager[3] `
        -ChangePasswordAtLogon $true `
        -StreetAddress $StreetAddress `
        -City $City `
        -Company $Company `
        -Country "PL" `
        -PostalCode $PostalCode `
        -AccountExpirationDate $ExpirationDateOneDay `
        -Enabled $true `
        -AccountPassword (ConvertTo-SecureString "$Pass" -AsPlainText -Force)
    }
}

function GLOBAL:Get-ChoiseManagerInOrganizationUnit {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Server,
        [Parameter(Mandatory=$true)]
        [string]$pathOU
    )
    
    Process {
    $i = 0
    do {
        $i++
        $star = "*"
        $nameManager = $(Write-Host "MENADŻER - SZUKAJ: " -ForegroundColor Cyan -NoNewLine; Read-Host);
        $starPlusManager = $star + $nameManager + $star

        if ([string]::IsNullOrWhiteSpace($nameManager)) {
            Write-Warning "Błędne dane wejściowe"
        }
        else {
            $findManagerInOu = @(Get-ADUser -Filter {DisplayName -like $starPlusManager} -Properties Surname, GivenName, EmailAddress, SamAccountName, Displayname, Title -Credential $Credentials -Server $Server -SearchBase $pathOU | Select-Object Displayname, Title, Surname, GivenName, EmailAddress, SamAccountName)

            if (-not $findManagerInOu) {
                Write-Warning "Błędne dane wejściowe"
            }
            else {
                break
        $onceMoreTurnOn = ""
        $onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
        if ($onceMoreTurnOn -eq "t") {
            C:\Scripts\ADDNewUSER.ps1
        }
            }
        }
    } while ($i -le 2)


    if (-not $findManagerInOu) {
        Write-Host -ForegroundColor DarkRed "Kończenie pracy.."
        [console]::beep(700, 1000)
        Start-Sleep -Seconds 2
        break
    }

    [int]$numberManager = 0
    $numberOfTablesManager = @{}
    foreach ($manager in $findManagerInOu) {
        $numberManager++
        Write-Host -ForegroundColor DarkCyan "$numberManager. $($manager.Displayname), $($manager.Title)"
        $numberOfTablesManager[$numberManager] = $manager.Surname, $manager.GivenName, $manager.EmailAddress, $manager.SamAccountName, $manager.Displayname
    }


        $i = 0
       $i = 0
    do {
        $i++
        try {
            $choiseManager = Read-Host "WYBIERZ MENADŻERA"
            if ($numberOfTablesManager.ContainsKey([int]$choiseManager)) {
                break
        $onceMoreTurnOn = ""
        $onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
        if ($onceMoreTurnOn -eq "t") {
            C:\Scripts\ADDNewUSER.ps1
        }
            }
            else {
                Write-Warning "Błędne dane wejściowe"
            }
        }
        catch {
            Write-Warning "Błędne dane wejściowe"
        }
    } while ($i -le 2)
    
    if ($choiseManager -notin $numberOfTablesManager.Keys) {
        Write-Host -ForegroundColor DarkRed "Kończenie pracy.."
        [console]::beep(700, 1000)
        Start-Sleep -Seconds 2
        break 

    }
    $GLOBAL:Manager = $numberOfTablesManager.Item([int]$choiseManager)
    }
}

function GLOBAL:Get-WriteToFileNewUsersList {

        [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Domain,
        [Parameter(Mandatory=$true)]
        [string]$DomainMail

    )
    
    Process {
        Write-Output ""
        Write-host -ForegroundColor Cyan "Dodano użytkownika '$SamAccountUser'"
        Write-host -ForegroundColor Yellow "Dane zostały zapisane: C:\NewUsersList.txt"
        Write-Output ""
        Write-Output "$Date" | Out-File  C:\NewUsersList.txt -append
        Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File  C:\NewUsersList.txt -append
        Write-Output "Użytkownik: $Domain\$SamAccountUser" | Out-File  C:\NewUsersList.txt -append
        Write-Output "E-mail: $SamAccountUserEmail@$DomainMail" | Out-File  C:\NewUsersList.txt -append
        Write-Output "Hasło do zmiany: $Pass" | Out-File  C:\NewUsersList.txt -append
        Write-Output "E-mail Menadżera: $($Manager[2])" | Out-File  C:\NewUsersList.txt -append
        Write-Output "ORGANIZATION UNIT: $mainPath" | Out-File  C:\NewUsersList.txt -append
        Write-Output "============================================" | Out-File  C:\NewUsersList.txt -append
    }
}

function GLOBAL:Set-CredentialMxCoig {
    
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Target
    )
    
    Process {
        $GLOBAL:Cred = Get-StoredCredential -Target "$Target"
        $GLOBAL:passwordMx = $Cred.GetNetworkCredential().Password
        $GLOBAL:userNameMx = $Cred.GetNetworkCredential().UserName
    }
}

function GLOBAL:Get-CreateTempFileToSendViaMail {

        [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$Domain,
        [Parameter(Mandatory=$true)]
        [string]$DomainMail
    )

    Process {
        $GLOBAL:tempNoteFile = "C:\$Firstname $Lastname - Poświadczenia.txt"
    
        Write-Output "============================================" | Out-File "$tempNoteFile" -Append
        Write-Output "Imię i Nazwisko: $Firstname $Lastname" | Out-File "$tempNoteFile" -Append
        Write-Output "Użytkownik: $Domain\$SamAccountUser" | Out-File "$tempNoteFile" -Append
        Write-Output "E-mail: $SamAccountUserEmail@$DomainMail" | Out-File "$tempNoteFile" -Append
        Write-Output "Hasło do zmiany: $Pass" | Out-File "$tempNoteFile" -Append
        Write-Output "============================================" | Out-File "$tempNoteFile" -Append
    }
}

function GLOBAL:Get-AddNewAccountSmarterMail {

    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    Param (   
        [Parameter(Mandatory=$true)]
        [string]$enableMailForwarding,
        [Parameter(Mandatory=$false)]
        [string]$owaDomainAddress

    )
    
    Process {
$mxCoigAccount = $(Write-Host "Smarter-Mail: Utworzyć konto $SamAccountUserEmail ? T [Tak]: " -ForegroundColor DarkYellow -NoNewline; Read-Host)

    if ($mxCoigAccount -eq "t") {
        $authBody = @{
            username = $userNameMx
            password = $passwordMx
        }
        
        $response = Invoke-RestMethod -Uri https://mx.coig.pl/api/v1/auth/authenticate-user -Method Post -Body $authBody
        
        $header = @{
            'Authorization' = "Bearer " + $response.accessToken
        }
        
        $samAccount = "$SamAccountUserEmail"
        $fullName = "$Firstname $Lastname"
        $passwordHead = "$Pass"
        $exchangeOwaAddress = "$SamAccountUserEmail@$owaDomainAddress"
        
        
    $user = @('{
    "userData": {
		    "userName": "' + $SamAccount + '",
            "fullName": "' + $FullName + '",
		    "password": "' + $PasswordHead + '",

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
            "' + $exchangeOwaAddress + '",
           ],

         "keepRecipients": false,
         "deleteOnForward": true
        },

    "userMailSettings": {
         "canReceiveMail": true,
         "enableMailForwarding": "' + $enableMailForwarding + '"
        },
    }')

        Invoke-RestMethod -uri https://mx.coig.pl/api/v1/settings/domain/user-put -Method Post -Headers $header -body $user -ContentType "application/json" | Out-Null
        }
    }
}

function GLOBAL:Get-SendEmailToManager {
    $GLOBAL:itUserName = ($userNameMx.split("@")[0])
    $sendMail = Read-Host "Wysłać poświadczenia do $($Manager[2]) ? T [Tak]"
    if ($sendMail -eq "T") {
        try {
            $outlook = New-Object -ComObject Outlook.Application
            $email = $outlook.CreateItem(0)
            $email.Recipients.Add([string]$Manager[2]) > $null
            $recip = $email.Recipients.Add("$itUserName@wasko.pl")
            $recip.Type = 3  # CC = 2
            $email.Subject = "Nowy użytkownik $Firstname $Lastname"
            $email.Body = "W załączniku poświadczenia dla użytkownika $Firstname $Lastname"
            $email.Attachments.Add($tempNoteFile) > $null
            $email.Send() > $null
            Write-Host -ForegroundColor Black -BackgroundColor White "E-mail został wysłany"
        }
        catch {
            Write-Host -ForegroundColor Black -BackgroundColor Yellow "Nie udało się wysłać wiadomości e-mail do: $($Manager[2])"
        }
        finally {
            if ($email) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($email) | Out-Null
                Remove-Variable -Name email -ErrorAction SilentlyContinue
            }
    
            if ($outlook) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                Remove-Variable -Name outlook -ErrorAction SilentlyContinue
            }
        }
    }


    $sendMail2 = Read-Host "Wysłać poświadczenia do innej osoby? T [Tak]"
    if ($sendMail2 -eq "T") {
        try {
            $NextUserMail = Read-Host "Podaj adres e-mail"
            if ([string]::IsNullOrEmpty($NextUserMail)) {
                Write-Host -ForegroundColor Black -BackgroundColor Yellow "Nie podano adresu e-mail"
            }
            $outlook = New-Object -ComObject Outlook.Application 
            $email = $outlook.CreateItem(0)
            $email.To = "$NextUserMail"
            $email.Subject = "Nowy użytkownik $Firstname $Lastname"
            $email.Body = "W załączniku poświadczenia dla użytkownika $Firstname $Lastname"
            $email.Attachments.Add($tempNoteFile) > $null
            $email.Send() > $null
            Write-Host -ForegroundColor Black -BackgroundColor White "E-mail został wysłany"
            }
        catch {
            Write-Host -ForegroundColor Black -BackgroundColor Yellow "Nie udało się wysłać wiadomości e-mail"
        }
        finally {
            if ($email) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($email) | Out-Null
                Remove-Variable -Name email -ErrorAction SilentlyContinue 
            }
    
            if ($outlook) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
                Remove-Variable -Name outlook -ErrorAction SilentlyContinue
        }
    }
}

Remove-Item -Path "$tempNoteFile"
$onceMoreTurnOn = $(Write-Host "Czy uruchomić ponownie? T [Tak]: " -ForegroundColor DarkGreen -NoNewline; Read-Host)
if ($onceMoreTurnOn -eq "T") { & C:\Scripts\ADDNewUSER.ps1 }
}


Clear-Host

Write-Output ""
Write-Output "1. WASKO"
Write-Output "----------------------------------------"
Write-Output "2. COIG"
Write-Output "----------------------------------------"
Write-Output "3. ENTE"
Write-Output "----------------------------------------"
Write-Output "4. FONON"
Write-Output "----------------------------------------"
Write-Output "5. DE2ES"
Write-Output "----------------------------------------"
Write-Output "6. GABOS"
Write-Output "----------------------------------------"
Write-Output "7. LOGICSYNERGY"
Write-Output "----------------------------------------"
Write-Output "8. FK"
Write-Output "----------------------------------------"
Write-Output ""
Write-Output ""

[int]$GLOBAL:Company = 0
$GLOBAL:Company = Read-Host "SIEDZIBA SPÓŁKI"

if ($Company -in 1..8) {}
else {break}

    if ($Company -eq "1")
    {
        Clear-Host

        Write-Output ""
        Write-Output "1. WASKO  >>>  Gliwice - Berbeckiego"
        Write-Output "----------------------------------------"
        Write-Output "2. WASKO  >>>  Warszawa - Płowiecka"
        Write-Output "----------------------------------------"
        Write-Output "3. WASKO  >>>  Warszawa - Czackiego"
        Write-Output "----------------------------------------"
        Write-Output ""
        Write-Output ""

        $i = 0
        do {
            $i++
            try {
                $CompanyAddress = Read-Host "WYBIERZ ADRES"
                if ([int]$CompanyAddress -ge 1 -and [int]$CompanyAddress -le 3) {
                    break ; New-Item C:\Scripts\ADDNewUSER.ps1 }
                
                else {
                    Write-Warning "Błędne dane wejściowe"
                }
            }
            catch {
                Write-Warning "Błędne dane wejściowe. Wprowadź cyfrę."
            }
        } while ($i -le 2)


    if ($Company -eq "1") {

    Clear-Host
    Write-Host -ForegroundColor DarkGray "WASKO GLIWICE BERBECKIEGO"

    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=Zarzad, OU=RN, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountWasko
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Berbeckiego 6" -City "Gliwice" -Company "WASKO S.A." -PostalCode "44-100" -Server "wasko.pl" -DomainMail "wasko.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Set-CredentialMxCoig -Target "mxwasko"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false" -owaDomainAddress "owa.wasko.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Get-SendEmailToManager

    }


        if ($Company -eq "2") {

    Clear-Host
    Write-Host -ForegroundColor DarkGray "WASKO WARSZAWA PŁOWIECKA"

    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target wasko.pl -Server wasko.pl -pathOU "OU=Zarzad, OU=RN, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountWasko
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Płowiecka 105/107"  -City "Warszawa" -Company "WASKO S.A." -PostalCode "44-100" -DomainMail "wasko.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Set-CredentialMxCoig -Target "mxwasko"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false" -owaDomainAddress "owa.wasko.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Get-SendEmailToManager

    }


        if ($Company -eq "3") {

    Clear-Host
    Write-Host -ForegroundColor DarkGray  "WASKO WARSZAWA CZACKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=Zarzad, OU=RN, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountWasko
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Czackiego 7/9/11" -City "Warszawa" -Company "WASKO S.A." -PostalCode "04-501" -mainPath $mainpath -Server "DOMENA-WASKO" -DomainMail "wasko.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Set-CredentialMxCoig -Target "mxwasko"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false" -owaDomainAddress "owa.wasko.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "wasko.pl"
    Get-SendEmailToManager

    }}
    elseif ($Company -eq "2")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "COIG KATOWICE MIKOLOWSKA"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "win.coig.com" -Server "win.coig.com" -pathOU "OU=D, OU=COIG SA, DC=win, DC=coig, DC=com"
    Get-OrganizationUnit -Server "win.coig.com" -pathOU "OU=D, OU=COIG SA, DC=win, DC=coig, DC=com"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountCoig
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -Server "WINCOIG" -DomainMail "coig.pl"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "WINCOIG" -DomainMail "coig.pl"
    Set-CredentialMxCoig -Target "mxcoig"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false" -owaDomainAddress "owa.coig.pl"
    Get-CreateTempFileToSendViaMail -Domain "WINCOIG" -DomainMail "coig.pl"
    Get-SendEmailToManager
    }
    elseif ($Company -eq "3")
        {

    Clear-Host
    Write-Host -ForegroundColor DarkGray  "ENTE GLIWICE Gaudiego"
    $Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "ente.local" -Server "ente.local" -pathOU "OU=ente, DC=ente, DC=local"
    Get-OrganizationUnit -Server "ente.local" -pathOU "OU=ente, DC=ente, DC=local"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountEnte
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "ente.local" -pathOU "OU=ente, DC=ente, DC=local"
    Get-AddNewUser -Company "ENTE" -Server "ente.local" -DomainMail "ente.com.pl" -mainPath "$mainpath" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "ENTE" -DomainMail "ente.com.pl"
    Set-CredentialMxCoig -Target "mxente"
    Get-AddNewAccountSmarterMail -enableMailForwarding "true" -owaDomainAddress "owa.ente.com.pl"
    Get-CreateTempFileToSendViaMail -Domain "ENTE" -DomainMail "ente.com.pl"
    Get-SendEmailToManager
    }
    elseif ($Company -eq "4")
        {
    Clear-Host
    Write-Output ""
    Write-Output "1. FONON  >>>  Gliwice - Berbeckiego"
    Write-Output "----------------------------------------"
    Write-Output "2. FONON  >>>  Warszawa - Czackiego"
    Write-Output "----------------------------------------"
    Write-Output "3. FONON  >>>  Kraków - Pańska"
    Write-Output "----------------------------------------"
    Write-Output ""
    Write-Output ""

    $Company=Read-Host "WYBIERZ ADRES"

    if ($Company -eq "1")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "FONON GLIWICE BERBECKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountFonon
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Berbeckiego 6" -City "Gliwice" -Company "WASKO S.A." -PostalCode "44-100" -mainPath $mainpath -Server "wasko.pl" -DomainMail "fonon.com.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Set-CredentialMxCoig -Target "mxfonon"
    Get-AddNewAccountSmarterMail -enableMailForwarding "true" -owaDomainAddress "mx.fonon.com.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Get-SendEmailToManager

        }
    elseif ($Company -eq "2") {

    Clear-Host
    Write-Host -ForegroundColor DarkGray  "FONON WARSZAWA CZACKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target wasko.pl -Server wasko.pl -pathOU "OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountFonon
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Czackiego 7/9/11" -City "Warszawa" -Company "WASKO S.A." -PostalCode "04-501" -mainPath $mainpath -Server "wasko.pl" -DomainMail "fonon.com.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Set-CredentialMxCoig -Target "mxfonon"
    Get-AddNewAccountSmarterMail -enableMailForwarding "true" -owaDomainAddress "mx.fonon.com.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Get-SendEmailToManager
    }

    elseif ($Company -eq "3") {

    Clear-Host
    Write-Host -ForegroundColor DarkGray  "FONON KRAKOW PANSKA"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target wasko.pl -Server wasko.pl -pathOU "OU=Fonon, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountFonon
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Pańska 23" -City "Kraków" -Company "Fonon Sp. z o.o." -PostalCode "30-565" -mainPath $mainpath -Server "wasko.pl" -DomainMail "fonon.com.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Set-CredentialMxCoig -Target "mxfonon"
    Get-AddNewAccountSmarterMail -enableMailForwarding "true" -owaDomainAddress "mx.fonon.com.pl"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "fonon.com.pl"
    Get-SendEmailToManager
    }}
    elseif ($Company -eq "5")
        {
    cls
    Write-Output ""
    Write-Output "1. DE2ES  >>>  Gliwice - Berbeckiego"
    Write-Output "----------------------------------------"
    Write-Output "2. DE2ES  >>>  Warszawa - Czackiego"
    Write-Output "----------------------------------------"
    Write-Output "3. DE2ES  >>>  Kraków - Pańska"
    Write-Output "----------------------------------------"
    Write-Output ""
    Write-Output ""


    $Company=Read-Host "WYBIERZ ADRES"

    if ($Company -eq "1")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "DE2ES GLIWICE BERBECKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=D2S, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountD2s
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Berbeckiego 6" -City "Gliwice" -Company "D2S sp. z o.o." -PostalCode "44-100" -Department 'D2S' -Server "wasko.pl" -DomainMail "de2es.pl" -Description "D2S, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Set-CredentialMxCoig -Target "mxde2es"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Get-SendEmailToManager
    }

    elseif ($Company -eq "2")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "DE2ES WARSZAWA CZACKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=D2S, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountD2s
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Czackiego 7/9/11" -City "Warszawa" -Company "D2S sp. z o.o." -PostalCode "00-043" -Department "D2S" -Server "wasko.pl" -DomainMail "de2es.pl" -Description "D2S, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Set-CredentialMxCoig -Target "mxde2es"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Get-SendEmailToManager
        }

    elseif ($Company -eq "3")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "DE2ES KRAKOW PANSKA"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "wasko.pl" -pathOU "OU=D2S, OU=Wspolpracownicy, OU=WASKO, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountD2s
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "wasko.pl" -pathOU "OU=WASKO, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Pańska 23" -City "Kraków" -Company "D2S sp. z o.o." -PostalCode "30-565" -Server "wasko.pl" -Department "D2S" -Server "wasko.pl" -DomainMail "de2es.pl" -Description "D2S, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Set-CredentialMxCoig -Target "mxde2es"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "DOMENA-WASKO" -DomainMail "de2es.pl"
    Get-SendEmailToManager

        }
    }
    elseif ($Company -eq "6")
        {
    
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "GABOS SOFTWARE"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "win.gabos.pl" -Server "win.gabos.pl" -pathOU "OU=GABOS, DC=win, DC=gabos, DC=pl"
    Get-OrganizationUnit -Server "win.gabos.pl" -pathOU "OU=GABOS, DC=win, DC=gabos, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountGabos
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "win.gabos.pl" -pathOU "OU=GABOS, DC=win, DC=gabos, DC=pl"
    Get-AddNewUser -Server "win.gabos.pl" -DomainMail "gabos.pl" -Description "$Department/$Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "WINGABOS" -DomainMail "gabos.pl"
    Set-CredentialMxCoig -Target "mxgabos"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "WINGABOS" -DomainMail "gabos.pl"
    Get-SendEmailToManager

        }
    elseif ($Company -eq "7")
        {
        C:\Scripts\addnewuser_logicsynergy.ps1
        }
    elseif ($Company -eq "8")
        {
    Clear-Host
    Write-Output ""
    Write-Output "1. FK  >>>  Gliwice - Berbeckiego"
    Write-Output "----------------------------------------"
    Write-Output "2. FK  >>>  Katowice - Mikołowska"
    Write-Output "----------------------------------------"
    Write-Output ""
    Write-Output ""

    $Company=Read-Host "WYBIERZ ADRES"

    if ($Company -eq "1")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "FK GLIWICE BERBECKIEGO"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountFk
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Berbeckiego 6" -City "Wasko - Centrala Gliwice" -Company "WASKO4BUSINESS SP. Z O.O." -PostalCode "44-100" -Server "fk.wasko.pl" -DomainMail "wasko4b.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "FK" -DomainMail "wasko4b.pl"
    Set-CredentialMxCoig -Target "mxwasko4b"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "FK" -DomainMail "wasko4b.pl"
    Get-SendEmailToManager
        }
    elseif ($Company -eq "2")
        {
    Clear-Host
    Write-Host -ForegroundColor DarkGray  "FK KATOWICE MIKOŁOWSKA"
    $GLOBAL:Pass = Get-RandomPassword 15
    Set-CredentialAD -Target "wasko.pl" -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-OrganizationUnit -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-DeclaringVariables
    Get-DeclaringVariablesGroupAndAccountFk
    Set-ChangePolishSign
    Get-CheckInputData
    Get-ChoiseManagerInOrganizationUnit -Server "fk.wasko.pl" -pathOU "OU=W4B, OU=Wspolpracownicy, OU=WASKO, DC=fk, DC=wasko, DC=pl"
    Get-AddNewUser -StreetAddress "Mikołowska 100" -City "Katowice" -Company "WASKO4BUSINESS SP. Z O.O." -PostalCode "40065" -Server "fk.wasko.pl" -DomainMail "wasko4b.pl" -Department "$Department" -Description "$Department, $Firstname $Lastname"
    Get-AddUserToGroupMember
    Get-WriteToFileNewUsersList -Domain "FK" -DomainMail "wasko4b.pl"
    Set-CredentialMxCoig -Target "mxwasko4b"
    Get-AddNewAccountSmarterMail -enableMailForwarding "false"
    Get-CreateTempFileToSendViaMail -Domain "FK" -DomainMail "wasko4b.pl"
    Get-SendEmailToManager
        }
}