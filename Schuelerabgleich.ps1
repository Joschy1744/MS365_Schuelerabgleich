# Installiere die erforderlichen Module, wenn nicht bereits installiert
if (-not (Get-Module -Name MicrosoftTeams -ListAvailable)) {
    Install-Module -Name MicrosoftTeams -Force -AllowClobber
}

if (-not (Get-Module -Name AzureAD -ListAvailable)) {
    Install-Module -Name AzureAD -Force -AllowClobber
}


if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module -Name Microsoft.Graph -Force -AllowClobber
}


# Attribute
# https://learn.microsoft.com/en-us/microsoft-365/enterprise/configure-user-account-properties-with-microsoft-365-powershell?view=o365-worldwide
Import-Module -Name ImportExcel
Import-Module -Name Communary.PASM
Import-Module Microsoft.Graph.Users


##################################################################################################################################################################
## Löschen der Nutzer, die nicht mehr in eingelesener Datei sind ACHTUNG: immer erst Dry Run mit False machen und Exportdateien kontrollieren!
##
$deleteOldUsers = $false
##
##
## Anlegen der Nutzer, die in eingelesener Datei sind und noch nicht in AD ACHTUNG: immer erst Dry Run mit False machen und Exportdateien kontrollieren!
##
$generateNewUsers = $false
##
## Funktioniert nur, wenn auch wirklich ein Nutzer angelegt wird, es werden alle Lehrer angeschrieben, die den Klassennamen in -Department haben
## Sendet jeden Zugang einzeln, gut für unter dem Jahr
$sendNewCredentialsToKL = $true
##
## Sendet die Zugangsdaten komprimiert nach Klasse, gut für Jahresanfang
$sendNewCredentialsToKLKomprimiert = $false
##
## Schüler zur Klassengruppe des Klassenlehrers hinzufügen.
$gruppenhinzufuegen = $true
##
## Schüler aus anderen Klassengruppen im gleichen Jahr entfernen, geht nur in Kombination mit $gruppenhinzufuegen
$susKlassengruppeVersetzen = $true
##
## Für die Jahresgruppen Klassen: Zuerst Lehrer durchlaufen lassen, dort wird das Klassenteam erstellt
## Schuljahresstart
$jahr = "2024"
##
##
## Schülerdomain
##
$TargetUsername = "@h.de"
##
##  Mailversand der neuen Zugangsdaten an folgende Domain (Abgeleich der Klasse über -Department
$LehrerUsername = "@hm.de"
##
##
## hinzufügen zu folgender Gruppe
##
$GroupName1 = "#HMS"
$GroupID1 = "d788d941-1f85-4ecf-97a6-ea5e1f2"
##
## Mail-Versand der Zugangsdaten in BCC an:
$MailBCC = "it@hm.de"
## Mail-Versand der Zugangsdaten in BCC an:
$MailBCC2 = "a@hm.de"
##
## Mail des Abensender, sollte identisch mit dem Login sein
$MailAbsender = "office365@onmicrosoft.com"
##
##
$lizenzSuS = 'ENTERPRISEPACKPLUS_STUUSEBNFT'
##
##
##################################################################################################################################################################


# Pfad zur Excel-Datei
# Den Ordnerpfad, in dem das Skript ausgeführt wird, abrufen
$ordnerPfad = $PSScriptRoot
$logFilePath = "$PSScriptRoot\logfile.txt"

# Öffnen oder erstellen Sie die Logdatei und leiten Sie die Ausgabe dorthin um
Start-Transcript -Path $logFilePath

#Einbinden der externen Methodendatei
. "$PSScriptRoot\functions.ps1"
HelloScriptDatei


# Verbindung mit Azure AD herstellen
Connect-AzureAD
Write-Host ("AzureAD verbunden") -ForegroundColor Green

# Verbindung mit Microsoft Teams herstellen
Connect-MicrosoftTeams
Write-Host ("MicrosoftTeam verbunden") -ForegroundColor Green

# Verbinden mit Microsoft Graph
Connect-MgGraph -NoWelcome -Scopes  User.ReadWrite.All, Organization.Read.All, Directory.ReadWrite.All, Mail.Send, Mail.Send.Shared
Write-Host ("Graph verbunden") -ForegroundColor Green



# Ausgabe des Ordnerpfads
Write-Host "Das Skript wird im Ordner ausgeführt: $ordnerPfad"


# Definieren Sie die Wildcard für den Dateinamen (z.B. #SPH*.txt)
$dateiWildcard = "tableExport*.txt"



# Suchen Sie nach Excel-Dateien im Ordner mit der Wildcard im Namen
$excelDatei = Get-ChildItem -Path $ordnerPfad -Filter $dateiWildcard | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# Überprüfen, ob eine passende Excel-Datei gefunden wurde
if ($excelDatei -ne $null) {
    # Jetzt können Sie die gefundene Excel-Datei einlesen
    $excelData = [System.Collections.ArrayList](Import-Csv -Path $excelDatei.FullName)
}
else {
    Write-Host "Keine passende Datei gefunden."
    Exit
}


# Alle Benutzer abrufen
# ArrayList für performatereres entfernen
$users = [System.Collections.ArrayList](Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -like "*$TargetUsername" } | Sort-Object -Property Surname)
 


$anzahlUserInAD = $users.Count
$anzahlUserInImport = $excelData.Count
# Array für fehlende Benutzer erstellen
$missingUsers = @()

$missingUsersinAD = @()

$usersAddToAD = @()

$usersDeletedFromAD = @()


$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

# check, ob eingelesene Nutzer in AD angelegt sind
$i = 1
foreach ($excelUser in $excelData) {

    Write-Host ("" + $i + " von " + $anzahlUserInImport + " Pruefen:  " + $excelUser.Nachname + ", " + $excelUser.Vorname + " Id: " + $excelUser.Id )
              
    $userFound = $false
    foreach ($user in $users) {
        [string]$GivenName = $user.GivenName | Out-String
        [string]$Vorname = $excelUser.Vorname | Out-String

        if ( ($user.ExtensionProperty.employeeId -eq $excelUser.Id) -or ($GivenName -eq $Vorname -and $user.Surname -eq $excelUser.Nachname) ) {
            $userFound = $true
            Write-Host ("" + $i + " von " + $anzahlUserInImport + " Gefunden: " + $excelUser.Nachname + ", " + $excelUser.Vorname + " Id: " + $user.ExtensionProperty.employeeId)
            break
        }
    }

    # Wenn der Benutzer nicht in der AD gefunden wurde, zur Liste der fehlenden Benutzer hinzufügen
    if (!$userFound) {
        $missingUserInAD = [PSCustomObject]@{
            Id           = $excelUser.Id
            Nachname     = $excelUser.Nachname
            Vorname      = $excelUser.Vorname
            Klassenstufe = $excelUser.Klasse
            Stufe        = $excelUser.Stufe
        }
        $missingUsersinAD += $missingUserInAD

        if ($generateNewUsers) {
            ## erstellen des neuen Nutzers, zu weisen der Lizenzen
            $dpname = $excelUser.Vorname + " " + $excelUser.Nachname

            if ($excelUser.Klasse) {
                $klasse = $excelUser.Klasse
            }
            else {
                $klasse = "keine"
            }
           
            $userPrincipalName = "" + $excelUser.Login + $TargetUsername
            $vorname = $excelUser.Vorname
            $nachname = $excelUser.Nachname
            
            $lizenzfuerSuS = Get-MgSubscribedSku -All | Where SkuPartNumber -eq $lizenzSuS
            $pw = "Hms" + $excelUser.Id
            $passwordProfile = @{
                forceChangePasswordNextSignIn = $true
                password                      = $pw
            }
           

            New-MgUser  -PasswordProfile $passwordProfile -GivenName $excelUser.Vorname -Surname $excelUser.Nachname -DisplayName $dpname  -Department $klasse -State $excelUser.Stufe -JobTitle "Schüler" -Country "Deutschland" -UserPrincipalName $userPrincipalName -AccountEnabled -EmployeeId $excelUser.Id -MailNickName $excelUser.Login -UsageLocation "DE"
            Set-MgUserLicense -UserId $userPrincipalName -AddLicenses @{SkuId = $lizenzfuerSuS.SkuId } -RemoveLicenses @()
            
            $userAddToAD = [PSCustomObject]@{
                Vorname      = $excelUser.Vorname
                Nachname     = $excelUser.Nachname
                Klasse       = $excelUser.Klasse
                Stufe        = $excelUser.Stufe
                Benutzername = $userPrincipalName
                Kennwort     = $pw
                Id           = $excelUser.Id
            }
            $usersAddToAD += $userAddToAD

            if ($sendNewCredentialsToKL) {
               
                $kls = Get-AzureADUser -Filter "startswith(Department,'$klasse')" -All:$true  | Where-Object { $_.UserPrincipalName -like "*$LehrerUsername" }  
                if ($kls) {
                 
                    foreach ($kl in $kls) {
                        $userID = $kl.UserPrincipalName
                        $vn = $kl.GivenName
                        $nn = $kl.Surname
                        $params = @{
                            Message         = @{
                                Subject       = "Neuer Nutzer in Klasse $klasse"
                                Body          = @{
                                    ContentType = "HTML"  
                                    Content     = "Hallo $vn $nn,<br>ein neuer Schüler wurde soeben in MS365/Teams deiner Klasse $klasse hinzugefügt.<br><br>-----------------------------<br><br><b>Name:</b> $vorname $nachname <br><b>E-Mail:</b> $userPrincipalName<br><b>Passwort:</b> $pw<br><br>-----------------------------<br><br>Bitte leite diese Daten an den/die Schüler/in weiter.<br><br><i>Das ist eine automatisch generierte E-Mail. Bitte antworte nicht auf diese E-Mail.<br>Fragen bitte an $MailBCC</i>"
                                }
                                ToRecipients  = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $userID
                                        }
                                    }
                                )
                                BccRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $MailBCC
                                        }
                                    }
                                )
                            }
                            SaveToSentItems = "false"
                        }
                        # A UPN can also be used as -UserId.
                    
                        Send-MgUserMail -UserId  $MailAbsender -BodyParameter $params
                        Write-Host ("Einzel-Mail für $klasse an $userID versendet.")
                    }
                }
            }
        }
    }
    else {
        #Entfernen des Datensatzes aus $Users um Suchlaufzeit zu verkürzen
        $users.Remove($user)
    }

    
    $i++
    Write-Host ("---------------------------------------------------------")
}


# Fehlende Benutzer aus LUSD in CSV-Datei schreiben
$missingUsersinADCsvPath = $ordnerPfad + "\nutzerInImportAberNichtInAD.csv"
$missingUsersinAD | Export-Csv -Path $missingUsersinADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

if ($generateNewUsers) {
    $usersAddToAD = $usersAddToAD | Sort-Object -Property Klasse

    $usersAddToADCsvPath = $ordnerPfad + "\neuErstellteNutzer_$timestamp.csv"
    $usersAddToAD | Export-Csv -Path $usersAddToADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

    if ($sendNewCredentialsToKLKomprimiert) {
             # Gruppiere die Benutzer nach Klasse
             $groupedUsers = $usersAddToAD | Group-Object -Property Klasse
             # Iteriere über jede Klasse
                foreach ($group in $groupedUsers) {
                    $klasse = $group.Name
                    $benutzer = $group.Group
    
                    # Erstellen des E-Mail-Inhalts
                    $emailBody = "Informationen für Klasse $klasse<br><br>"
                    $emailBody += "<table><tr>"

                    foreach ($user in $benutzer) {
                        $emailBody += "<tr>"
                        $emailBody += "<td><b>Vorname:</b> $($user.Vorname)</td>"
                        $emailBody += "<td><b>Nachname:</b> $($user.Nachname)</td>"
                        $emailBody += "<td><b>Benutzername:</b> $($user.Benutzername)</td>"
                        $emailBody += "<td><b>Kennwort:</b> $($user.Kennwort)</td>"
                        $emailBody += "</tr>"
                    }
                    $emailBody += "</table>"
                 



               $kls = Get-AzureADUser -Filter "startswith(Department,'$klasse')" -All:$true  | Where-Object { $_.UserPrincipalName -like "*$LehrerUsername" }  
                if ($kls) {
                     foreach ($kl in $kls) {
                        $userID = $kl.UserPrincipalName
                        $vn = $kl.GivenName
                        $nn = $kl.Surname
                        $params = @{
                            Message         = @{
                                Subject       = "Neue Nutzer in Klasse $klasse"
                                Body          = @{
                                    ContentType = "HTML"  
                                    Content     = "Hallo $vn $nn,<br>neue Schüler wurde soeben in MS365/Teams deiner Klasse $klasse hinzugefügt.<br><br>
                                    -----------------------------<br><br>
                                    $emailBody
                                    <br><br>-----------------------------<br><br>
                                    Bitte leite diese Daten an den/die Schüler/innen weiter.<br><br>
                                    <i>Das ist eine automatisch generierte E-Mail. Bitte antworte nicht auf diese E-Mail.<br>Fragen bitte an $MailBCC</i>"
                                }
                                ToRecipients  = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $userID
                                        }
                                    }
                                )
                                BccRecipients = @(
                                    @{
                                        EmailAddress = @{
                                            Address = $MailBCC
                                        }
                                    },
                                    @{
                                        EmailAddress = @{
                                            Address = $MailBCC2
                                        }
                                    }
                                )
                            }
                            SaveToSentItems = "false"
                        }
                        # A UPN can also be used as -UserId.
                    
                        Send-MgUserMail -UserId  $MailAbsender -BodyParameter $params
                        Write-Host ("Gruppen-Mail für $klasse an $userID versendet.")
                    }
                }
             }
    }

}

#### Neu Nutzer laden, da potentiell neu angelegt.

$users = Get-AzureADUser -All $true | Where-Object { $_.UserPrincipalName -like "*$TargetUsername" } | Sort-Object -Property Surname
$AllTeamsInOrg = Get-Team 

# Alle Teams in der Organisation abrufen und nach dem Schema filtern um zu überprüfen, ob ein Schüler falsch sitzt
$filterPattern = "* Klassengruppe $jahr"  # Das Schema, nach dem du filtern möchtest



# Erstellen einer Liste von Klassen und deren Mitgliedern
$classes = Get-TeamsByPattern -FilterPattern $filterPattern


$i = 0
# Schleife durch alle Benutzer
foreach ($user in $users) {
    $i++

    $userFound = $false
     
    Write-Host ("" + $i + " von " + $anzahlUserInAD )


    # Schleife durch die Excel-Daten
    foreach ($excelUser in $excelData) {
   
        [string]$GivenName = $user.GivenName | Out-String
        [string]$Vorname = $excelUser.Vorname | Out-String
         
        # Überprüfen, ob Vorname und Nachname übereinstimmen oder IDs gleich.
          
        if (($user.ExtensionProperty.employeeId -eq $excelUser.Id) -or ($user.GivenName -eq $excelUser.Vorname -and $user.Surname -eq $excelUser.Nachname)) { 
            $userFound = $true
            if ($excelUser.Klasse) {
                $klasse = $excelUser.Klasse
            }
            else {
                $klasse = "keine"
            }
            # Department auf den Wert der Klasse setzen
            $dpname = $excelUser.Vorname + " " + $excelUser.Nachname
            $extensionProps = New-Object System.Collections.Generic.Dictionary"[String,String]"
            $extensionProps.Add("employeeId", $excelUser.Id)
          

            Set-AzureADUser -ObjectId $user.ObjectId -GivenName $excelUser.Vorname -Surname $excelUser.Nachname -DisplayName $dpname  -Department $klasse -State $excelUser.Stufe -Country "Deutschland" -ExtensionProperty $extensionProps -JobTitle "Schüler"
           
            Write-Host ("Eigenschaften für Benutzer " + $user.UserPrincipalName + " gesetzt.")


            # hinzufügen zu der SCHULE-Gruppe

            # Benutzer zum Teams-Team hinzufügen
            try {
                $membershipType = "Member"
                        
                Add-TeamUser -GroupId $GroupID1 -User $user.UserPrincipalName -Role $membershipType

                Write-Host ("Benutzer " + $user.UserPrincipalName + " wurde als " + $membershipType + " zum Team $GroupName1 hinzugefügt.")
            }
            catch {
                Write-Host ("Fehler beim Hinzufügen von Benutzer " + $user.UserPrincipalName + " zum Team " + $_.Exception.Message)
            }

            break
        }
    }
    
    # Wenn der Benutzer nicht in der Excel-Datei gefunden wurde, zur Liste der fehlenden Benutzer in Ursprungsdatei hinzufügen
    if (!$userFound -and $user.JobTitle -eq "Schüler") {
        $missingUser = [PSCustomObject]@{
            Id                = $user.ExtensionProperty.employeeId
            Klasse            = $user.Department
            Stufe             = $user.State
            Name              = $user.DisplayName
            Nachname          = $user.Surname
            Vorname           = $user.GivenName
            ObjectId          = $user.ObjectId
            UserPrincipalName = $user.UserPrincipalName
        }
        $missingUsers += $missingUser

        #Wenn Löschen $true, dann löschen
        if ($deleteOldUsers -and $user.JobTitle -eq "Schüler") {
            Remove-AzureADUser -ObjectId $user.ObjectId
            $usersDeletedFromAD += $missingUser
        }
    }
    elseif ($userFound -and $user.JobTitle -eq "Schüler") {
        $excelData.Remove($excelUser)
    } 

    ######################################################
    ## Nutzer zum Klassenteam hinzufügen
    ######################################################
    # Wenn Klasse verfügbar
    if ($user.Department -and $gruppenhinzufuegen) {

        $klasse = $user.Department
            
        $classTeamName = $klasse + " Klassengruppe " + $jahr
        $classTeam = $null
            
        foreach ($team in $AllTeamsInOrg) {
            if ($team.DisplayName -eq $classTeamName) {
                $classTeam = $team
                break
            }
        }            
           
            
        # Überprüfen, ob ein Microsoft Teams-Team mit dem Namen der Klasse existiert
                   
            
        if ($classTeam) {
                           
            # Benutzer zum Teams-Team hinzufügen
            try {
                $membershipType = "Member"
                Add-TeamUser -GroupId $classTeam.GroupId -User $user.UserPrincipalName -Role $membershipType
                        
                Write-Host ("Benutzer " + $user.UserPrincipalName + " wurde als " + $membershipType + " zum Team " + $classTeam.DisplayName + " hinzugefügt.")
            }
            catch {
                Write-Host ("Fehler beim Hinzufügen von Benutzer " + $user.UserPrincipalName + " zum Team " + $classTeam.DisplayName + ": " + $_.Exception.Message)
            }

        }

        # Prüft, ob der Nutzer in einem Klassenteam sitzt, in das er nicht gehört und entfernt ihn daraus.
        if($susKlassengruppeVersetzen){
            checkKlassenteam -Name $user.UserPrincipalName -Klasse $classTeamName -Teams $classes
         }  
    }
    Write-Host ("---------------------------------------------------------")
}

# Fehlende Benutzer in CSV-Datei schreiben
$missingUsersCsvPath = $ordnerPfad + "\nutzerInAdAberNichtInImport.csv"
$missingUsers | Export-Csv -Path $missingUsersCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8

if ($deleteOldUsers) {
    $usersDeletedFromADCsvPath = $ordnerPfad + "\geloeschteNutzer_$timestamp.csv"
    $usersDeletedFromAD | Export-Csv -Path $usersDeletedFromADCsvPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
}


# Verbindung zu Azure AD trennen
Disconnect-AzureAD

# Beenden der Transkription und Schließen der Logdatei
Stop-Transcript