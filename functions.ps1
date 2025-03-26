function HelloScriptDatei {
    <# 
        .SYNOPSIS
        Sagt Hallo, wenn die Datei korrekt eingebunden wurde.

        .DESCRIPTION 
        Sagt Hallo, wenn die Datei korrekt eingebunden wurde.
    #>
    Write-Host "Scriptdatei eingebunden"
}


function checkKlassenteam {
     <# 
        .SYNOPSIS
       Prüft, ob der Schüler im Klassenteam seiner Klasse sitzt und entfernt ihn aus einer alten Klasse, fals er verschoben wurde

        .DESCRIPTION 
         Prüft, ob der Schüler im Klassenteam seiner Klasse sitzt und entfernt ihn aus einer alten Klasse, fals er verschoben wurde

        .EXAMPLE
        checkKlassenteamByID "Lampe"
    #>
    param(
    [string]$Name,
    [string]$Klasse,
    $Teams
    )
    Write-Host "Prüft die falsche Klassengruppenzugehörigkeit für Name $Name"



    foreach ($class in $Teams) {
    

    #eigenes Klasssenteam muss nicht beachtet werden.
        if($class.Klasse -eq $Klasse){
             continue
        }
   
        foreach ($member in $class.Mitglieder) {
            #Wenn die Mailadresse in den Membern gefunden wird, löschen.
            if($member.UserPrincipalName -eq $Name){
              Write-Host "Benutzer $Name ist Mitglied des Teams $($class.Klasse)"
                # Benutzer aus dem Team entfernen
                try{
                Remove-TeamUser -GroupId $class.GroupId -User $Name
                Write-Host "Benutzer $Name wurde aus dem Team $($class.Klasse) entfernt."
                } catch {
                 Write-Host ("Fehler beim löschen von $Name aus Team  $($class.Klasse)")
    
                }
            }
             
         
        }
    
    #Write-Host "-----------------------------------"
    }
}


function Get-TeamsByPattern {
    param(
       [string]$FilterPattern
        
    )
    
    # Erstellen einer Struktur, um die gefilterten Teams und ihre Mitglieder zu speichern
    $teamsStructure = @()

    # Alle Teams abrufen und filtern
    $filteredTeams = Get-Team | Where-Object { $_.DisplayName -like $FilterPattern }
    
    
    foreach ($team in $filteredTeams) {
   
          
        # Mitglieder des aktuellen Teams abrufen
        $teamUsers = Get-TeamUser -GroupId $team.GroupId -Role Member

        # Struktur für das Team erstellen
        $teamInfo = [PSCustomObject]@{
            Klasse = $team.DisplayName
            GroupId = $team.GroupId
            Mitglieder = $teamUsers | ForEach-Object {
                [PSCustomObject]@{
                    UserPrincipalName = $_.User
                }
            }
        }

        # Teaminfo zur Struktur hinzufügen
        $teamsStructure += $teamInfo
    }

    # Rückgabe der Struktur
    return $teamsStructure
}





function checkKlassenteamParallel {
    <# 
        .SYNOPSIS
       Prüft, ob der Schüler im Klassenteam seiner Klasse sitzt und entfernt ihn aus einer alten Klasse, falls er verschoben wurde.

        .DESCRIPTION 
         Prüft, ob der Schüler im Klassenteam seiner Klasse sitzt und entfernt ihn aus einer alten Klasse, falls er verschoben wurde.

        .EXAMPLE
        checkKlassenteam -Name "Lampe" -Klasse "10A" -Teams $filteredTeams
    #>
    param(
        [string]$Name,
        [string]$Klasse,
        [array]$Teams
    )

    # In PowerShell 7.x verfügbar, um parallele Verarbeitung zu nutzen
    $Teams | ForEach-Object -Parallel {
        param (
            $team,
            $Name,
            $Klasse
        )

        Write-Host "Prüft die Klassengruppenzugehörigkeit für Name $Name und Klasse $Klasse"

        # Eigenes Klassenteam muss nicht beachtet werden.
        if ($team.DisplayName -eq $Klasse) {
            return
        }

        # Mitglieder des aktuellen Teams abrufen
        $teamMembers = Get-TeamUser -GroupId $team.GroupId

        # Überprüfen, ob der Benutzer Mitglied des Teams ist
        $isMember = $teamMembers | Where-Object { $_.UserPrincipalName -eq $Name }

        if ($isMember) {
            Write-Host "Benutzer $Name ist Mitglied des Teams $($team.DisplayName)"
            
            # Benutzer aus dem Team entfernen
            Remove-TeamUser -GroupId $team.GroupId -User $Name

            Write-Host "Benutzer $Name wurde aus dem Team $($team.DisplayName) entfernt."
        }
    } -ArgumentList $_, $Name, $Klasse -ThrottleLimit 10
}

