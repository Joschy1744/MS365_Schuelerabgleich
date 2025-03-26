# MS365_Schuelerabgleich
 Ein Powershell Schülerabgleich mit Daten aus der BNV des SPH Hessen. 

Datenbasis: 
	SPH Benutzerverwaltung Lernende -> Deaktivierte unsichtbar, sonst sind alte accounts noch mit aufgeführt, ggf langen dann die Lizenzen nicht.
	Auswahl Zeige "Alle" Einträge ganz unten links auf seite, dann oben auf Tabelle rechts Export (Balken mit Pfeil nach oben rechts) und als !!!! TXT exportieren, weil csv BOM-Order Zeichen im Export hat, was Powershell nicht so mag.
	
	Die Datei tableExport.txt in den Ordner des Scripts legen, diese wird automatisch gesucht und gefunden.
	
	Öffnen der Datei Schuelerabgleich.ps1 und eisntellen der Parameter $deleteOldUsers, $generateNewUsers, $jahr
	
	wenn true, dann wird gelöscht und angelegt, wenn false, wird nur in eien Datei ausgegeben, welches resultat der Abgleich hat. 
	
	$jahr gibt das Schuljahr an und setzt Schüler in Klassenteams des Jahres. z.b. 05ah Klassenteam 2023
	
	
	Account benötigt Adminrechte und auf dem PC muss powershell mit MicrosoftTeams, AzureAd und MS.Graph installiet sein, das Script sollte das prüfen und nachinstallieren. 
	
	Es empfielt sich die Nutzung von Windows PowerShell ISE
	
	
	
	Beim Anlegen wir der Nutzer mit eine A3 Lizenz Schüler versehen, und in die entsprechenden Gruppen gelegt, am Ende kommt eine Export csv in den Ordner, mit den Zugangsdaten.
	
	Hinweis für Jahreübergang, vermutlich empfielt es sich, erst ein dryrun mit beiden parametern auf false, dann erst löschen um die Lizenzen frei zu machen und im letzen Schritt die Nutzer zu erstellen.