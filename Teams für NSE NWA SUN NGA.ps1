# Überprüfung und Installation fehlender Module
$requiredModules = @('MicrosoftTeams', 'PnP.PowerShell')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -Scope CurrentUser
    }
    Import-Module -Name $module
}

# Log-Datei-Pfad
$logPath = "C:\Users\tristan.boehling\OneDrive - Wittl-IT\Desktop\Powershell Scripte\Lofiles bei Skriptausführung"
if (-not (Test-Path -Path $logPath)) {
    New-Item -ItemType Directory -Path $logPath -Force
}
$logFile = Join-Path -Path $logPath -ChildPath "TeamCreationLog.txt"

function Log-Error {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp ERROR: $Message" | Out-File -FilePath $logFile -Append
}

function Log-Info {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp INFO: $Message" | Out-File -FilePath $logFile -Append
}

# Einleitung
Write-Host -ForegroundColor Green "Das ist ein Script, um ein MS Team automatisch zu erstellen. Bevor wir loslegen, melde dich mit dem Niesberger Admin-Account an."


try {
    # Stellt die Verbindung zum MS Konto her
    $tenant = Connect-MicrosoftTeams
    $account = $tenant.account
    $tenantdomain = $tenant.TenantDomain
    $host.ui.RawUI.WindowTitle = "Windows PowerShell ISE Remote Connection to Teams Tenant: $tenantdomain with Account $account"
    Write-Host -ForegroundColor Green "Verbindung erfolgreich hergestellt."
    Log-Info "Verbindung zu Microsoft Teams erfolgreich hergestellt."
} catch {
    Write-Host -ForegroundColor Red "Verbindung zu Microsoft Teams fehlgeschlagen. Bitte überprüfe deine Anmeldedaten und Internetverbindung."
    Log-Error "Verbindung zu Microsoft Teams fehlgeschlagen."
    exit
}

# Anleitung was zu tun ist
Write-Host -ForegroundColor Green "Nun brauche ich noch einige Parameter von dir."

$displayname = Read-Host "Name für das Team"
$MailNickName = Read-Host "Name für die Mail WICHTIG! Ohne Domäne eingeben!"
$number = Read-Host "Für welche Abteilung soll das Team erstellt werden? (1 for NWA, 2 for NWE, 3 for SUN, 4 for RME, 5 for SF-Bau, 6 for NGA, 7 for Donat)"

try {
    $group = New-Team -MailNickname $MailNickName -DisplayName $displayname -Visibility Private `
    -AllowGiphy $true -AllowStickersAndMemes $true -AllowCustomMemes $true `
    -AllowGuestCreateUpdateChannels $true -AllowGuestDeleteChannels $true `
    -AllowCreateUpdateChannels $true -AllowDeleteChannels $true -AllowAddRemoveApps $true `
    -AllowCreateUpdateRemoveTabs $true -AllowCreateUpdateRemoveConnectors $true `
    -AllowUserEditMessages $true -AllowUserDeleteMessages $true -AllowOwnerDeleteMessages $true `
    -AllowTeamMentions $true -AllowChannelMentions $true -ShowInTeamsSearchAndSuggestions $true

    Write-Host -ForegroundColor Cyan "Team '$displayname' wurde erfolgreich erstellt."
    Log-Info "Team '$displayname' erfolgreich erstellt."
} catch {
    Write-Host -ForegroundColor Red "Fehler beim Erstellen des Teams. Bitte überprüfe die eingegebenen Daten und versuche es erneut."
    Log-Error "Fehler beim Erstellen des Teams '$displayname'."
    exit
}


# Funktion zur Kanal- und Benutzerzuweisung mit Fehlerbehandlung und Fortschrittsanzeige
function Add-TeamDetails {
    param (
        [string]$GroupId,
        [array]$Channels,
        [hashtable]$Users,
        [string]$ImagePath
    )

    $totalOperations = $Channels.Count + $Users.Count + ($ImagePath -ne '')
    $completedOperations = 0

    foreach ($channel in $Channels) {
        try {
            New-TeamChannel -GroupId $GroupId -DisplayName $channel
            Log-Info "Kanal '$channel' erfolgreich hinzugefügt."
            $completedOperations++
            Write-Progress -Activity "Füge Kanäle hinzu" -Status "$completedOperations von $totalOperations abgeschlossen" -PercentComplete (($completedOperations / $totalOperations) * 100)
        } catch {
            Log-Error "Fehler beim Hinzufügen des Kanals '$channel'."
        }
    }

    foreach ($user in $Users.GetEnumerator()) {
        try {
            Add-TeamUser -GroupId $GroupId -User $user.Name -Role $user.Value
            Log-Info "Benutzer '$($user.Name)' erfolgreich hinzugefügt als $($user.Value)."
            $completedOperations++
            Write-Progress -Activity "Füge Benutzer hinzu" -Status "$completedOperations von $totalOperations abgeschlossen" -PercentComplete (($completedOperations / $totalOperations) * 100)
        } catch {
            Log-Error "Fehler beim Hinzufügen des Benutzers '$($user.Name)'."
        }
    }

    if ($ImagePath -ne '') {
        try {
            Set-TeamPicture -GroupId $GroupId -ImagePath $ImagePath
            Log-Info "Team-Bild erfolgreich gesetzt."
            $completedOperations++
            Write-Progress -Activity "Setze Team-Bild" -Status "$completedOperations von $totalOperations abgeschlossen" -PercentComplete 100
        } catch {
            Log-Error "Fehler beim Setzen des Team-Bildes."
        }
    }

    Write-Progress -Activity "Team Details hinzufügen" -Completed
}

# SharePoint-Integration mit Fehlerbehandlung
function Add-SharePointDetails {
    param (
        [string]$SiteUrl
    )

    try {
        Connect-PnPOnline -Url $SiteUrl -UseWebLogin
        Log-Info "Verbindung zu SharePoint Online erfolgreich hergestellt."
    } catch {
        Log-Error "Verbindung zu SharePoint Online fehlgeschlagen."
        return
    }

    try {
        Add-PnPField -List "Dokumente" -DisplayName "Aktenplan" -InternalName "Aktenplan" `
        -Type Choice -Group "Benutzerdefinierte Spalten" -Choices @(
            "AG-Kalkulation",
            "AG-Angebot",
            "AG-Vertrag",
            "AG-Nachtrag",
            "AG-Abnahme",
            "AN-Angebot",
            "AN-Vertrag",
            "AN-Nachtrag",
            "AN-Abnahme",
            "Schriftverkehr",
            "Technische Unterlagen",
            "Kaufmännische Unterlagen",
            "Bescheinigungen",
            "AG-Mängel",
            "AN-Mängel",
            "Bürgschaften/Avale",
            "Aktennotizen + Protokolle",
            "Dokumentation",
            "Immobilie",
            "Fuhrpark"
        ) -AddToDefaultView
        Log-Info "Felder zu SharePoint-Dokumentenliste erfolgreich hinzugefügt."

    } catch {
        Log-Error "Fehler beim Hinzufügen von Feldern zu SharePoint-Dokumentenliste."
    }
}

# Auswahl der Abteilung und Durchführung der spezifischen Konfiguration

switch ($number) {
    "1" {
    # Channels beibehalten
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("01 Auftraggeber", "02 Nachunternehmer", "03 Schriftverkehr - Sonstiges", "04 Aktennotizen - Protokolle", "05 Kaufmännische Berichte", "06 Technische Unterlagen", "07 Planung", "08 Terminplan", "09 Fotos", "10 Abnahmen", "11 Bestandsunterlagen", "Archivierung") `
    -Users @{
        "rene.fabian@niersberger.de"="Member";
        "andreas.bartl@niersberger.de"="Member";
        "joeran.brenn@niersberger.de"="Member";
        "pal.boelcskei@niersberger.de"="Member";
        "enrico.canonico@niersberger.de"="Member";
        "alexandra.gawenda@niersberger.de"="Member";
        "doris.giese@niersberger.de"="Owner";
        "karolina.heil@niersberger.de"="Member";
        "markus.hirt@niersberger.de"="Member";
        "martin.kalinowski@niersberger.de"="Member";
        "glikeria.kranvogel@niersberger.de"="Owner";
        "dietmar.leichsenring@niersberger.de"="";
        "sandra.liedtke@niersberger.de"="Member";
        "christian.pichl@niersberger.de"="Member";
        "stefanie.piller@niersberger.de"="Member";
        "eduard.reimche@niersberger.de"="Member";
        "christof.reiner@niersberger.de"="Member";
        "ralph.scherber@niersberger.de"="Member";
        "stefan.schleicher@niersberger.de"="Member";
        "martin.scholz@niersberger.de"="Owner";
        "christopher.sheffield@niersberger.de"="Member";
        "torsten.ulbrich@niersberger.de"="Member";
        "hans-peter.ulrich@niersberger.de"="Member";
        "hans-guenther.wachter@niersberger.de"="Member";
        "natanel.wuensch@niersberger.de"="Member";
        "torsten.wybranitz@niersberger.de"="Owner";
        "hueseyin.yildiz@niersberger.de"="Member";
        "hande.yilmaz@niersberger.de"="Member";
        "hendrik.zink@niersberger.de"="Member"
    } `
    -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\Niersberger-N.jpg"
}
"2" {
    # Channels beibehalten
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("01 Angebote und Verträge", "02 Bauunterlagen", "03 Pläne", "04 Kostenschätzung -berechnung -verfolgung", "05 Schriftverkehr - Protokolle - Fotos - Sonstiges", "06 Berechnungen", "07 Herstellerunterlagen", "08 Leistungsverzeichnisse", "09 Objektüberwachung", "10 Abgabe", "11 CAD", "Archivierung") `
    -Users @{
        "alessa.hahn@nwe.gmbh"="Owner";
        "daniel.wich@nwe.gmbh"="Owner";
        "bernhard.schierl@nwe.gmbh"="Member";
        "david.zhu@nwe.gmbh"="Member";
        "dirk.bruckhaus@nwe.gmbh"="Member";
        "igor.pomozov@nwe.gmbh"="Member";
        "annika.schimon@niersberger.de"="Member";
        "jonathan.griener@nwe.gmbh"="Member";
        "iris.wehner@nwe.gmbh"="Member";
        "linh.nguyen@nwe.gmbh"="Member";
	"luis.frankenhauser@nwe.gmbh"="Member";
        "marco.kuehn@nwe.gmbh"="Member";
        "marion.kuehn@nwe.gmbh"="Member";
        "michael.spangler@nwe.gmbh"="Member";
        "oliver.reichel@nwe.gmbh"="Member";
        "phil.spindler@nwe.gmbh"="Member";
	"sebastian.kammerer@nwe.gmbh"="Member";
        "tatjana.korobkina@nwe.gmbh"="Member";
        "tobias.hirt@nwe.gmbh"="Member";
        "victor.mironenko@nwe.gmbh"="Member";
        "xiaoyan.lin@niersberger.de"="Member";
        "frank.baessler@hnb-energie.de"="Member";
        "yulia.kuznetsova@nwe.gmbh"="Member";
        "cesar.herranz@nwe.gmbh"="Member";
        "ines.aldaz@nwe.gmbh"="Member";
        "jose.andreu@nwe.gmbh"="Member";
	"jose.collado@nwe.gmbh"="Member";
        "thomas.ennich@nwe.gmbh"="Owner"
    } `
    -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\NWE 84x84.jpg"
}

"3" {
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("Archivierung") `
    -Users @{
        "mesut.sehir@sunplanung.de"="Owner";
        "riikka.spaeth@sunplanung.de"="Member";
        "annika.schimon@niersberger.de"="Member";
        "frank.bildhauer@niersberger.de"="Member";
        "ayman.marshan@sunplanung.de"="Member"
    } `
    -ImagePath "" # Pfad zum Bild, falls vorhanden, oder leer lassen
}

"4" {
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("Archivierung") `
    -Users @{
        "maximilian.hass@niersberger.de"="Owner";
        "eric.posse@rme.eu"="Member";
        "christian.putsche@rme.eu"="Member";
        "michael.gumpert@rme.eu"="Member";
        "thomas.brauer@rme.eu"="Member";
        "anna-maria.reiber@rme.eu"="Member";
        "paul.zoller@rme.eu"="Member";
        "lisa.michaelis@rme.eu"="Member";
        "claudia.lunderstedt-georgi@rme.eu"="Member";
        "sabrina.peters@rme.eu"="Member";
        "andrea.lohmann@rme.eu"="Member"
    } `
    -ImagePath "C:\Users\tristan.boehling\OneDrive - Wittl-IT\Dokumente\Kunden\Niersberger-Dokumente\Projekte-Teamserstellung\RME-Logo-RGB.png"
}


    "5" {
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("01 Planung", "02 Bau", "03 Inbetriebnahme", "04 Wartung", "05 Administration", "Archivierung") `
    -Users @{
        "frank.bildhauer@niersberger.de"="Owner"
    } `
    -ImagePath "" # Pfad zum Bild, falls vorhanden, oder leer lassen
}

"6" {
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("Archivierung") `
    -Users @{
        "maximilian.hass@niersberger.de"="Member";
        "christopher.sheffield@niersberger.de"="Member";
        "rainer.schneiderbanger@niersberger.de"="Member";
        "alessandro.eckert@niersberger.de"="Member";
        "lorena.kqira@niersberger.de"="Member";
        "manuela.mauser@niersberger.de"="Member";
       
    } `
    
}
"7" {
    # Channels beibehalten
    Add-TeamDetails -GroupId $group.GroupId `
    -Channels @("Archivierung") `
    -Users @{
        "maximilian.hass@niersberger.de"="Owner";
        "enrico.donat@donat-haustechnik.de"="Member";
        "christian.putsche@rme.eu"="Member";
        "luisa.koppe@donat-haustechnik.de"="Member";
        "thomas.brauer@rme.eu"="Member";
        "sebastian.schmidt@donat-haustechnik.de"="Member"

    } `
    -ImagePath "C:\Users\tristan.boehling\OneDrive - Wittl-IT\Dokumente\Kunden\Niersberger-Dokumente\Projekte-Teamserstellung\Donat-klein.png"

}


}


# Aufruf der Funktion für SharePoint-Details mit der generierten Site URL
Add-SharePointDetails -SiteUrl $siteUrl
Log-Info "SharePoint-Details Funktion wurde aufgerufen."

Write-Host "Fertig. Drücke Enter, um das Skript zu beenden..."
Read-Host
Disconnect-PSSession
