Clear


#Einleitung
Write-Host -ForegroundColor Green "Das ist ein Script um ein MS Team automatisches zu erstellen.
Bevor wir loslegen melde dich mit einem Admin Account bei der Frima an, wo du das Team erstellen willst."


#Warten zum Lesen 
Start-Sleep -s 5



#Stellt die Verbindung zum MS Konto her
$tenant = (Connect-MicrosoftTeams)
$account = ($tenant.account)
$tenantdomain = ($tenant.TenantDomain)
$host.ui.RawUI.WindowTitle = "Windows Powershell ISE remote connection to Teams Tenant: $tenantdomain with account $account "


#Andere Farbe der Schrift mit Anleitung was zu tun ist
Write-Host -ForegroundColor Green "Nun brauche ich noch einige Parameter von dir"



$displayname = Read-Host "Name für das Team" 

$description = $displayname

$MailNickName = Read-Host "Name für die Mail WICHTIG! Ohne Domäne eingeben!"


   $number = Read-Host "Für Welche Abteilung soll das Team erstellt werden?
1 for NWA
2 for NWE
3 for HNB   
4 for SUN
5 for NGA
6 for RME  
7 for SF-Bau      "











$boolean = "True"

$group = New-Team -MailNickname $MailNickName -displayname $displayname -Visibility private -AllowGiphy $true -AllowStickersAndMemes $true -AllowCustomMemes $true -AllowGuestCreateUpdateChannels $true -AllowGuestDeleteChannels $true -AllowCreateUpdateChannels $true -AllowDeleteChannels $true -AllowAddRemoveApps $true -AllowCreateUpdateRemoveTabs $true -AllowCreateUpdateRemoveConnectors $true -AllowUserEditMessages $true -AllowUserDeleteMessages $true -AllowOwnerDeleteMessages $true -AllowTeamMentions $true -AllowChannelMentions $true -ShowInTeamsSearchAndSuggestions $true





Foreach($number in $number)

{




   if($number.contains("1"))



		{



    Start-Sleep -s 40        


   New-TeamChannel -GroupId $group.GroupId -DisplayName "01 Auftraggeber"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "02 Nachunternehmer"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "03 Schriftverkehr - Sonstiges"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "04 Aktennotizen - Protokolle"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "05 Kaufmännische Berichte"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "06 Technische Unterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "07 Planung"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "08 Terminplan"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "09 Fotos"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "10 Abnahmen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "11 Bestandsunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "Archivierung"


   Add-TeamUser -GroupId $group.GroupId  -User fabian.rene@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User andreas.bartl@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User joeran.brenn@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User enrico.canonico@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User alexandra.gawenda@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User doris.giese@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User karolina.heil@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User martin.kalinowski@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User lorena.Kqira@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -groupid $group.groupid  -User glikeria.kranvogel@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User sandra.liedtke@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User julia.niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User christian.pichl@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User stefanie.piller@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User eduard.reimche@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User christof.reiner@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User ralph.scherber@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User stefan.schleicher@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User martin.scholz@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User silvia.schubbert@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User christopher.sheffield@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User torsten.ulbrich@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User hans-peter.ulrich@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User hans-guenther.wachter@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User natanel.wuensch@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User torsten.wybranitz@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User hueseyin.yildiz@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -groupid $group.groupid  -User hande.yilmaz@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User hendrik.zink@niersberger.de
   

   


   
   Start-Sleep -s 12
           
   Set-TeamPicture -GroupId $group.GroupId -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\Niersberger-N.jpg"

   

		}	



   if($number.contains("2"))
		{
		

   Start-Sleep -s 40

   New-TeamChannel -GroupId $group.GroupId -DisplayName "01 Angebote und Verträge"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "02 Bauunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "03 Pläne"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "04 Kostenschätzung -berechnung -verfolgung"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "05 Schriftverkehr - Protokolle - Fotos - Sonstiges"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "06 Berechnungen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "07 Herstellerunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "08 Leistungsverzeichnisse"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "09 Objektüberwachung"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "10 Abgabe"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "11 CAD"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "Archivierung"
          

   Start-Sleep -s 12
   
   Set-TeamPicture -GroupId $group.GroupId -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\NSE 84x84.jpg"
            
   
   Add-TeamUser -GroupId $group.GroupId  -User daniel.wich@nwe.gmbh -Role Owner
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User alessa.hahn@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User felix.schoenberger@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User tobias.hirt@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User michael.kaleta@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User marco.kuehn@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User xiaoyan.lin@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User torsten.wybranitz@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User annika.schimon@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -groupid $group.groupid  -User rene.fabian@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User oliver.reichel@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User igor.pomozov@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User frank.baessler@hnb-energie.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User sebastian.kammerer@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User alessandro.eckert@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User yulia.kuznetsova@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User david.zhu@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User iris.wehner@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User linh.nguyen@nwe.gmbh
   Start-Sleep -s 2



   Add-TeamUser -GroupId $group.GroupId  -User thomas.ennich@nwe.gmbh -Role Owner


              
              

		}	


   if($number.contains("3"))
		{


  
  Start-Sleep -s 40

   New-TeamChannel -GroupId $group.GroupId -DisplayName "01 Auftrag"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "02 Bauunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "03 Planeingänge"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "04 Elektroplanungen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "05 Versorgungsträger"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "06 Technikberechnungen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "07 Herstellerunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "08 Protokolle"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "09 Schriftverkehr"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "10 Kostenermittlung"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "11 Vergabeunterlagen"
   Start-Sleep -s 2
   New-TeamChannel -GroupId $group.GroupId -DisplayName "12 Bauüberwachung"
   Start-Sleep -s 2

   Start-Sleep -s 12

   Set-TeamPicture -GroupId $group.GroupId -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\HNB.jpg"
            

   Add-TeamUser -GroupId $group.GroupId  -User daniel.wich@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User alessa.hahn@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User felix.schoenberger@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User tobias.hirt@hnb-energie.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User sebastian.kammerer@hnb-energie.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User mesut.sehir@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User thomas.ennich@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User xiaoyan.lin@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User igor.pomozov@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User rene.fabian@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User annika.schimon@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User oliver.reichel@nwe.gmbh
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User michael.dotterweich@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User marco.kuehn@nwe.gmbh

    Add-TeamUser -GroupId $group.GroupId  -User michael.kaleta@nwe.gmbh

   Add-TeamUser -GroupId $group.GroupId  -User frank.baessler@hnb-energie.de -Role Owner

		}	



if($number.contains("4"))
		{
		

   New-TeamChannel -GroupId $group.GroupId -DisplayName "Archivierung"  
   
   Add-TeamUser -GroupId $group.GroupId  -User mesut.sehir@sunplanung.de -Role Owner
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User riikka.spaeth@sunplanung.de
   Start-Sleep -s 2
   Add-TeamUser -GroupId $group.GroupId  -User annika.schimon@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -groupid $group.groupid  -User frank.bildhauer@niersberger.de
   Start-Sleep -s 2
   Add-TeamUser -groupid $group.groupid  -User ayman.marshan@sunplanung.de
 

 

              
              

		}	


if($number.contains("5"))
		{
		
          
   
   Set-TeamPicture -GroupId $group.GroupId -ImagePath "$env:HOMEPATH\Wittl-IT\Wittl-IT Kunden - Dokumente\Niersberger\Projekt Teams V2\NSE 84x84.jpg"
            
   



}

if($number.contains("6"))
		{
		
          
   
#New-TeamChannel -GroupId $group.GroupId -DisplayName "Archivierung"

#$displaynameEXT = $displayname + "EXT"

#$groupEXT = New-Team -MailNickname $displaynameEXT  -displayname $displaynameEXT -Visibility private -AllowGiphy $true -AllowStickersAndMemes $true -AllowCustomMemes $true -AllowGuestCreateUpdateChannels $true -AllowGuestDeleteChannels $true -AllowCreateUpdateChannels $true -AllowDeleteChannels $true -AllowAddRemoveApps $true -AllowCreateUpdateRemoveTabs $true -AllowCreateUpdateRemoveConnectors $true -AllowUserEditMessages $true -AllowUserDeleteMessages $true -AllowOwnerDeleteMessages $true -AllowTeamMentions $true -AllowChannelMentions $true -ShowInTeamsSearchAndSuggestions $true
            
   
       	Add-TeamUser -GroupId $group.GroupId  -User maximilian.hass@niersberger.de -Role Owner

	


   	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User eric.posse@rme.eu
 	Add-TeamUser -GroupId $group.GroupId  -User christian.putsche@rme.eu
  	Add-TeamUser -GroupId $group.GroupId  -User michael.gumpert@rme.eu
  	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User thomas.brauer@rme.eu
  	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User anna-maria.reiber@rme.eu
	Add-TeamUser -GroupId $group.GroupId  -User paul.zoller@rme.eu
  	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User lisa.michaelis@rme.eu
  	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User claudia.lunderstedt-georgi@rme.eu
	Add-TeamUser -GroupId $group.GroupId  -User sabrina.peters@rme.eu
  	Start-Sleep -s 2
	Add-TeamUser -GroupId $group.GroupId  -User andrea.lohmann@rme.eu
  	
   	


}

 if($number.contains("7"))



		{



    Start-Sleep -s 40        


 
   Start-Sleep -s 2
   
   Add-TeamUser -GroupId $group.GroupId  -User frank.bildhauer@niersberger.de -Role Owner
   
   Start-Sleep -s 12
           
   Set-TeamPicture -GroupId $group.GroupId -ImagePath "$env:HOMEPATH\C:\Users\tristan.boehling\OneDrive - Wittl-IT\Bilder\Niersbeger-Logos\Niersberger-N 192x192.png"


   

		}	


}

write-host $Error

$prompt = Read-Host "Hit enter to end the script..."

Disconnect-PSSession

