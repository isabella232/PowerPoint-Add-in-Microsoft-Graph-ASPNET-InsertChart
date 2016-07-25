# Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In 

Erfahren Sie, wie Sie ein Microsoft Office-Add-In erstellen, das eine Verbindung mit Microsoft Graph herstellt, nach allen in OneDrive for Business gespeicherten Arbeitsmappen sucht, alle Diagramme in den Arbeitsmappen mithilfe der Excel REST-APIs abruft und ein Bild eines Diagramms in eine PowerPoint-Folie mit Office.js einfügt.

![Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In – Beispiel](../images/InsertChart.png)

## Einführung

Integrieren von Daten von Onlinedienstanbietern erhöht den Wert und die Akzeptanz Ihrer Add-Ins. Dieses Codebeispiel zeigt, wie Sie das Add-In mit Microsoft Graph verbinden. Verwenden Sie dieses Codebeispiel für folgende Aufgaben:

* Herstellen einer Verbindung zwischen einem Office-Add-In und Microsoft Graph
* Verwenden des OAuth 2.0-Autorisierungsframeworks in einem Add-In
* Verwenden von Excel- und OneDrive-REST-APIs in Microsoft Graph
* Anzeigen eines Dialogfelds mit dem Office-Benutzeroberflächennamespace
* Erstellen eines Add-Ins mithilfe von ASP.NET MVC und Office.js 
* Verwenden von Add-In-Befehlen in PowerPoint


## Anforderungen
Damit dieses Codebeispiel ausgeführt wird, gelten die folgenden Anforderungen.

* Visual Studio 2015

* Ein Office 365-Konto, das Sie beim Beitreten zum <a herf="https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=14b845d0-938c-45af-b061-f798fbb4d170&amp;lcid=1033">Office 365 Developer Program</a> erhalten, das ein kostenloses 1-Jahres-Abonnement für Office 365 enthält.

* Excel-Arbeitsmappen (mit Diagrammen), die in OneDrive for Business in Ihrem Office 365-Abonnement gespeichert sind

* PowerPoint für Windows Desktop Version 16.0.6769.2001 oder höher
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Ein Microsoft Azure-Mandant. Dieses Add-In erfordert Azure Active Directory (AD). Von Azure Active Directory (AD) werden Identitätsdienste bereitgestellt, die durch Anwendungen für die Authentifizierung und Autorisierung verwendet werden. Hier kann ein Testabonnement erworben werden: [Microsoft Azure](https://account.windowsazure.com/SignUp)

## Konfigurieren des Projekts

1. Wählen Sie in **Visual Studio** das Projekt **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. Stellen Sie unter **Eigenschaften** sicher, dass **SSL-aktiviert** den Wert **True** aufweist. Überprüfen Sie, ob die **SSL-URL** den gleichen Domänennamen und gleiche Portnummer wie in Schritt 3 aufgeführt verwendet.
 
2. Sie müssen sicherstellen, dass Ihr Azure-Abonnement an Ihren Office 365-Mandanten gebunden ist. Rufen Sie für weitere Informationen dazu den Blogpost [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx) des Active Directory-Teams auf. Im Abschnitt **Adding a new directory** finden Sie Informationen über die entsprechende Vorgehensweise. Weitere Informationen finden Sie zudem unter [Einrichten Ihrer Office 365-Entwicklungsumgebung](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) im Abschnitt **Verknüpfen Ihres Office 365-Kontos mit Azure AD zum Erstellen und Verwalten von Apps**.

3. Registrieren Sie Ihre Anwendung über das [Azure-Verwaltungsportal](https://manage.windowsazure.com). Informationen zur Registrierung Ihrer Anwendung finden Sie unter [Registrieren der browserbasierten Web-App mit dem Azure-Verwaltungsportal](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp). Verwenden Sie die folgenden Einstellungen:

 - ANMELDE-URL: https://localhost:44301/AzureADAuth/Authorize 
 - APP-ID-URI: https://localhost:44301
 - ANTWORT-URL: https://localhost:44301/AzureADAuth/Authorize	

	> Hinweis: Kopieren Sie nach Registrierung der Anwendung die Client-ID und den geheimen Clientschlüssel, der im Azure-Verwaltungsportal angezeigt wird.
	 
4. Gewähren Sie Ihrer Anwendung entsprechende Berechtigungen.
	*  Wählen Sie im Azure-Verwaltungsportal die Registerkarte **Active Directory** und einen Office 365-Mandanten.
	*  Wählen Sie die Registerkarte **Anwendungen**, und klicken Sie auf die Anwendung, die Sie konfigurieren möchten. Wählen Sie **Konfigurieren**.
	*  Fügen Sie unter **Berechtigungen für andere Anwendungen****Microsoft Graph** hinzu.
	*  Wählen Sie unter **Delegierte Berechtigungen**, die Option **Benutzerdateien und Dateien lesen, die für den Benutzer freigegeben wurden**.

5.  Legen Sie in web.config **AAD:ClientID** auf Ihre Client-ID fest, und legen Sie **AAD:ClientSecret** auf Ihren geheimen Clientschlüssel fest. 

## Ausführen des Projekts
1. Öffnen Sie die Visual Studio-Projektmappe. 
2. Klicken Sie mit der rechten Maustaste auf **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, und wählen Sie dann **Als Startprojekt festlegen **.
2. Drücken Sie F5. 
3. Wählen Sie in PowerPoint **Einfügen** > **Diagramm auswählen**, um das Aufgabenbereich-Add-In zu öffnen.

## Bekannte Probleme

* Szenario: Beim Versuch, das Codebeispiel auszuführen, wird das Add-In nicht geladen.
	* Lösung: 
		1. Öffnen Sie in Visual Studio **SQL Server-Objekt-Explorer**.
		2. Erweitern Sie **(localdb)\MSSQLLocalDB** > **Datenbanken**.
		3. Klicken Sie mit der rechten Maustaste auf **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, und wählen Sie dann **Löschen**. 
* Szenario: Beim Ausführen des Codebeispiels tritt in der Zeile *Office.context.ui.messageParent* ein Fehler auf.	
	* Lösung: Beenden Sie die Ausführung des Codebeispiels, und starten Sie es erneut. 
* Beim Herunterladen der ZIP-Datei wird beim Extrahieren der Dateien eine Fehlermeldung mit dem Hinweis angezeigt, dass der Dateipfad zu lang ist.
	* Lösung: Entpacken Sie Ihre Dateien in einen Ordner direkt unter dem Stamm (z. B. C:\sample).

## Fragen und Kommentare
Wir schätzen Ihr Feedback hinsichtlich des Beispiels *Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In*. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden. Allgemeine Fragen zur Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen mit [office-js], [MicrosoftGraph] und [API] markiert sind.

## Zusätzliche Ressourcen

* [Microsoft Graph (Excel) ToDo-Codebeispiel](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Microsoft Graph-Dokumentation](https://graph.microsoft.io/en-us/docs)
* [Dokumentation zu Office-Add-Ins](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* Schauen Sie sich das Video an unter //Build – [Übersicht über die Office-Plattform](https://channel9.msdn.com/Events/Build/2016/B872 "Übersicht über die Office-Plattform").

## Copyright
Copyright (c) 2016 Microsoft Corporation. Alle Rechte vorbehalten.


