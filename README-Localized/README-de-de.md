---
page_type: sample
products:
- office-powerpoint
- office-excel
- office-365
- office-onedrive
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  - Microsoft Graph
  services:
  - Excel
  - Office 365
  - OneDrive
  createdDate: 3/17/2016 9:42:20 AM
---
 Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In 

Erfahren Sie, wie Sie ein Microsoft Office-Add-In erstellen, das eine Verbindung mit Microsoft Graph herstellt, nach allen in OneDrive for Business gespeicherten Arbeitsmappen sucht, alle Diagramme in den Arbeitsmappen mithilfe der Excel REST-APIs abruft und ein Bild eines Diagramms in eine PowerPoint-Folie mit Office.js einfügt.

![Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In – Beispiel](images/InsertChart.png)

## Einführung

Integrieren von Daten von Onlinedienstanbietern erhöht den Wert und die Akzeptanz Ihrer Add-Ins. Dieses Codebeispiel zeigt, wie Sie das Add-In mit Microsoft Graph verbinden. Verwenden Sie dieses Codebeispiel für folgende Aufgaben:

* Herstellen einer Verbindung zwischen einem Office-Add-In und Microsoft Graph.
* Verwenden der MSAL .NET-Bibliothek, um das OAuth 2.0-Autorisierungsframework in einem Add-In zu implementieren.
* Verwenden von Excel- und OneDrive-REST-APIs in Microsoft Graph.
* Anzeigen eines Dialogfelds mit dem Office-Benutzeroberflächen-Namespace.
* Erstellen eines Add-Ins mithilfe von ASP.NET MVC, MSAL und Office.js. 
* Verwenden von Add-In-Befehlen in PowerPoint.


## Anforderungen

Damit dieses Codebeispiel ausgeführt wird, gelten die folgenden Anforderungen.

* Visual Studio 2019 oder höher.

* SQL Server Express (wird mit neueren Versionen von Visual Studio nicht mehr automatisch installiert.)

* Ein Office 365-Konto mit einem kostenlosen 1-jährigen Abonnement für Office 365, das Sie durch die Teilnahme am [Office 365-Entwicklerprogramm](https://aka.ms/devprogramsignup) erhalten.

* Excel-Arbeitsmappen (mit Diagrammen), die in OneDrive for Business in Ihrem Office 365-Abonnement gespeichert sind.

* PowerPoint für Windows Desktop Version 16.0.6769.2001 oder höher.
* [Office-Entwicklertools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Ein Microsoft Azure-Mandant. Für dieses Add-In ist Azure Active Directiory (AD) erforderlich. Von Azure AD werden Identitätsdienste bereitgestellt, die durch Anwendungen für die Authentifizierung und Autorisierung verwendet werden. Hier kann ein Testabonnement erworben werden: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Konfigurieren des Projekts

1. Wählen Sie in **Visual Studio** das Projekt **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb** aus. Stellen Sie unter **Eigenschaften** sicher, dass **SSL-aktiviert** den Wert **True** aufweist. Überprüfen Sie, ob die **SSL-URL** den gleichen Domänennamen und gleiche Portnummer wie in Schritt 3 aufgeführt verwendet.
 
2. Stellen Sie sicher, dass Ihr Azure-Abonnement an Ihren Office 365-Mandanten gebunden ist. Weitere Informationen finden Sie im folgenden Blogbeitrag des Active Directory-Teams: [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx) (Erstellen und Verwalten mehrerer Windows Azure Active Directory-Instanzen). Im Abschnitt **Adding a new directory** (Hinzufügen eines neuen Verzeichnisses) finden Sie Informationen über die entsprechende Vorgehensweise. Weitere Informationen finden Sie zudem unter [Einrichten Ihrer Office 365-Entwicklungsumgebung](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) im Abschnitt **Verknüpfen Ihres Office 365-Kontos mit Azure AD zum Erstellen und Verwalten von Apps**.

3. Registrieren Sie Ihre Anwendung über das [Azure-Verwaltungsportal](https://manage.windowsazure.com). Melden Sie sich mit einem Administratorkonto oder Ihrem Office 365-Abonnement an. Weitere Informationen zum Registrieren Ihrer Anwendungen finden Sie unter [Registrieren einer Anwendung bei der Microsoft Identity Platform](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually). Verwenden Sie die folgenden Einstellungen:

 - REDIRCT URI: https://localhost:44301/AzureADAuth/Authorize	
 - UNTERSTÜTZTE KONTOTYPEN: "Nur Konten in diesem Organisationsverzeichnis"
 - IMPLIZITE GEWÄHRUNG: Aktivieren Sie keine impliziten Gewährungsoptionen.
 - API-BERECHTIGUNGEN: **Files.Read.All** und **User.Read**

	> Hinweis: Nachdem Sie die Anwendung registriert haben, kopieren Sie die **Anwendungs-ID (Client-ID)** und die **Verzeichnis-ID (Mandanten-ID)** auf dem Blatt **Übersicht** der App-Registrierung im Azure-Verwaltungsportal. Wenn Sie auf dem Blatt **Zertifikate und Geheimnisse** den geheimen Clientschlüssel erstellen, kopieren Sie auch diesen Wert. 
	 
4.  Verwenden Sie die im vorherigen Schritt kopierten in "web.config". Legen Sie **AAD:ClientID** auf Ihre Client-ID, **AAD:ClientSecret** auf Ihren geheimen Clientschlüssel und **"AAD:O365TenantID"** auf Ihre Mandanten-ID fest. 

## Ausführen des Projekts
1. Öffnen Sie die Visual Studio-Projektmappe. 
2. Klicken Sie mit der rechten Maustaste auf **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, und wählen Sie dann **Als Startprojekt festlegen** aus.
2. Drücken Sie F5. 
3. Öffnen Sie in PowerPoint die Registerkarte **Einfügen**, und wählen Sie **Diagramm auswählen** aus, um das Aufgabenbereich-Add-In zu öffnen.

## Bekannte Probleme

* Szenario: Beim Versuch, das Codebeispiel auszuführen, wird das Add-In nicht geladen.
	* Lösung: 
		1. Öffnen Sie in Visual Studio **SQL Server-Objekt-Explorer**.
		2. Erweitern Sie **(localdb)\\MSSQLLocalDB** > **Datenbanken**.
		3. Klicken Sie mit der rechten Maustaste auf **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, und wählen Sie dann **Löschen** aus. 
* Szenario: Beim Ausführen des Codebeispiels tritt in der Zeile *Office.context.ui.messageParent* ein Fehler auf.	
	* Lösung: Beenden Sie die Ausführung des Codebeispiels, und starten Sie es erneut. 
* Beim Herunterladen der ZIP-Datei wird beim Extrahieren der Dateien eine Fehlermeldung mit dem Hinweis angezeigt, dass der Dateipfad zu lang ist.
	* Lösung: Entpacken Sie Ihre Dateien in einen Ordner direkt unter dem Stamm (z. B. C:\\sample).

## Fragen und Kommentare
Wir schätzen Ihr Feedback hinsichtlich des Beispiels *Einfügen von Excel-Diagrammen mit Microsoft Graph in einem PowerPoint-Add-In*. Sie können uns Ihr Feedback über den Abschnitt *Probleme* dieses Repositorys senden. Allgemeine Fragen zur Office 365-Entwicklung sollten in [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API) gestellt werden. Stellen Sie sicher, dass Ihre Fragen mit \[office-js], \[MicrosoftGraph] und \[API] markiert sind.

## Zusätzliche Ressourcen

* [Microsoft Graph (Excel) ToDo-Codebeispiel](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Microsoft Graph-Dokumentation](https://docs.microsoft.com/en-us/graph/)
* [Dokumentation zu Office-Add-Ins](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2016-2019 Microsoft Corporation. Alle Rechte vorbehalten.



In diesem Projekt wurden die [Microsoft Open Source-Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/) übernommen. Weitere Informationen finden Sie unter [Häufig gestellte Fragen zu Verhaltensregeln](https://opensource.microsoft.com/codeofconduct/faq/), oder richten Sie Ihre Fragen oder Kommentare an [opencode@microsoft.com](mailto:opencode@microsoft.com).
