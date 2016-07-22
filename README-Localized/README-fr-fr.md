# Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint 

Découvrez comment créer un complément Microsoft Office qui se connecte à Microsoft Graph, recherche tous les classeurs stockés dans OneDrive Entreprise, extrait tous les graphiques dans des classeurs à l’aide de l’API REST Excel et insère une image d’un graphique dans une diapositive PowerPoint à l’aide d’Office.js.

![Insérer des graphiques Excel à l’aide de Microsoft Graph dans un exemple de complément PowerPoint](../images/InsertChart.png)

## Présentation

Le fait d’intégrer des données à partir de fournisseurs de services en ligne augmente la valeur et l’adoption de vos compléments. Cet exemple de code vous montre comment connecter votre complément à Microsoft Graph. Utilisez cet exemple de code pour :

* Se connecter à Microsoft Graph à partir d’un complément Office.
* Utiliser l’authentification OAuth 2.0 dans un complément.
* Utiliser les API REST Excel et OneDrive à partir de Microsoft Graph.
* Afficher une boîte de dialogue à l’aide de l’espace de noms de l’interface utilisateur Office.
* Créer un complément à l’aide d’ASP.NET MVC et d’Office.js. 
* Utiliser les commandes de complément dans PowerPoint.


## Conditions requises
Pour exécuter cet exemple de code, les éléments suivants sont requis.

* Visual Studio 2015.

* Un compte Office 365 que vous pouvez obtenir en rejoignant le <a herf="https://profile.microsoft.com/RegSysProfileCenter/wizardnp.aspx?wizid=14b845d0-938c-45af-b061-f798fbb4d170&amp;lcid=1033">programme de développement Office 365</a> qui inclut un abonnement gratuit d’un an à Office 365.

* Des classeurs Excel (avec des graphiques) stockés sur OneDrive Entreprise dans votre abonnement Office 365.

* PowerPoint pour Bureau Windows version 16.0.6769.2001 ou ultérieure.
* [Outils de développement Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un client Microsoft Azure. Ce complément exige Azure Active Directory (AD). Azure AD fournit des services d’identité que les applications utilisent à des fins d’authentification et d’autorisation. Un abonnement d’évaluation peut être demandé ici : [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configurer le projet

1. Dans **Visual Studio**, choisissez le projet **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. Dans **Propriétés**, assurez-vous que **SSL activé** est défini sur **True**. Vérifiez que l’**URL SSL** utilise le même nom de domaine et le même numéro de port que ceux répertoriés à l’étape 3 ci-dessous.
 
2. Assurez-vous que votre abonnement Azure est lié à votre client Office 365. Pour plus d’informations, consultez le billet du blog de l’équipe d’Active Directory relatif à la [création et la gestion de plusieurs fenêtres dans les répertoires Azure Active Directory](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). La section sur l’**ajout d’un nouveau répertoire** vous explique comment procéder. Pour en savoir plus, vous pouvez également consulter la rubrique relative à la [configuration de votre environnement de développement Office 365](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) et la section sur l’**association de votre compte Office 365 à Azure Active Directory pour créer et gérer des applications**.

3. Inscrivez votre application à l’aide du [portail de gestion Azure](https://manage.windowsazure.com). Pour découvrir comment inscrire votre application, consultez la page relative à l’[enregistrement de votre application web basée sur le navigateur dans le portail de gestion Azure](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp). Utilisez les paramètres suivants :

 - URL de connexion : https://localhost:44301/AzureADAuth/Authorize 
 - URI d’ID d’application : https://localhost:44301
 - URL de réponse : https://localhost:44301/AzureADAuth/Authorize	

	> Remarque : Après avoir enregistré votre application, copiez l’ID client et le secret client figurant sur le portail de gestion Azure.
	 
4. Accordez des autorisations à votre application.
	*  Dans le portail de gestion Azure, sélectionnez l’onglet **Active Directory** et un client Office 365.
	*  Sélectionnez l’onglet **Applications**, puis cliquez sur l’application que vous souhaitez configurer. Choisissez **Configurer**.
	*  Dans **Autorisations pour d’autres applications**, ajoutez **Microsoft Graph**.
	*  Dans **Autorisations déléguées**, choisissez **Lire les fichiers utilisateur et les fichiers partagés avec l’utilisateur**.

5.  Dans web.config, définissez **AAD:ClientID** sur votre ID client et définissez **AAD:ClientSecret** sur votre secret client. 

## Exécuter le projet
1. Ouvrez le fichier de solution Visual Studio. 
2. Cliquez avec le bouton droit sur **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, puis choisissez **Définir comme projet de démarrage**.
2. Appuyez sur la touche F5. 
3. Dans PowerPoint, choisissez **Insertion** > **Sélectionner un graphique** pour ouvrir le complément de volet Office.

## Problèmes connus

* Scénario : Lorsque vous tentez d’exécuter l’exemple de code, le complément ne se charge pas.
	* Résolution : 
		1. Dans Visual Studio, ouvrez **Explorateur d’objets SQL Server**.
		2. Développez **(localdb)\MSSQLLocalDB** > **Bases de données**.
		3. Cliquez avec le bouton droit sur **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart**, puis choisissez **Supprimer**. 
* Scénario : Lorsque vous exécutez l’exemple de code, vous obtenez une erreur sur la ligne *Office.context.ui.messageParent*.	
	* Résolution : Arrêtez l’exécution de l’exemple de code et redémarrez-le. 
* Si vous téléchargez un fichier ZIP, lorsque vous extrayez les fichiers, vous obtenez une erreur indiquant que le chemin d’accès du fichier est trop long.
	* Résolution : Décompressez vos fichiers dans un dossier directement sous la racine (par exemple, C:\sample).

## Questions et commentaires
Nous aimerions recevoir vos commentaires relatifs à l’exemple *Insérer des graphiques Excel à l’aide de Microsoft Graph dans un complément PowerPoint*. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel. Si vous avez des questions sur le développement d’Office 365, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Assurez-vous que vos questions comportent les balises [office-js], [MicrosoftGraph] et [API].

## Ressources supplémentaires

* [Exemple de code ToDo pour Microsoft Graph (Excel)](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Documentation Microsoft Graph](https://graph.microsoft.io/en-us/docs)
* [Documentation de compléments Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* Visionnez la vidéo à partir de //Build - [Vue d’ensemble de la plateforme Office](https://channel9.msdn.com/Events/Build/2016/B872 "Présentation de la plate-forme Office").

## Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.


