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
 Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint 

Узнайте, как создать надстройку Microsoft Office, которая подключается к Microsoft Graph, находит все книги, сохраненные в OneDrive для бизнеса, получает все диаграммы из них с помощью REST API для Excel и вставляет изображение диаграммы в слайд PowerPoint с помощью Office.js.

![Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](../images/InsertChart.png)

## Введение

Интегрируя данные поставщиков интернет-служб, вы повышаете ценность и популярность своих надстроек. В этом примере кода показано, как подключить надстройку к Microsoft Graph. С его помощью можно:

* подключиться к Microsoft Graph из надстройки Office;
* использовать библиотеку MSAL .NET для внедрения инфраструктуры авторизации OAuth 2.0 в надстройке;
* использовать REST API для Excel и OneDrive из Microsoft Graph;
* отображать диалоговое окно с использованием пространства имен пользовательского интерфейса Office;
* создать надстройку с помощью ASP.NETMVC,MSAL и Office.js; 
* использовать команды надстроек в PowerPoint.


## Необходимые компоненты

Чтобы запустить этот пример кода, необходимо следующее:

* Visual Studio 2019 или более поздней версии.

* SQL Server Express (больше не устанавливается автоматически с последними версиями Visual Studio).

* Учетная запись Office 365, которую получают участники [Программы для разработчиков Office 365](https://aka.ms/devprogramsignup), предоставляется вместе с бесплатной годичной подпиской на Office 365.

* Книги Excel (с диаграммами), сохраненные в OneDrive для бизнеса в составе подписки на Office 365.

* PowerPoint для Windows Desktop версии не ниже 16.0.6769.2001.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Клиент Microsoft Azure. Эта надстройка требует наличия Azure Active Directiory (AD). В Azure AD доступны службы идентификации, которые приложения используют для проверки подлинности и авторизации. Здесь можно получить пробную подписку: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Настройка проекта

1. В **Visual Studio** выберите проект **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. Убедитесь, что в окне **Свойства** для параметра **SSL включен** задано значение **Иcтина**. Убедитесь, что в поле **URL-адрес SSL** используются доменное имя и номер порта, указанные на этапе 3.
 
2. Убедитесь, что ваша подписка Azure привязана к клиенту Office 365. Для получения дополнительных сведений просмотрите запись в блоге команды Active Directory, посвященную [созданию нескольких каталогов Windows Azure Active Directories и управлению ими](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). Инструкции приведены в разделе о **добавлении нового каталога**. Дополнительные сведения см. в статье [Как настроить среду разработки для Office 365](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) и, в частности, в разделе **Связывание Azure AD и учетной записи Office 365 для создания приложений и управления ими**.

3. Зарегистрируйте свое приложение на [портале управления Azure](https://manage.windowsazure.com). Войдите в систему с учетной записью администратора или своей подписки на Office 365. Сведения о регистрации приложений см. в статье [Регистрация приложения с помощью платформы удостоверений Майкрософт](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually). Используйте указанные ниже параметры:

 - URI ПЕРЕНАПРАВЛЕНИЯ: https://localhost:44301/AzureADAuth/Authorize	
 - ПОДДЕРЖИВАЕМЫЕ ТИПЫ УЧЕТНЫХ ЗАПИСЕЙ "Учетные записи только в этом каталоге организации"
 - НЕЯВНОЕ ПРЕДОСТАВЛЕНИЕ РАЗРЕШЕНИЯ: Не включайте никакие параметры неявного предоставления разрешений
 - РАЗРЕШЕНИЯ API: **Files.Read.All** и **User.Read**

	> Примечание. После регистрации приложения скопируйте **идентификатор приложения (клиента)** и **идентификатор директории (клиента)** в колонке **Обзор** регистрации приложения на портале управления Azure. Также скопируйте секретный код клиента, созданный в колонке **Сертификаты и секреты**. 
	 
4.  В узле web.config используйте значения, скопированные на предыдущем этапе. Для параметра **AAD:ClientID** задайте значение идентификатора клиента, а для параметра **AAD:ClientSecret** — значение секретного кода клиента. Задайте ваш идентификатор клиента Office 365 в **"AAD:O365TenantID"**. 

## Запуск проекта
1. Откройте файл решения в Visual Studio. 
2. Щелкните правой кнопкой мыши проект **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** и выберите команду **Назначить запускаемым проектом**.
2. Нажмите клавишу F5. 
3. В PowerPoint откройте вкладку **Вставка** и нажмите **Выбрать диаграмму**, чтобы открыть надстройку области задач.

## Известные проблемы

* Сценарий. При попытке запустить пример кода надстройка не загружается.
	* Решение: 
		1. В Visual Studio откройте **обозреватель объектов SQL Server**.
		2. Разверните узел **(localdb)\\MSSQLLocalDB** > **Базы данных**.
		3. Щелкните правой кнопкой мыши элемент **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** и выберите команду **Удалить**. 
* Сценарий: При запуске примера кода в строке *Office.context.ui.messageParent* возникает ошибка.	
	* Решение: Остановите выполнение примера кода и перезапустите его. 
* При распаковке скачанного ZIP-файла возникает ошибка и отображается сообщение о том, что путь к файлу слишком длинный.
	* Решение. Распакуйте файлы в папку непосредственно под корневой (например, C:\\sample).

## Вопросы и комментарии
Мы будем рады получить ваши отзывы о примере *вставки диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint*. Своими мыслями можете поделиться на вкладке *Проблемы* этого репозитория. Общие вопросы о разработке решений для Office 365 следует публиковать на сайте [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Помечайте свои вопросы тегами \[office-js], \[MicrosoftGraph] и \[API].

## Дополнительные ресурсы

* [Пример кода для использования списка дел в Microsoft Graph (Excel)](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Документация по Microsoft Graph](https://docs.microsoft.com/en-us/graph/)
* [Документация по надстройкам Office](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## Авторские права
(с) Корпорация Майкрософт (Microsoft Corporation), 2016 - 2019. Все права защищены.



Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
