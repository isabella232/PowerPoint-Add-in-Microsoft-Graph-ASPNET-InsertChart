# <a name="insert-excel-charts-using-microsoft-graph-in-a-powerpoint-add-in"></a>Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint 

Узнайте, как создать надстройку Microsoft Office, которая подключается к Microsoft Graph, находит все книги, сохраненные в OneDrive для бизнеса, получает все их диаграммы с помощью REST API для Excel и вставляет изображение диаграммы в слайд PowerPoint с помощью Office.js.

![Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](../images/InsertChart.png)

## <a name="introduction"></a>Введение

Интегрируя данные поставщиков интернет-служб, вы повышаете ценность и популярность своих надстроек. В этом примере кода показано, как подключить надстройку к Microsoft Graph. С его помощью можно:

* подключиться к Microsoft Graph из надстройки Office;
* использовать в надстройке платформу проверки подлинности OAuth 2.0;
* использовать REST API для Excel и OneDrive из Microsoft Graph;
* отображать диалоговое окно с использованием пространства имен пользовательского интерфейса Office;
* создать надстройку с помощью ASP.NET MVC и Office.js; 
* использовать команды надстроек в PowerPoint.


## <a name="prerequisites"></a>Необходимые компоненты
Чтобы запустить этот пример кода, необходимо следующее:

* Visual Studio 2015.

* Учетная запись Office 365, которую получают участники [Программы для разработчиков Office 365](https://aka.ms/devprogramsignup), предоставляется вместе с бесплатной годичной подпиской на Office 365.

* Книги Excel (с диаграммами), сохраненные в OneDrive для бизнеса в составе подписки на Office 365.

* PowerPoint для Windows Desktop версии не ниже 16.0.6769.2001.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Клиент Microsoft Azure. Для использования этой надстройки требуется Azure Active Directory (AD), где доступны службы удостоверений, которые приложения используют для проверки подлинности и авторизации. Здесь можно получить пробную подписку: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="configure-the-project"></a>Настройка проекта

1. В **Visual Studio** выберите проект **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. Убедитесь, что в окне **Свойства** для параметра **SSL включен** задано значение **True**. Убедитесь, что в поле **URL-адрес SSL** указаны доменное имя и номер порта, упомянутые на этапе 3.
 
2. Убедитесь, что ваша подписка на Azure связана с клиентом Office 365. Дополнительные сведения об этом см. в записи блога команды Active Directory, посвященной [созданию нескольких каталогов Azure AD и управлению ими](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). Соответствующие инструкции приведены в разделе о **добавлении нового каталога**. Дополнительные сведения см. в статье [Как настроить среду разработки для Office 365](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) и, в частности, разделе **Связывание Azure AD и учетной записи Office 365 для создания приложений и управления ими**.

3. Зарегистрируйте свое приложение на [портале управления Azure](https://manage.windowsazure.com). Сведения о том, как это сделать, см. в разделе [Регистрация браузерного веб-приложения на портале управления Azure](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp). Используйте следующие параметры:

 - URL-АДРЕС ВХОДА: https://localhost:44301/AzureADAuth/Authorize 
 - URI КОДА ПРИЛОЖЕНИЯ: https://localhost:44301.
 - URL-АДРЕС ОТВЕТА: https://localhost:44301/AzureADAuth/Authorize. 

    > Примечание. Зарегистрировав приложение, скопируйте идентификатор и секрет клиента, отображенные на портале управления Azure.
     
4. Предоставьте разрешения для приложения.
    *  На портале управления Azure откройте вкладку **Active Directory** и выберите клиент Office 365.
    *  Откройте вкладку **Приложения** и выберите приложение, которое нужно настроить. Выберите элемент **Настроить**.
    *  В разделе **Разрешения для других приложений** добавьте **Microsoft Graph**.
    *  В разделе **Делегированные разрешения** выберите элемент **Чтение файлов пользователя, а также файлов, которыми с ним поделились**.

5.  В узле web.config для параметра **AAD:ClientID** задайте значение идентификатора клиента, а для параметра **AAD:ClientSecret** — значение секрета клиента. 

## <a name="run-the-project"></a>Запуск проекта
1. Откройте файл решения в Visual Studio. 
2. Щелкните правой кнопкой мыши проект **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** и выберите команду **Назначить запускаемые проекты**.
2. Нажмите клавишу F5. 
3. В PowerPoint откройте вкладку **Вставка**  >  **Выбор диаграммы**, чтобы открыть надстройку области задач.

## <a name="known-issues"></a>Известные проблемы

* Сценарий. При попытке запустить пример кода надстройка не загружается.
    * Решение. 
        1. В Visual Studio откройте **обозреватель объектов SQL Server**.
        2. Разверните **(localdb)\MSSQLLocalDB** > **Базы данных**.
        3. Щелкните правой кнопкой мыши элемент **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** и выберите команду **Удалить**. 
* Сценарий: При запуске примера кода в строке *Office.context.ui.messageParent* возникает ошибка.   
    * Решение. Остановите выполнение примера кода и перезапустите его. 
* При распаковке скачанного ZIP-файла возникает ошибка и отображается сообщение о том, что путь к файлу слишком длинный.
    * Решение. Распакуйте файлы в папку непосредственно под корневой (например, C:\sample).

## <a name="questions-and-comments"></a>Вопросы и комментарии
Мы будем рады получить ваши отзывы о примере *вставки диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint*. Своими мыслями можете поделиться на вкладке *Issues* (Проблемы) этого репозитория. Общие вопросы о разработке решений для Office 365 следует задавать на сайте [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Помечайте свои вопросы тегами [office-js], [MicrosoftGraph] и [API].

## <a name="additional-resources"></a>Дополнительные ресурсы

* [Пример кода для использования списка дел в Microsoft Graph (Excel)](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Документация по Microsoft Graph](https://graph.microsoft.io/en-us/docs)
* [Документация по надстройкам Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* Посмотрите видео в разделе Build, содержащее [обзор платформы Office](https://channel9.msdn.com/Events/Build/2016/B872 "Обзор платформы Office").

## <a name="copyright"></a>Авторское право
© Корпорация Майкрософт (Microsoft Corporation), 2016. Все права защищены.



Этот проект соответствует [правилам поведения Майкрософт, касающимся обращения с открытым кодом](https://opensource.microsoft.com/codeofconduct/). Дополнительную информацию см. в разделе [часто задаваемых вопросов по правилам поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
