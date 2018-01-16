# <a name="insert-excel-charts-using-microsoft-graph-in-a-powerpoint-add-in"></a>Insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint 

Aprenda a crear un complemento de Microsoft Office que se conecte a Microsoft Graph, busque todos los libros almacenados en OneDrive para la Empresa, capture todos los gráficos en los libros mediante las API de REST de Excel e inserte una imagen de un gráfico en una diapositiva de PowerPoint con Office.js.

![Ejemplo de insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint](../images/InsertChart.png)

## <a name="introduction"></a>Introducción

Integrar datos de proveedores de servicios en línea aumenta el valor y la adopción de los complementos. En este ejemplo de código se muestra cómo conectar el complemento con Microsoft Graph. Use este ejemplo de código para:

* Conectarse a Microsoft Graph desde un complemento de Office.
* Usar el marco de autorización OAuth 2.0 en un complemento.
* Usar las API de REST de OneDrive y Excel desde Microsoft Graph.
* Mostrar un diálogo mediante el espacio de nombres de la interfaz de usuario de Office.
* Crear un complemento mediante ASP.NET MVC y Office.js. 
* Usar comandos de complementos en PowerPoint.


## <a name="prerequisites"></a>Requisitos previos
Para ejecutar este ejemplo de código, se requiere lo siguiente.

* Visual Studio 2015.

* Una cuenta de Office 365, que puede obtener si se une al [Programa de desarrolladores de Office 365](https://aka.ms/devprogramsignup), que incluye una suscripción gratuita de 1 año a Office 365.

* Libros de Excel (con gráficos) almacenados en OneDrive para la Empresa en su suscripción a Office 365.

* PowerPoint para escritorio de Windows, versión 16.0.6769.2001 o superior.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un inquilino de Microsoft Azure. Este complemento requiere Azure Active Directory (AD). Azure AD proporciona servicios de identidad que las aplicaciones usan para autenticación y autorización. Puede adquirir una suscripción de prueba aquí: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## <a name="configure-the-project"></a>Configurar el proyecto

1. En **Visual Studio**, elija el proyecto **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. En **Propiedades**, asegúrese de que **SSL habilitado** esté en **True**. Compruebe que la **Dirección URL de SSL** usa el mismo nombre de dominio y número de puerto que los enumerados en el paso 3 siguiente.
 
2. Asegúrese de que su suscripción de Azure está enlazada a su inquilino de Office 365. Para obtener más información, consulte la publicación del blog del equipo de Active Directory, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx) (Crear y administrar varios directorios de Windows Azure Active Directory). En la sección **Adding a new directory** (Agregar un nuevo directorio) se le explicará cómo hacerlo. Para más información, también puede consultar [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) (Configurar el entorno de desarrollo de Office 365) y la sección **Associate your Office 365 account with Azure AD to create and manage apps** (Asociar su cuenta de Office 365 con Azure AD para crear y administrar aplicaciones).

3. Registre la aplicación mediante el [Portal de administración de Azure](https://manage.windowsazure.com). Para saber cómo registrar la aplicación, consulte [Register your browser-based web app with the Azure Management Portal](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp) (Registrar la aplicación web basada en explorador con el Portal de administración de Azure). Use la configuración siguiente:

 - URL DE INICIO DE SESIÓN: https://localhost:44301/AzureADAuth/Authorize 
 - URI del identificador de la aplicación: https://localhost:44301
 - URL de respuesta: https://localhost:44301/AzureADAuth/Authorize 

    > Nota: Después de registrar la aplicación, copie el id. de cliente y el secreto de cliente que se muestra en el Portal de administración de Azure.
     
4. Conceda permisos a la aplicación.
    *  En el Portal de administración de Azure, seleccione la pestaña **Active Directory** y un espacio empresarial de Office 365.
    *  Seleccione la pestaña **Aplicaciones** y haga clic en la aplicación que quiera configurar. Elija **Configurar**.
    *  En los **permisos para otras aplicaciones**, agregue **Microsoft Graph**.
    *  En **Permisos delegados**, elija **Leer los archivos del usuario y los archivos compartidos con el usuario**.

5.  En web.config, establezca **AAD:ClientID** en el id. de cliente y establezca **AAD:ClientSecret** en el secreto de cliente. 

## <a name="run-the-project"></a>Ejecutar el proyecto
1. Abra el archivo de la solución de Visual Studio. 
2. Haga clic con el botón derecho en **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** y, después, elija **Establecer como proyecto de inicio**.
2. Pulse F5. 
3. En PowerPoint, elija **Insertar** > **Elegir un gráfico** para abrir el complemento del panel de tareas.

## <a name="known-issues"></a>Problemas conocidos

* Escenario: Al intentar ejecutar el código de ejemplo, el complemento no se cargará.
    * Solución: 
        1. En Visual Studio, abra **Explorador de objetos de SQL Server**.
        2. Expanda **(localdb)\MSSQLLocalDB** > **Bases de datos**.
        3. Haga clic con el botón derecho en **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** y, después, elija **Eliminar**. 
* Escenario: Cuando ejecuta el ejemplo de código, obtiene un error en la línea *Office.context.ui.messageParent*.   
    * Solución: Deje de ejecutar el ejemplo de código y reinícielo. 
* Si descarga el archivo zip, al extraer los archivos obtiene un error que indica que la ruta del archivo es demasiado larga.
    * Solución: Descomprima los archivos en una carpeta directamente bajo la raíz (por ejemplo, c:\sample).

## <a name="questions-and-comments"></a>Preguntas y comentarios
Nos encantaría recibir sus comentarios sobre el ejemplo de *insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint*. Puede enviarnos comentarios a través de la sección *Problemas* de este repositorio. Las preguntas generales sobre desarrollo en Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrese de que sus preguntas están etiquetadas con [office-js], [MicrosoftGraph] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* [Ejemplo de código de tarea pendiente de Microsoft Graph (Excel)](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Documentación de Microsoft Graph](https://graph.microsoft.io/en-us/docs)
* [Documentación de complementos de Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* Vea el vídeo de //Build - [Office Platform Overview (Información general sobre la plataforma de Office)](https://channel9.msdn.com/Events/Build/2016/B872 "Office Platform Overview (Información general sobre la plataforma de Office)").

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte las [preguntas más frecuentes sobre el Código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
