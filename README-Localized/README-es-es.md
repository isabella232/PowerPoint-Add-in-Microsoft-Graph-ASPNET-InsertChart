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
 Insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint 

Obtenga información acerca de cómo crear un complemento de Microsoft Office que se conecte a Microsoft Graph, busque todos los libros almacenados en OneDrive para la Empresa, capture todos los gráficos en los libros mediante las API de REST de Excel, e inserte una imagen de un gráfico en una diapositiva de PowerPoint con Office.js.

![Ejemplo de insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint](../images/InsertChart.png)

## Introducción

Integrar datos de proveedores de servicios en línea aumenta el valor y la adopción de los complementos. En este ejemplo de código se muestra cómo conectar el complemento con Microsoft Graph. Use este ejemplo de código para:

* Conectarse a Microsoft Graph desde un complemento de Office.
* Use la biblioteca MSAL .NET para implementar el marco de autorización OAuth 2.0 en un complemento.
* Usar las APIs de REST de OneDrive y Excel desde Microsoft Graph.
* Mostrar un diálogo mediante el espacio de nombres de la interfaz de usuario de Office.
* Crear un complemento mediante ASP.NET MVC y Office.js. 
* Usar comandos de complementos en PowerPoint.


## Requisitos previos

Para ejecutar este ejemplo de código, se requiere lo siguiente.

* Visual Studio 2019 o posterior.

* SQL Server Express (ya no se instala automáticamente con versiones recientes de Visual Studio).

* Una cuenta de Office 365 que puede obtener al unirse al [programa para desarrolladores de Office 365](https://aka.ms/devprogramsignup) que incluye una suscripción gratuita de 1 año a Office 365.

* Libros de Excel (con gráficos) almacenados en OneDrive para la Empresa en su suscripción de Office 365. 

* PowerPoint para escritorio de Windows, versión 16.0.6769.2001 o superior.
* [Herramientas para desarrolladores de Office](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Un inquilino de Microsoft Azure. Este complemento requiere Azure Active Directiory (AD).  Azure (AD) le ofrece servicios de identidad que las aplicaciones usan para autenticación y autorización. Las suscripciones de prueba se pueden adquirir aquí: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configurar el proyecto

1. En **Visual Studio**, elija el proyecto**PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. En **Propiedades**, asegúrese de que **SSL esté activado** y sea **Verdadero**. Compruebe que la **Dirección URL de SSL ** use el mismo nombre de dominio y número de puerto que se indica en el paso 3 que se muestra a continuación.
 
2. Debe asegurarse de que su suscripción de Azure esté vinculada a su inquilino de Office 365. Para obtener más información, puede consultar la entrada del blog del equipo de Active Directory, [Creación y administración de varios directorios de Windows Azure Active](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). En la **sección agregar un directorio** se explica cómo realizar esta acción. También puede ver [configurar el entorno](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) de desarrollo de Office 365 y la **sección asociar la cuenta de Office 365 Azure AD para crear** y administrar aplicaciones y obtener más información.

3. Registre la aplicación mediante el [Portal de administración de Azure](https://manage.windowsazure.com). Inicie sesión con la cuenta de un administrador o con su suscripción a Office 365. Para saber cómo registrar aplicaciones, consulte [Registrar una aplicación con la Plataforma de identidad de Microsoft para desarrolladores](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually). Utilice la siguiente configuración:

 - URI REDIRCT: https://localhost:44301/AzureADAuth/Authorize	
 - TIPOS DE CUENTA ADMITIDAS: «Solo las cuentas de este directorio organizativo»
 - CONCESIÓN IMPLÍCITA: No habilitar las opciones implícitas de concesión
 - PERMISOS DE LA API: **Files.Read.All** y **User.Read**

	> Nota: Después de registrar la aplicación, copie la Id. de la aplicación (cliente) y elId. del directorio (inquilino) en la hoja de **información general** del registro de la aplicación en el Portal de administración de Azure. Cuando cree el secreto de cliente en la hoja de **Certificados y Secretos**, cópielo. 
	 
4.  En web.config, use los valores que copió en el paso anterior. Establezca **AAD:ClientID** en el Id. de cliente, **AAD:ClientSecret** en el secreto de cliente, y finalmente **"AAD:O365TenantID"** en el Id. de inquilino  

## Ejecutar el proyecto
1. Abra el archivo de solución de Visual Studio. 
2. Haga clic con el botón derecho en **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** y, después, seleccione **Establecer como proyecto de inicio**.
2. Pulse F5. 
3. En PowerPoint, abre la pestaña **Insertar** y luego seleccione **Elegir un gráfico** para abrir el complemento del panel de tareas.

## Problemas conocidos

* Escenario: Al intentar ejecutar el código de ejemplo, el complemento no se cargará.
	* Resolución: 
		1. En Visual Studio, abra **Explorador de objetos de SQL Server**.
		2. Expanda **(localdb)\\MSSQLLocalDB** > **Bases de datos**.
		3. Haga clic con el botón derecho en **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** y, después, elija **Eliminar**. 
* Escenario: Cuando ejecuta el ejemplo de código, obtiene un error en la línea *Office.context.ui.messageParent*.	
	* Resolución: Deje de ejecutar el ejemplo de código y reinícielo. 
* Si descarga el archivo zip, al extraer los archivos obtiene un error que indica que la ruta del archivo es demasiado larga.
	* Solución: Descomprima los archivos en una carpeta directamente bajo la raíz (por ejemplo, c:\\sample).

## Preguntas y comentarios
Nos encantaría recibir sus comentarios sobre el ejemplo de *Insertar gráficos de Excel con Microsoft Graph en un complemento de PowerPoint*. Usted puede enviarnos sus comentarios a través de la sección *Problemas* de este repositorio. Las preguntas generales sobre desarrollo en Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Asegúrese de que sus preguntas están etiquetadas con \[office-js], \[MicrosoftGraph] y \[API].

## Recursos adicionales

* [Ejemplo de código de lista de tarea pendiente de Microsoft Graph (Excel)](https://github.com/microsoftgraph/aspnet-todo-rest-sample)
* [Documentación de Microsoft Graph](https://docs.microsoft.com/en-us/graph/)
* [Documentación de complementos de Office](https://docs.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)

## Derechos de autor
Copyright (c) 2016 - 2019 Microsoft Corporation. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
