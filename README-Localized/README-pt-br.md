# Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint 

Saiba como criar um Suplemento do Microsoft Office que se conecta ao Microsoft Graph, localiza todas as pastas de trabalho armazenadas no OneDrive for Business, busca todos os gráficos nas pastas de trabalho usando as APIs REST do Excel e insere a imagem de um gráfico em um slide do PowerPoint usando Office.js.

![Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint](../images/InsertChart.png)

## Introdução

A integração de dados de provedores de serviço online aumenta o valor e a adoção de seus suplementos. O código a seguir mostra como conectar seu suplemento ao Microsoft Graph. Use este exemplo de código para:

* Conectar-se ao Microsoft Graph a partir de um Suplemento do Office.
* Use a estrutura de autorização OAuth 2.0 em um suplemento.
* Use o Excel e as APIs REST do OneDrive a partir do Microsoft Graph.
* Exiba uma caixa de diálogo usando o namespace da interface do usuário do Office.
* Crie um Suplemento usando ASP.NET MVC e Office.js. 
* Use comandos de suplemento no PowerPoint.


## Pré-requisitos
Para executar este exemplo de código, são necessários.

* Visual Studio 2015.

* Uma conta do Office 365 que pode ser obtida ingressando no <a herf="https://aka.ms/devprogramsignup">Programa do Desenvolvedor do Office 365</a> que inclui uma assinatura gratuita de 1 ano para o Office 365.

* Pastas de trabalho (com gráficos) do Excel armazenadas no OneDrive for Business em sua assinatura do Office 365.

* PowerPoint para Windows Desktop, versão 16.0.6769.2001 ou superior.
* [Office Developer Tools](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)

* Um Locatário do Microsoft Azure. Esse suplemento requer o Microsoft Azure Active Directory (AD). O Azure AD fornece serviços de identidade que os aplicativos usam para autenticação e autorização. Você pode adquirir uma assinatura de avaliação aqui: [Microsoft Azure](https://account.windowsazure.com/SignUp).

## Configurar o projeto

1. No **Visual Studio**, escolha o projeto **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChartWeb**. Em **Propriedades**, certifique-se de **SSL Habilitado** seja **True**. Verifique se a **URL do SSL** usa o mesmo nome de domínio e número de porta como esses listado na etapa 3 abaixo.
 
2. Assegure que sua assinatura do Azure esteja vinculada ao seu locatário do Office 365. Para obter mais informações, confira a postagem de blog da equipe do Active Directory: [Criar e Gerenciar Vários Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). A seção **Adicionar um novo diretório** explica como fazer isso. Para obter mais informações, também é possível consultar o artigo [Configurar seu ambiente de desenvolvimento do Office 365](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) e a seção **Associar sua conta do Office 365 ao Azure AD para criar e gerenciar aplicativos**.

3. Registre seu aplicativo usando o [Portal de Gerenciamento do Azure](https://manage.windowsazure.com). Para saber como registrar seu aplicativo, consulte [Registrar seu aplicativo Web baseado em navegador com o Portal de Gerenciamento do Azure](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp). Use as seguintes configurações:

 - URL DE ENTRADA: https://localhost:44301/AzureADAuth/Authorize 
 - URI DA ID DO APLIVATIVO: https://localhost:44301
 - URL DE RESPOSTA/localhost:44301/AzureADAuth/Authorize	

	> Observação: Depois que você registrar seu aplicativo, copie a id do cliente e o segredo do cliente exibidos no Portal de Gerenciamento do Azure.
	 
4. Conceda permissões para seu aplicativo.
	*  No Portal de Gerenciamento do Azure, selecione a guia **Active Directory** e um locatário do Office 365.
	*  Selecione a guia **Aplicativos** e clique no aplicativo que você deseja configurar. Escolha **Configurar**.
	*  Em **permissões para outros aplicativos**, adicione **Microsoft Graph**.
	*  Em **Permissões Delegadas**, escolha **Ler arquivos do usuário e arquivos compartilhados com o usuário**.

5.  Em web.config, defina **AAD:ClientID** como id do cliente e **AAD:ClientSecret** como seu segredo de cliente. 

## Executar o projeto
1. Abra o arquivo de solução do Visual Studio. 
2. Clique com o botão direito do mouse em **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** e escolha **Definir como Projeto de Inicialização**.
2. Pressione F5. 
3. No PowerPoint, escolha **Inserir** > **Escolher um gráfico** para abrir o suplemento do painel de tarefas.

## Problemas conhecidos

* Situação: Ao tentar executar o exemplo de código, o suplemento não será carregado.
	* Resolução: 
		1. No Visual Studio, abra o **Pesquisador de Objetos do SQL Server**.
		2. Expanda **(localdb) \MSSQLLocalDB** > **Bancos de Dados**.
		3. Clique com o botão direito do mouse em **PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart** e escolha **Excluir**. 
* Situação: Quando você executa o código de exemplo, você recebe um erro na linha *Office.context.ui.messageParent*.	
	* Resolução: Interrompa a execução do código de exemplo e reinicie-o. 
* Se baixar o arquivo zip, ao extrair os arquivos, você receberá um erro indicando que o caminho do arquivo é muito longo.
	* Resolução: Descompacte os arquivos em uma pasta diretamente na raiz (por exemplo, c:\exemplo).

## Perguntas e comentários
Gostaríamos de receber seu comentário sobre o exemplo de *Inserir gráficos do Excel usando o Microsoft Graph em um Suplemento do PowerPoint*. Você pode enviar comentários na seção *Problemas* deste repositório. As perguntas sobre o desenvolvimento do Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/Office365+API). Verifique se suas perguntas são marcadas com [office-js], [MicrosoftGraph] e [API].

## Recursos adicionais

* [Exemplo de código de Tarefas Pendentes do Microsoft Graph (Excel)](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-ExcelREST-ToDo)
* [Documentação do Microsoft Graph](https://graph.microsoft.io/en-us/docs)
* [Documentação dos Suplementos do Office](https://dev.office.com/docs/add-ins/overview/office-add-ins)
* Confira o vídeo em //Compilar - [Visão Geral da Plataforma do Office](https://channel9.msdn.com/Events/Build/2016/B872 "Visão Geral da Plataforma do Office").

## Direitos autorais
Copyright (C) 2016 Microsoft Corporation. Todos os direitos reservados.


