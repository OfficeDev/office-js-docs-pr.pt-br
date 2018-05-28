---
title: Criar e depurar Suplementos do Office no Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="c4168-102">Criar e depurar Suplementos do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="c4168-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="c4168-p101">Esse artigo descreve como usar o Visual Studio para criar o seu primeiro suplemento do Office. As etapas desse artigo t?m como base o Visual Studio 2015. Se voc? estiver usando outra vers?o do Visual Studio, os procedimentos poder?o variar um pouco.</span><span class="sxs-lookup"><span data-stu-id="c4168-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="c4168-106">Para come?ar a usar um suplemento do OneNote, confira o artigo [Crie seu primeiro suplemento do OneNote](../onenote/onenote-add-ins-getting-started.md).</span><span class="sxs-lookup"><span data-stu-id="c4168-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="c4168-107">Criar um projeto de suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="c4168-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="c4168-p102">Para come?ar, verifique se voc? tem as [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) instaladas e uma vers?o do Microsoft Office. ? poss?vel ingressar no [Programa do Desenvolvedor do Office 365](https://developer.microsoft.com/en-us/office/dev-program) ou seguir estas instru??es para receber a [?ltima vers?o](../develop/install-latest-office-version.md).</span><span class="sxs-lookup"><span data-stu-id="c4168-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/en-us/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>


1. <span data-ttu-id="c4168-110">Na barra de menus do Visual Studio, selecione **Arquivo**  >  **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="c4168-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="c4168-111">Na lista de tipos de projeto em **Visual C#** ou **Visual Basic**, expanda **Office/SharePoint**, escolha **Suplementos Web** e selecione um dos projetos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>  
    
3. <span data-ttu-id="c4168-112">Nomeie o projeto e escolha **OK** para cri?-lo.</span><span class="sxs-lookup"><span data-stu-id="c4168-112">Name the project, and then choose  **OK** to create the project.</span></span>
    
4. <span data-ttu-id="c4168-p103">O Visual Studio cria uma solu??o, e os dois projetos dele s?o exibidos no **Gerenciador de Solu??es**. A p?gina padr?o Home.html ? exibida no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="c4168-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.</span></span>
    
<span data-ttu-id="c4168-115">No Visual Studio 2015, alguns dos modelos de projetos de suplementos foram atualizados para refletir a funcionalidade adicional:</span><span class="sxs-lookup"><span data-stu-id="c4168-115">In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:</span></span>


- <span data-ttu-id="c4168-p104">Os suplementos de conte?do podem aparecer no corpo de documentos do Access e do PowerPoint, e em planilhas do Excel. Voc? tamb?m pode escolher a op??o Projeto B?sico para criar um projeto de suplemento de conte?do b?sico com c?digo inicial m?nimo, ou a op??o Projeto de Visualiza??o de Documento (apenas para Access e Excel) para criar um suplemento de conte?do mais completo que inclui c?digo inicial para visualizar e associar dados.</span><span class="sxs-lookup"><span data-stu-id="c4168-p104">Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.</span></span>
    
- <span data-ttu-id="c4168-118">Os suplementos do Outlook incluem op??es para incluir o suplemento em mensagens de email ou compromissos e para especificar se o suplemento est? dispon?vel quando uma mensagem de email ou um compromisso est? sendo redigido ou lido.</span><span class="sxs-lookup"><span data-stu-id="c4168-118">Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.</span></span>
    

> [!NOTE]
> <span data-ttu-id="c4168-p105">No Visual Studio, a maioria das op??es pode ser compreendida por meio das pr?prias descri??es, exceto a caixa de sele??o **Mensagem de Email**. Use essa caixa de sele??o se quiser criar um suplemento do Outlook exibido em itens de email e em solicita??es, respostas e cancelamentos de reuni?o.</span><span class="sxs-lookup"><span data-stu-id="c4168-p105">In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.</span></span>

<span data-ttu-id="c4168-121">Ao concluir o assistente, o Visual Studio cria uma solu??o que cont?m dois projetos.</span><span class="sxs-lookup"><span data-stu-id="c4168-121">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span>



|<span data-ttu-id="c4168-122">**Projeto**</span><span class="sxs-lookup"><span data-stu-id="c4168-122">**Project**</span></span>|<span data-ttu-id="c4168-123">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="c4168-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c4168-124">Projeto de suplemento</span><span class="sxs-lookup"><span data-stu-id="c4168-124">Add-in project</span></span>|<span data-ttu-id="c4168-p106">Cont?m somente um arquivo de manifesto XML, que cont?m todas as configura??es que descrevem o suplemento. As configura??es ajudam o host do Office a determinar quando o suplemento dever? ser ativado e onde ele dever? aparecer. O Visual Studio gera o conte?do desse arquivo para que voc? possa executar o projeto e usar o suplemento imediatamente . Voc? pode alterar as configura??es a qualquer momento usando o editor de Manifesto.</span><span class="sxs-lookup"><span data-stu-id="c4168-p106">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="c4168-129">Projeto de aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="c4168-129">Web application project</span></span>|<span data-ttu-id="c4168-p107">Cont?m as p?ginas de conte?do do suplemento, incluindo todos os arquivos e refer?ncias de arquivo de que voc? precisa para desenvolver p?ginas HTML e JavaScript com reconhecimento do Office. Enquanto voc? desenvolve o suplemento, o Visual Studio hospeda o aplicativo Web no servidor IIS local. Quando estiver pronto para publicar, voc? ter? de localizar um servidor para hospedar o projeto. Para saber mais sobre projetos de aplicativos Web ASP.NET, confira [Projetos Web ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="c4168-p107">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="c4168-133">Modificar as configura??es de suplemento</span><span class="sxs-lookup"><span data-stu-id="c4168-133">Modify your add-in settings</span></span>


<span data-ttu-id="c4168-p108">Para alterar as configura??es do seu suplemento, edite o arquivo de manifesto XML do projeto. No **Gerenciador de Solu??es**, expanda o n? de projeto do suplemento, expanda a pasta que cont?m o manifesto XML e escolha o manifesto XML. Voc? pode apontar para qualquer elemento do arquivo para exibir uma dica de ferramenta que descreve a finalidade do elemento. Para saber mais sobre o arquivo de manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="c4168-p108">To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="c4168-138">Desenvolver o conte?do do suplemento</span><span class="sxs-lookup"><span data-stu-id="c4168-138">Develop the contents of your add-in</span></span>


<span data-ttu-id="c4168-139">Enquanto o projeto de suplemento permite modificar as configura??es que descrevem o suplemento, o aplicativo Web fornece o conte?do que aparece no suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-139">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="c4168-p109">O projeto de aplicativo Web cont?m uma p?gina HTML padr?o e o arquivo JavaScript que voc? pode usar para come?ar. O projeto tamb?m cont?m um arquivo JavaScript que ? comum a todas as p?ginas que voc? adiciona ao projeto. Esses arquivos s?o convenientes porque cont?m refer?ncias a outras bibliotecas JavaScript, incluindo a API JavaScript para Office.</span><span class="sxs-lookup"><span data-stu-id="c4168-p109">The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> 

<span data-ttu-id="c4168-p110">? medida que o suplemento se tornar mais sofisticado, voc? poder? adicionar mais arquivos HTML e JavaScript. Voc? pode usar o conte?do dos arquivos HTML e JavaScript padr?o como exemplos dos tipos de refer?ncias que talvez queira adicionar a outras p?ginas no projeto para faz?-las funcionar com o suplemento. A tabela a seguir descreve os arquivos HTML e JavaScript padr?o.</span><span class="sxs-lookup"><span data-stu-id="c4168-p110">As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.</span></span>



|<span data-ttu-id="c4168-146">**Arquivo**</span><span class="sxs-lookup"><span data-stu-id="c4168-146">**File**</span></span>|<span data-ttu-id="c4168-147">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="c4168-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c4168-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="c4168-148">**Home.html**</span></span>|<span data-ttu-id="c4168-p111">Localizado na pasta **Home** do projeto, essa ? a p?gina HTML padr?o do suplemento. Essa p?gina ? exibida como a primeira no suplemento quando ele ? ativado em um documento, mensagem de email ou item de compromisso. Esse arquivo ? conveniente porque cont?m todas as refer?ncias de arquivo de que voc? precisa para come?ar. Quando estiver pronto para criar seu primeiro suplemento, basta adicionar o c?digo HTML a esse arquivo.</span><span class="sxs-lookup"><span data-stu-id="c4168-p111">Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.</span></span>|
|<span data-ttu-id="c4168-153">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="c4168-153">**Home.js**</span></span>|<span data-ttu-id="c4168-p112">Localizado na pasta **Home** do projeto, esse ? o arquivo JavaScript associado ? p?gina Home.js. Voc? pode colocar qualquer c?digo que seja espec?fico ao comportamento da p?gina Home.html no arquivo Home.js. O arquivo Home.js cont?m c?digo de exemplo para voc? come?ar.</span><span class="sxs-lookup"><span data-stu-id="c4168-p112">Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="c4168-157">**App.js**</span><span class="sxs-lookup"><span data-stu-id="c4168-157">**App.js**</span></span>|<span data-ttu-id="c4168-p113">Localizado na pasta **Add-in** do projeto, esse ? o arquivo JavaScript padr?o do suplemento inteiro. Voc? pode colocar c?digo comum ao comportamento de v?rias p?ginas do suplemento no arquivo App.js. O arquivo App.js cont?m c?digo de exemplo para voc? come?ar.</span><span class="sxs-lookup"><span data-stu-id="c4168-p113">Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.</span></span>|

> [!NOTE]
> <span data-ttu-id="c4168-p114">N?o ? necess?rio usar esses arquivos. Fique ? vontade para adicionar outros arquivos ao projeto e us?-los. Se desejar que outro arquivo HTML apare?a como a p?gina inicial do suplemento, abra o editor de manifesto e aponte a propriedade **SourceLocation** para o nome do arquivo.</span><span class="sxs-lookup"><span data-stu-id="c4168-p114">You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>


## <a name="debug-your-add-in"></a><span data-ttu-id="c4168-164">Depurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="c4168-164">Debug your add-in</span></span>


<span data-ttu-id="c4168-165">Quando estiver pronto para iniciar o suplemento, examine as propriedades relacionadas ? compila??o e ? depura??o e inicie a solu??o.</span><span class="sxs-lookup"><span data-stu-id="c4168-165">When you are ready to start your add-in, review build and debug related properties, and then start the solution.</span></span>


### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="c4168-166">Examinar as propriedades de compila??o e depura??o</span><span class="sxs-lookup"><span data-stu-id="c4168-166">Review the build and debug properties</span></span>

<span data-ttu-id="c4168-p115">Antes de iniciar a solu??o, verifique se o Visual Studio abrir? o aplicativo host desejado. Essa informa??o ? exibida nas p?ginas de propriedades do projeto, com v?rias outras propriedades relacionadas ? compila??o e ? depura??o do suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-p115">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>


### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="c4168-169">Para abrir as p?ginas de propriedades de um projeto</span><span class="sxs-lookup"><span data-stu-id="c4168-169">To open the property pages of a project</span></span>


1. <span data-ttu-id="c4168-170">No **Gerenciador de Solu??es**, escolha o nome do projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-170">In  **Solution Explorer**, choose the project name.</span></span>
    
2. <span data-ttu-id="c4168-171">Na barra de menus, escolha **Exibir**, **Janela Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="c4168-171">On the menu bar, choose  **View**,  **Properties Window**.</span></span>
    
<span data-ttu-id="c4168-172">A tabela a seguir descreve as propriedades do projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-172">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="c4168-173">**Propriedade**</span><span class="sxs-lookup"><span data-stu-id="c4168-173">**Property**</span></span>|<span data-ttu-id="c4168-174">**Descri??o**</span><span class="sxs-lookup"><span data-stu-id="c4168-174">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="c4168-175">**Iniciar A??o**</span><span class="sxs-lookup"><span data-stu-id="c4168-175">**Start Action**</span></span>|<span data-ttu-id="c4168-176">Especifica se o suplemento deve ser depurado em um cliente da ?rea de trabalho do Office ou em um cliente do Office Online no navegador especificado.</span><span class="sxs-lookup"><span data-stu-id="c4168-176">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="c4168-177">**Iniciar Documento** (apenas suplementos de conte?do e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="c4168-177">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="c4168-178">Especifica o documento a ser aberto quando voc? iniciar o projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-178">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="c4168-179">**Projeto da Web**</span><span class="sxs-lookup"><span data-stu-id="c4168-179">**Web Project**</span></span>|<span data-ttu-id="c4168-180">Especifica o nome do projeto Web associado ao suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-180">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="c4168-181">**Endere?o de Email** (apenas suplementos do Outlook)</span><span class="sxs-lookup"><span data-stu-id="c4168-181">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="c4168-182">Especifica o endere?o de email da conta de usu?rio no Exchange Server ou no Exchange Online com a qual voc? deseja testar o suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c4168-182">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="c4168-183">**Url EWS** (apenas suplementos do Outlook)</span><span class="sxs-lookup"><span data-stu-id="c4168-183">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="c4168-184">URL do servi?o Web do Exchange (por exemplo: https://www.contoso.com/ews/exchange.aspx).</span><span class="sxs-lookup"><span data-stu-id="c4168-184">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="c4168-185">**Url OWA** (apenas suplementos do Outlook)</span><span class="sxs-lookup"><span data-stu-id="c4168-185">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="c4168-186">URL do Outlook Web App (Por exemplo: https://www.contoso.com/owa).</span><span class="sxs-lookup"><span data-stu-id="c4168-186">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="c4168-187">**Nome de usu?rio** (apenas suplementos do Outlook)</span><span class="sxs-lookup"><span data-stu-id="c4168-187">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="c4168-188">Especifica o nome de sua conta de usu?rio no Exchange Server ou no Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="c4168-188">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="c4168-189">**Arquivo do projeto**</span><span class="sxs-lookup"><span data-stu-id="c4168-189">**Project File**</span></span>|<span data-ttu-id="c4168-190">Especifica o nome do arquivo que cont?m informa??es de compila??o, configura??o e outras informa??es sobre o projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-190">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="c4168-191">**Pasta do projeto**</span><span class="sxs-lookup"><span data-stu-id="c4168-191">**Project Folder**</span></span>|<span data-ttu-id="c4168-192">O local do arquivo do projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-192">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="c4168-193">Use um documento existente para depurar o suplemento (apenas suplementos de conte?do e de painel de tarefas)</span><span class="sxs-lookup"><span data-stu-id="c4168-193">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>


<span data-ttu-id="c4168-p116">Voc? pode adicionar documentos ao projeto de suplemento. Se voc? tiver um documento que contenha os dados de teste que deseja usar com o suplemento, o Visual Studio abrir? esse documento quando voc? iniciar o projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-p116">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="c4168-196">Para usar um documento existente para depurar o suplemento</span><span class="sxs-lookup"><span data-stu-id="c4168-196">To use an existing document to debug the add-in</span></span>


1. <span data-ttu-id="c4168-197">No **Gerenciador de Solu??es**, escolha a pasta do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-197">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="c4168-198">Escolha o projeto do suplemento, n?o o projeto do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="c4168-198">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="c4168-199">No menu **Projeto**, escolha **Adicionar Item Existente**.</span><span class="sxs-lookup"><span data-stu-id="c4168-199">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="c4168-200">Na caixa de di?logo **Adicionar Item Existente**, localize e selecione o documento que voc? deseja adicionar.</span><span class="sxs-lookup"><span data-stu-id="c4168-200">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="c4168-201">Escolha o bot?o **Adicionar** para adicionar o documento ao projeto.</span><span class="sxs-lookup"><span data-stu-id="c4168-201">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="c4168-202">No **Gerenciador de Solu??es**, abra o menu de atalho do projeto e escolha  **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="c4168-202">In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.</span></span>
    
    <span data-ttu-id="c4168-203">As p?ginas de propriedades do projeto s?o exibidas.</span><span class="sxs-lookup"><span data-stu-id="c4168-203">The property pages for the project appear.</span></span>
    
6. <span data-ttu-id="c4168-204">Na lista **Iniciar Documento**, escolha o documento que voc? adicionou ao projeto e escolha o bot?o **OK** para fechar as p?ginas de propriedades.</span><span class="sxs-lookup"><span data-stu-id="c4168-204">In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.</span></span>
    

### <a name="start-the-solution"></a><span data-ttu-id="c4168-205">Iniciar a solu??o</span><span class="sxs-lookup"><span data-stu-id="c4168-205">Start the solution</span></span>


<span data-ttu-id="c4168-p117">O Visual Studio compilar? automaticamente a solu??o ao iniciar. Voc? pode iniciar a solu??o por meio da barra de **Menus** escolhendo **Depurar**, **Iniciar**.</span><span class="sxs-lookup"><span data-stu-id="c4168-p117">Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**.</span></span> 


> [!NOTE]
> <span data-ttu-id="c4168-p118">Se a depura??o de script n?o estiver habilitada no Internet Explorer, voc? n?o poder? iniciar o depurador no Visual Studio. ? poss?vel habilitar a depura??o de scripts abrindo a caixa de di?logo **Op??es da Internet**, escolhendo a guia **Avan?ado** e desmarcando as caixas de sele??o **Desabilitar depura??o de script (Internet Explorer)** e **Desabilitar a depura??o de script (outros)**.</span><span class="sxs-lookup"><span data-stu-id="c4168-p118">If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.</span></span>

<span data-ttu-id="c4168-210">O Visual Studio compila o projeto e faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="c4168-210">Visual Studio builds the project and does the following:</span></span>


1. <span data-ttu-id="c4168-p119">Cria uma c?pia do arquivo de manifesto XML e a adiciona ao diret?rio _NomedoProjeto_\Output. O aplicativo host consome esta c?pia quando voc? inicia o Visual Studio e depura o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-p119">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="c4168-213">Cria um conjunto de entradas de registro no computador que permitem que o suplemento seja exibido no aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="c4168-213">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="c4168-214">Compila o projeto de aplicativo da Web e o implanta no servidor Web IIS local (http://localhost).</span><span class="sxs-lookup"><span data-stu-id="c4168-214">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="c4168-215">Depois, o Visual Studio faz o seguinte:</span><span class="sxs-lookup"><span data-stu-id="c4168-215">Next, Visual Studio does the following:</span></span>


1. <span data-ttu-id="c4168-216">Modifica o elemento [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) do arquivo de manifesto XML, substituindo o token ~ remoteAppUrl pelo endere?o totalmente qualificado da p?gina inicial (por exemplo, http://localhost/MyAgave.html).</span><span class="sxs-lookup"><span data-stu-id="c4168-216">Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="c4168-217">Inicia o projeto de aplicativo da Web no IIS Express.</span><span class="sxs-lookup"><span data-stu-id="c4168-217">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="c4168-218">Abre o aplicativo host.</span><span class="sxs-lookup"><span data-stu-id="c4168-218">Opens the host application.</span></span> 
    
<span data-ttu-id="c4168-p120">O Visual Studio n?o mostra erros de valida??o na janela **OUTPUT** ao compilar o projeto. O Visual Studio relata erros e avisos na janela **ERRORLIST** ? medida que eles ocorrem. O Visual Studio tamb?m relata erros de valida??o mostrando sublinhados ondulados (conhecidos como rabiscos) de cores diferentes no editor de c?digo e texto. Essas marcas o notificam de problemas que o Visual Studio detectou no c?digo. Para saber mais, confira [Editor de c?digo e texto](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Para saber mais sobre como habilitar ou desabilitar a valida??o, confira:</span><span class="sxs-lookup"><span data-stu-id="c4168-p120">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- <span data-ttu-id="c4168-225">[Op??es, editor de texto, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="c4168-225">[Options, Text Editor, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span></span>
    
- <span data-ttu-id="c4168-226">[Tutorial: Definir op??es de valida??o para edi??o de HTML no Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="c4168-226">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="c4168-227">[CSS, confira Valida??o, CSS, editor de texto, caixa de di?logo Op??es](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="c4168-227">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="c4168-228">Para examinar as regras de valida??o do arquivo de manifesto XML no projeto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="c4168-228">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a><span data-ttu-id="c4168-229">Mostrar um suplemento no Excel, no Word ou no Project e percorrer o c?digo</span><span class="sxs-lookup"><span data-stu-id="c4168-229">Show an add-in in Excel, Word, or Project and step through your code</span></span>


<span data-ttu-id="c4168-p121">Se voc? definir a propriedade **Start Document** do projeto de suplemento para o Excel ou o Word, o Visual Studio criar? um novo documento e o suplemento ser? exibido. Se voc? definir a propriedade **Start Document** do projeto de suplemento para usar um documento existente, o Visual Studio abrir? o documento, mas voc? precisar? inserir manualmente o suplemento. Se definir **Start Document** como **Microsoft Project**, voc? precisar? inserir manualmente o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-p121">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.</span></span>


### <a name="to-show-an-office-add-in-in-excel-or-word"></a><span data-ttu-id="c4168-233">Para mostrar um suplemento do Office no Excel ou no Word</span><span class="sxs-lookup"><span data-stu-id="c4168-233">To show an Office Add-in in Excel or Word</span></span>


1. <span data-ttu-id="c4168-234">No Excel ou no Word, na guia **Inserir**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="c4168-234">In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="c4168-235">Na lista exibida, escolha o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-235">In the list that appears, choose your add-in.</span></span>
    

### <a name="to-show-an-office-add-in-in-project"></a><span data-ttu-id="c4168-236">Para mostrar um suplemento do Office no Project</span><span class="sxs-lookup"><span data-stu-id="c4168-236">To show an Office Add-in in Project</span></span>


1. <span data-ttu-id="c4168-237">No Project, na guia **Projeto**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="c4168-237">In Project, on the  **Project** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="c4168-238">Na lista exibida, escolha o suplemento.</span><span class="sxs-lookup"><span data-stu-id="c4168-238">In the list that appears, choose your add-in.</span></span>
    
<span data-ttu-id="c4168-p122">No Visual Studio, voc? pode ent?o definir pontos de interrup??o. Depois, voc? interage com o suplemento e percorre o c?digo nos arquivos de c?digo HTML, JavaScript e C# ou VB.</span><span class="sxs-lookup"><span data-stu-id="c4168-p122">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="c4168-241">Mostrar o suplemento do Outlook no Outlook e percorrer o c?digo</span><span class="sxs-lookup"><span data-stu-id="c4168-241">Show the Outlook add-in in Outlook and step through your code</span></span>


<span data-ttu-id="c4168-242">Para exibir o suplemento no Outlook, abra uma mensagem de email ou um item de compromisso.</span><span class="sxs-lookup"><span data-stu-id="c4168-242">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="c4168-p123">O Outlook ativa o suplemento para o item, contanto que os crit?rios de ativa??o sejam atendidos. A barra de suplementos aparece na parte superior da janela Inspetor ou Painel de Leitura, e o suplemento do Outlook aparece como um bot?o na barra de suplementos. Se o suplemento tiver um comando de suplemento, aparecer? um bot?o na faixa de op??es, na guia padr?o ou em uma guia personalizada especificada, e o suplemento n?o aparecer? na barra de suplementos.</span><span class="sxs-lookup"><span data-stu-id="c4168-p123">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="c4168-246">Para exibir o suplemento do Outlook, escolha o bot?o do suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="c4168-246">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="c4168-p124">No Visual Studio, voc? pode definir pontos de interrup??o. Depois, voc? interage com o suplemento do Outlook e percorre o c?digo nos arquivos de c?digo HTML, JavaScript e C# ou VB.</span><span class="sxs-lookup"><span data-stu-id="c4168-p124">In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span> 

<span data-ttu-id="c4168-p125">Voc? tamb?m pode alterar o c?digo e examinar os efeitos das altera??es no suplemento do Outlook sem ter que fechar o Suplemento do Office e reiniciar o projeto. No Outlook, basta abrir o menu de atalho do suplemento do Outlook e escolher **Recarregar**.</span><span class="sxs-lookup"><span data-stu-id="c4168-p125">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="c4168-251">Modificar o c?digo e continuar a depurar o suplemento sem ter que reiniciar o projeto</span><span class="sxs-lookup"><span data-stu-id="c4168-251">Modify code and continue to debug the add-in without having to start the project again</span></span>


<span data-ttu-id="c4168-p126">Voc? pode alterar o c?digo e examinar os efeitos das altera??es no suplemento sem ter que fechar o aplicativo host e reiniciar o projeto. Ap?s alterar o c?digo, abra o menu de atalho do suplemento e escolha **Recarregar**. Quando voc? recarregar o suplemento, ele ? desconectado do depurador do Visual Studio. Portanto, voc? pode exibir os efeitos da altera??o, mas n?o pode percorrer o c?digo novamente at? anexar o depurador do Visual Studio a todos os processos Iexplore.exe dispon?veis.</span><span class="sxs-lookup"><span data-stu-id="c4168-p126">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.</span></span>


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a><span data-ttu-id="c4168-256">Para anexar o depurador do Visual Studio a todos os processos Iexplore.exe dispon?veis</span><span class="sxs-lookup"><span data-stu-id="c4168-256">To attach the Visual Studio debugger to all of the available Iexplore.exe processes</span></span>


1. <span data-ttu-id="c4168-257">No Visual Studio, escolha **DEPURAR**, **Anexar ao Processo**.</span><span class="sxs-lookup"><span data-stu-id="c4168-257">In Visual Studio, choose  **DEBUG**,  **Attach to Process**.</span></span>
    
2. <span data-ttu-id="c4168-258">Na caixa de di?logo **Anexar ao Processo**, escolha todos os processos **Iexplore.exe** dispon?veis e, em seguida, selecione o bot?o **Anexar**.</span><span class="sxs-lookup"><span data-stu-id="c4168-258">In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="c4168-259">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="c4168-259">Next steps</span></span>

- [<span data-ttu-id="c4168-260">Implantar e publicar seu suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="c4168-260">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
