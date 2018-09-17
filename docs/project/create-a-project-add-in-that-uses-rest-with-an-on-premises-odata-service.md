---
title: Criar um suplemento de Project que usa REST com um serviço OData local do Project Server
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 462b151a12c7c19f445e86bf59af633b699e7e60
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945705"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="8392a-102">Criar um suplemento de Project que usa REST com um serviço OData local do Project Server</span><span class="sxs-lookup"><span data-stu-id="8392a-102">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="8392a-p101">Este artigo descreve como criar um suplemento de painel tarefas do Project Professional 2013 que compara dados de custo e de trabalho no projeto ativo com médias de todos os projetos da instância atual do Project Web App. O suplemento usa REST com a biblioteca jQuery para acessar o serviço de relatório OData **ProjectData** no Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="8392a-p101">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance. The add-in uses REST with the jQuery library to access the  **ProjectData** OData reporting service in Project Server 2013.</span></span>


<span data-ttu-id="8392a-105">O código deste artigo é baseado em um exemplo desenvolvido por Saurabh Sanghvi e Arvind Iyer, da Microsoft Corporation.</span><span class="sxs-lookup"><span data-stu-id="8392a-105">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="8392a-106">Pré-requisitos para a criação de um suplemento de painel de tarefas que lê dados de relatório do Project Server</span><span class="sxs-lookup"><span data-stu-id="8392a-106">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>


<span data-ttu-id="8392a-107">A seguir temos os pré-requisitos para a criação de um suplemento de painel de tarefas do Project que lê o serviço **ProjectData** de uma instância do Project Web App em uma instalação local do Project Server 2013:</span><span class="sxs-lookup"><span data-stu-id="8392a-107">The following are the prerequisites for creating a Project task pane add-in that reads the  **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013:</span></span>


- <span data-ttu-id="8392a-p102">Verifique se você instalou os service packs e as atualizações mais recentes do Windows em seu computador de desenvolvimento local. O sistema operacional pode ser Windows 7, Windows 8, Windows Server 2008 ou Windows Server 2012.</span><span class="sxs-lookup"><span data-stu-id="8392a-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>
    
- <span data-ttu-id="8392a-p103">O Project Professional 2013 é necessário para a conexão com o Project Web App. O computador de desenvolvimento deve ter o Project Professional 2013 instalado para habilitar a depuração **F5** com o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="8392a-p103">Project Professional 2013 is required to connect with Project Web App. The development computer must have Project Professional 2013 installed to enable  **F5** debugging with Visual Studio.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="8392a-112">O Project Standard 2013 também pode hospedar suplementos de painel de tarefas, mas não pode fazer logon no Project Web App.</span><span class="sxs-lookup"><span data-stu-id="8392a-112">Project Standard 2013 can also host task pane add-ins, but cannot log on to Project Web App.</span></span>

- <span data-ttu-id="8392a-113">O Visual Studio 2015 com Office Developer Tools para Visual Studio inclui modelos para criar suplementos do Office e do SharePoint. Verifique se você instalou a versão mais recente do Office Developer Tools. Confira a seção _Ferramentas_ de [Download de suplementos do Office e do SharePoint](https://developer.microsoft.com/office/docs).</span><span class="sxs-lookup"><span data-stu-id="8392a-113">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](https://developer.microsoft.com/office/docs).</span></span>
    
- <span data-ttu-id="8392a-p104">Os procedimentos e exemplos de código neste artigo acessam o serviço **ProjectData** do Project Server 2013 em um domínio local. Os métodos jQuery neste artigo não funcionam com o Project Online.</span><span class="sxs-lookup"><span data-stu-id="8392a-p104">The procedures and code examples in this article access the  **ProjectData** service of Project Server 2013 in a local domain. The jQuery methods in this article do not work with Project Online.</span></span>
    
    <span data-ttu-id="8392a-116">Verifique se o serviço **ProjectData** está acessível do seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="8392a-116">Verify that the  **ProjectData** service is accessible from your development computer.</span></span>
    

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="8392a-p105">Procedimento 1. Para verificar se o serviço ProjectData está acessível</span><span class="sxs-lookup"><span data-stu-id="8392a-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>


1. <span data-ttu-id="8392a-p106">Para permitir que seu navegador mostre os dados XML de consultas REST diretamente, desative o modo de exibição de leitura de feed. Para saber mais sobre como fazer isso no Internet Explorer, confira o Procedimento 1, etapa 4 em [Consultar feeds OData para dados de relatório do Project](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="8392a-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>
    
2. <span data-ttu-id="8392a-121">Consulte o serviço  **ProjectData** usando seu navegador com a seguinte URL: **http://ServerName /ProjectServerName /_api/ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="8392a-121">Query the  ProjectData service by using your browser with the following URL: http://ServerName /ProjectServerName /_api/ProjectData. For example, if the Project Web App instance is  , the browser shows the following results:</span></span> <span data-ttu-id="8392a-122">Por exemplo, se a instância de aplicativo Web do Project for `http://MyServer/pwa`, o navegador mostrará os seguintes resultados:</span><span class="sxs-lookup"><span data-stu-id="8392a-122">Query the  ProjectData service by using your browser with the following URL: http://ServerName /ProjectServerName /_api/ProjectData. For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results:</span></span>
    
    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/" 
        xmlns="http://www.w3.org/2007/app" 
        xmlns:atom="http://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

3. <span data-ttu-id="8392a-p108">Pode ser necessário fornecer as credenciais de rede para ver os resultados. Se o navegador exibir "Erro 403, acesso negado", você não tem permissão de logon para essa instância do Project Web App ou há algum problema de rede que exige ajuda administrativa.</span><span class="sxs-lookup"><span data-stu-id="8392a-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>
    

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="8392a-125">Usar o Visual Studio para criar um suplemento de painel de tarefas para o Project</span><span class="sxs-lookup"><span data-stu-id="8392a-125">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="8392a-p109">O Office Developer Tools para Visual Studio inclui um modelo de suplemento de painel de tarefas para o Project 2013. Se você criar uma solução denominada **HelloProjectOData**, ela conterá os dois projetos do Visual Studio a seguir:</span><span class="sxs-lookup"><span data-stu-id="8392a-p109">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013. If you create a solution named  **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>


- <span data-ttu-id="8392a-p110">O projeto de suplemento usa o nome da solução. Ele inclui o arquivo de manifesto XML para o suplemento e serve para o .NET Framework 4.5. O Procedimento 3 mostra as etapas para modificar o manifesto para o suplemento **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p110">The add-in project takes the name of the solution. It includes the XML manifest file for the add-in and targets the .NET Framework 4.5. Procedure 3 shows the steps to modify the manifest for the  **HelloProjectOData** add-in.</span></span>
    
- <span data-ttu-id="8392a-p111">O projeto Web é denominado **HelloProjectODataWeb**. Ele inclui as páginas da Web, os arquivos JavaScript, os arquivos CSS, as imagens, as referências e os arquivos de configuração para o conteúdo Web no painel de tarefas. O projeto Web serve para o .NET Framework 4. Os Procedimentos 4 e 5 mostram como modificar os arquivos no projeto Web para criar a funcionalidade do suplemento **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p111">The web project is named  **HelloProjectODataWeb**. It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane. The web project targets the .NET Framework 4. Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the  **HelloProjectOData** add-in.</span></span>
    

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="8392a-p112">Procedimento 2. Para criar o suplemento HelloProjectOData para o Project</span><span class="sxs-lookup"><span data-stu-id="8392a-p112">Procedure 2. To create the HelloProjectOData add-in for Project</span></span>


1. <span data-ttu-id="8392a-137">Execute o Visual Studio 2015 como administrador e selecione **Novo Projeto** na página Iniciar.</span><span class="sxs-lookup"><span data-stu-id="8392a-137">Run Visual Studio 2015 as an administrator, and then select  **New Project** on the Start page.</span></span>
    
2. <span data-ttu-id="8392a-p113">Na caixa de diálogo **Novo Projeto**, expanda os nós **Modelos**, **Visual C#** e **Office/SharePoint** e selecione **Suplementos do Office**. Selecione **.NET Framework 4.5.2** na lista suspensa de estrutura de destino na parte superior do painel central e, em seguida, selecione **Suplemento do Office** (veja a captura de tela a seguir).</span><span class="sxs-lookup"><span data-stu-id="8392a-p113">In the  **New Project** dialog box, expand the **Templates**,  **Visual C#**, and  **Office/SharePoint** nodes, and then select \*\* Office Add-ins\*\*. Select  **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>
    
3. <span data-ttu-id="8392a-140">Para colocar ambos os projetos do Visual Studio no mesmo diretório, selecione **Criar diretório para solução** e navegue até o local desejado.</span><span class="sxs-lookup"><span data-stu-id="8392a-140">To place both of the Visual Studio projects in the same directory, select  **Create directory for solution**, and then browse to the location you want.</span></span>
    
4. <span data-ttu-id="8392a-141">No campo **Nome**, digite HelloProjectOData e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="8392a-141">In the  **Name** field, typeHelloProjectOData, and then choose  **OK**.</span></span>
    
    <span data-ttu-id="8392a-142">*Figura 1. Criação de um suplemento do Office*</span><span class="sxs-lookup"><span data-stu-id="8392a-142">*Figure 1. Creating an Office Add-in*</span></span>

    ![Criar um Suplemento do Office](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="8392a-144">Na caixa de diálogo **Escolha o tipo de suplemento**, selecione **Painel de tarefas** e escolha **Avançar** (veja a captura de tela a seguir).</span><span class="sxs-lookup"><span data-stu-id="8392a-144">In the  **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>
    
    <span data-ttu-id="8392a-145">*Figura 2. Como escolher o tipo de suplemento a criar*</span><span class="sxs-lookup"><span data-stu-id="8392a-145">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![Escolher o tipo de suplemento a criar](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="8392a-147">Na caixa de diálogo **Escolha os aplicativos host**, desmarque todas as caixas de seleção, exceto o **Project** (veja a captura de tela a seguir) e escolha **Concluir**.</span><span class="sxs-lookup"><span data-stu-id="8392a-147">In the  **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>
    
    <span data-ttu-id="8392a-148">*Figura 3. Como escolher o aplicativo host*</span><span class="sxs-lookup"><span data-stu-id="8392a-148">*Figure 3. Choosing the host application*</span></span>

    ![Escolher o Project como o único aplicativo host](../images/create-office-add-in.png)
    
    <span data-ttu-id="8392a-150">O Visual Studio cria o projeto **HelloProjectOdata** e o projeto **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="8392a-150">Visual Studio creates the  **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>
    
<span data-ttu-id="8392a-p114">A pasta **AddIn** (veja a captura de tela a seguir) contém o arquivo App.css para estilos CSS personalizados. Na subpasta **Home**, o arquivo Home.html contém referências para arquivos CSS e JavaScript que o suplemento usa, e o conteúdo HTML5 para o suplemento. Além disso, o arquivo Home.js é para o seu código JavaScript personalizado. A pasta **Scripts** inclui os arquivos da biblioteca jQuery. A subpasta **Office** inclui as bibliotecas JavaScript, como office.js e project-15.js, além das bibliotecas de linguagem para sequências de caracteres padrão nos suplementos do Office. Na pasta **Content**, o arquivo Office.css contém os estilos padrão para todos os suplementos do Office.</span><span class="sxs-lookup"><span data-stu-id="8392a-p114">The  **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles. In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in. Also, the Home.js file is for your custom JavaScript code. The **Scripts** folder includes the jQuery library files. The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office add-ins.</span></span>

<span data-ttu-id="8392a-156">*Figura 4. Exibição de arquivos de projeto Web padrão no Gerenciador de Soluções*</span><span class="sxs-lookup"><span data-stu-id="8392a-156">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![Exibir os arquivos do projeto Web no Gerenciador de Soluções](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="8392a-p115">O manifesto para o projeto **HelloProjectOData** é o arquivo HelloProjectOData.xml. Opcionalmente, você pode modificar o manifesto para adicionar uma descrição do suplemento, uma referência a um ícone, informações de linguagem adicionais e outras configurações. O Procedimento 3 simplesmente modifica o nome de exibição e a descrição do suplemento e adiciona um ícone.</span><span class="sxs-lookup"><span data-stu-id="8392a-p115">The manifest for the  **HelloProjectOData** project is the HelloProjectOData.xml file. You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings. Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="8392a-161">Para saber mais sobre o manifesto, confira [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md) e [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md#see-also).</span><span class="sxs-lookup"><span data-stu-id="8392a-161">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="8392a-p116">Procedimento 3. Para modificar o manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="8392a-p116">Procedure 3. To modify the add-in manifest</span></span>


1. <span data-ttu-id="8392a-164">No Visual Studio, abra o arquivo HelloProjectOData.xml.</span><span class="sxs-lookup"><span data-stu-id="8392a-164">In Visual Studio, open the HelloProjectOData.xml file.</span></span>
    
2. <span data-ttu-id="8392a-p117">O nome de exibição padrão é o nome do projeto do Visual Studio ("HelloProjectOData"). Por exemplo, altere o valor padrão do elemento **DisplayName** para "Hello ProjectData".</span><span class="sxs-lookup"><span data-stu-id="8392a-p117">The default display name is the name of the Visual Studio project ("HelloProjectOData"). For example, change the default value of the  **DisplayName** element to"Hello ProjectData".</span></span>
    
3. <span data-ttu-id="8392a-p118">A descrição padrão também é "HelloProjectOData". Por exemplo, altere o valor padrão do elemento Description para "Testar consultas REST do serviço ProjectData".</span><span class="sxs-lookup"><span data-stu-id="8392a-p118">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>
    
4. <span data-ttu-id="8392a-p119">Adicione um ícone para mostrar a lista suspensa **Suplementos do Office** na guia **PROJETO** da faixa de opções. Você pode adicionar um arquivo de ícone na solução do Visual Studio ou usar uma URL para um ícone.</span><span class="sxs-lookup"><span data-stu-id="8392a-p119">Add an icon to show in the  **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon. You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="8392a-171">As etapas a seguir mostram como adicionar um arquivo de ícone à solução do Visual Studio:</span><span class="sxs-lookup"><span data-stu-id="8392a-171">The following steps show how to add an icon file to the Visual Studio solution:</span></span>
    
1. <span data-ttu-id="8392a-172">No **Gerenciador de Soluções**, vá até a pasta chamada Imagens.</span><span class="sxs-lookup"><span data-stu-id="8392a-172">In  **Solution Explorer**, go to the folder named Images.</span></span>
    
2. <span data-ttu-id="8392a-p120">Para ser exibido na lista suspensa **Suplementos do Office**, o ícone deve ter 32 x 32 pixels. Por exemplo, instale o SDK do Project 2013, escolha a pasta **Imagens** e adicione o seguinte arquivo do SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span><span class="sxs-lookup"><span data-stu-id="8392a-p120">To be displayed in the  **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels. For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>
    
    <span data-ttu-id="8392a-175">Como alternativa, use seu próprio ícone de 32 x 32 ou copie a imagem a seguir para um arquivo chamado NewIcon.png e, em seguida, adicione esse arquivo à pasta `HelloProjectODataWeb\Images`:</span><span class="sxs-lookup"><span data-stu-id="8392a-175">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder:</span></span>
    
    ![Ícone do aplicativo HelloProjectOData](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="8392a-p121">No manifesto HelloProjectOData.xml, adicione um elemento **IconUrl** abaixo do elemento **Description**, em que o valor da URL do ícone é o caminho relativo para o arquivo do ícone de 32 x 32 pixels. Por exemplo, adicione a seguinte linha: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. O arquivo de manifesto HelloProjectOData.xml agora contém o seguinte (seu valor **Id** será diferente):</span><span class="sxs-lookup"><span data-stu-id="8392a-p121">In the HelloProjectOData.xml manifest, add an  **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file. For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. The HelloProjectOData.xml manifest file now contains the following (your  **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82 </Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />

        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="8392a-180">Criar conteúdo HTML para o suplemento HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="8392a-180">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="8392a-p122">O suplemento **HelloProjectOData** é um exemplo que inclui as saídas de erro e de depuração. Ele não se destina a uso em produção. Antes de começar a escrever conteúdo HTML, crie a interface do usuário e a experiência para o suplemento, e descreva as funções JavaScript que interagem com o código HTML. Para saber mais, confira [Diretrizes de design para suplementos do Office](../design/add-in-design.md).</span><span class="sxs-lookup"><span data-stu-id="8392a-p122">The  **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use. Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code. For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="8392a-p123">O painel de tarefas mostra o nome de exibição do suplemento na parte superior, que é o valor do elemento **DisplayName** no manifesto. O elemento **body** no arquivo HelloProjectOData.html contém outros elementos de interface do usuário, da seguinte maneira:</span><span class="sxs-lookup"><span data-stu-id="8392a-p123">The task pane shows the add-in display name at the top, which is the value of the  **DisplayName** element in the manifest. The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="8392a-186">Um subtítulo indica a funcionalidade ou o tipo de operação geral, por exemplo, **CONSULTA REST ODATA**.</span><span class="sxs-lookup"><span data-stu-id="8392a-186">A subtitle indicates the general functionality or type of operation, for example,  **ODATA REST QUERY**.</span></span>
    
- <span data-ttu-id="8392a-p124">O botão **Obter Ponto de Extremidade de ProjectData** chama a função **setOdataUrl** para obter o ponto de extremidade do serviço **ProjectData** e exibi-lo em uma caixa de texto. Se o projeto não estiver conectado ao Project Web App, o suplemento chama um identificador de erro para exibir uma mensagem de erro pop-up.</span><span class="sxs-lookup"><span data-stu-id="8392a-p124">The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function to get the endpoint of the **ProjectData** service, and display it in a text box. If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>
    
- <span data-ttu-id="8392a-p125">O botão **Comparar Todos os Projetos** fica desabilitado até que o suplemento obtenha um ponto de extremidade OData válido. Ao selecionar o botão, ele chama a função **retrieveOData**, que usa uma consulta REST para obter os dados de trabalho e custo do projeto do serviço **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p125">The  **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint. When you select the button, it calls the **retrieveOData** function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>
    
- <span data-ttu-id="8392a-p126">Uma tabela exibe os valores médios de custo do projeto, custo real, trabalho e porcentagem concluída. A tabela também compara os valores atuais do projeto ativo com a média. Se o valor atual for maior que a média de todos os projetos, será exibido em vermelho. Se o valor atual for menor que a média, será exibido em verde. Se o valor atual não estiver disponível, a tabela exibirá **NA** em azul.</span><span class="sxs-lookup"><span data-stu-id="8392a-p126">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue  **NA**.</span></span>
    
    <span data-ttu-id="8392a-196">A função **retrieveOData** aciona a função **parseODataResult**, que calcula e exibe os valores da tabela.</span><span class="sxs-lookup"><span data-stu-id="8392a-196">The  **retrieveOData** function calls the **parseODataResult** function, which calculates and displays values for the table.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="8392a-p127">Neste exemplo, os dados de trabalho e custo para o projeto ativo são derivados dos valores publicados. Se você alterar os valores no Project, o serviço **ProjectData** não terá as alterações até que o projeto seja publicado.</span><span class="sxs-lookup"><span data-stu-id="8392a-p127">In this example, cost and work data for the active project are derived from the published values. If you change values in Project, the  **ProjectData** service does not have the changes until the project is published.</span></span>


### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="8392a-p128">Procedimento 4. Para criar o conteúdo HTML</span><span class="sxs-lookup"><span data-stu-id="8392a-p128">Procedure 4. To create the HTML content</span></span>

1. <span data-ttu-id="8392a-p129">No elemento **head** do arquivo Home.html, adicione elementos **link** extras para os arquivos CSS que seu suplemento usa. O modelo de projeto do Visual Studio inclui um link para o arquivo App.css que você pode usar para os estilos CSS personalizados.</span><span class="sxs-lookup"><span data-stu-id="8392a-p129">In the  **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses. The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>
    
2. <span data-ttu-id="8392a-p130">Adicione elementos **script** extras para bibliotecas JavaScript que o suplemento usa. O modelo de projeto inclui links para os arquivos jQuery- _[versão]_.js, office.js e MicrosoftAjax.js na pasta **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p130">Add any additional  **script** elements for JavaScript libraries that your add-in uses. The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the  **Scripts** folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="8392a-p131">Antes de implantar o suplemento, mude a referência office.js e a referência jQuery para a referência CDN (rede de distribuição de conteúdo). A referência CDN fornece a versão mais recente e melhora o desempenho.</span><span class="sxs-lookup"><span data-stu-id="8392a-p131">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="8392a-p132">O suplemento **HelloProjectOData** também usa o arquivo SurfaceErrors.js, que exibe os erros em uma mensagem pop-up. Você pode copiar o código da seção _Programação Robusta_ de [Cria seu primeiro suplemento do painel de tarefas do Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md) e adicionar um arquivo SurfaceErrors.js na pasta **Scripts\Office** do projeto **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p132">The  **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message. You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the  **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>
    
    <span data-ttu-id="8392a-209">Esse é o código HTML atualizado para o elemento **head**, com a linha adicional para o arquivo SurfaceErrors.js:</span><span class="sxs-lookup"><span data-stu-id="8392a-209">Following is the updated HTML code for the  **head** element, with the additional line for the SurfaceErrors.js file:</span></span>
    
    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>
    
    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    
    <!-- Add your CSS styles to the following file -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />
    
    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>
    
    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->
    
    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>
    
    <!-- Add your JavaScript to the following files -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="8392a-p133">No elemento **corpo**, exclua o código existente do modelo e adicione o código para a interface de usuário. Se um elemento deve ser preenchido com os dados ou manipulado por uma instrução jQuery, deve incluir um atributo **id** exclusivo. No código a seguir, os atributos **id** para os elementos **button**, **span** e **td** (definição de célula de tabela) que as funções jQuery usam são mostrados em negrito.</span><span class="sxs-lookup"><span data-stu-id="8392a-p133">In the **body** element, delete the existing code from the template, and then add the code for the user interface. If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute. In the following code, the **id** attributes for the **button**,  **span**, and  **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>
    
   <span data-ttu-id="8392a-p134">O seguinte HTML adiciona uma imagem gráfica, que poderia ser um logotipo da empresa. Você pode usar o logotipo que quiser ou copiar o arquivo NewLogo.png do download do Project 2013 SDK e depois usar o **Gerenciamento de Soluções** para adicionar o arquivo à pasta `HelloProjectODataWeb\Images`.</span><span class="sxs-lookup"><span data-stu-id="8392a-p134">The following HTML adds a graphic image, which could be a company logo. You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use  **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>
    
    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```


## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="8392a-215">Criar o código JavaScript para o suplemento</span><span class="sxs-lookup"><span data-stu-id="8392a-215">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="8392a-p135">O modelo para um suplemento de painel de tarefas do Project inclui código de inicialização padrão que foi projetado para demonstrar ações get e set básicas para dados em um documento no caso de um suplemento típico do Office 2013. Como o Project 2013 não dá suporte a ações que gravam no projeto ativo e o suplemento **HelloProjectOData** não usa o método **getSelectedDataAsync**, você pode excluir o script na função **Office.initialize** e excluir as funções **setData** e **getData** do arquivo HelloProjectOData.js padrão.</span><span class="sxs-lookup"><span data-stu-id="8392a-p135">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in. Because Project 2013 does not support actions that write to the active project, and the  **HelloProjectOData** add-in does not use the **getSelectedDataAsync** method, you can delete the script within the **Office.initialize** function, and delete the **setData** function and **getData** function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="8392a-p136">O JavaScript inclui constantes globais para a consulta REST e variáveis globais que são usadas em várias funções. O botão **Obter Ponto de Extremidade de ProjectData** chama a função **setOdataUrl**, que inicia as variáveis globais e determina se o Project está conectado ao Project Web App.</span><span class="sxs-lookup"><span data-stu-id="8392a-p136">The JavaScript includes global constants for the REST query and global variables that are used in several functions. The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="8392a-220">O restante do arquivo HelloProjectOData.js inclui duas funções: a função **retrieveOData** é chamada quando o usuário seleciona **Comparar Todos os Projetos** e a função **parseODataResult** calcula médias e preenche a tabela de comparação com valores que são formatados com cores e unidades.</span><span class="sxs-lookup"><span data-stu-id="8392a-220">The remainder of the HelloProjectOData.js file includes two functions: the  **retrieveOData** function is called when the user selects **Compare All Projects**; and the  **parseODataResult** function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="8392a-p137">Procedimento 5. Para criar o código JavaScript</span><span class="sxs-lookup"><span data-stu-id="8392a-p137">Procedure 5. To create the JavaScript code</span></span>

1. <span data-ttu-id="8392a-p138">Exclua todo o código do arquivo HelloProjectOData.js padrão e adicione as variáveis globais e a função **Office.initialize**. Nomes de variáveis que estão totalmente em maiúsculas sugerem que estas são constantes. Elas serão usadas mais tarde com a variável **_pwa** para criar a consulta REST neste exemplo.</span><span class="sxs-lookup"><span data-stu-id="8392a-p138">Delete all code in the default HelloProjectOData.js file, and then add the global variables and  **Office.initialize** function. Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>
    
    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData
    
    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

2. <span data-ttu-id="8392a-p139">Adicione **setOdataUrl** e as funções relacionadas. A função **setOdataUrl** chama **getProjectGuid** e **getDocumentUrl** para iniciar as variáveis globais. No [método getProjectFieldAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js), a função anônima para o parâmetro _callback_ habilita o botão **Comparar Todos os Projetos** usando o método **removeAttr** na biblioteca jQuery e exibe a URL do serviço **ProjectData**. Se o Project não estiver conectado ao Project Web App, a função gera um erro e exibe uma mensagem de erro pop-up. O arquivo SurfaceErrors.js inclui o método **throwError**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p139">Add  **setOdataUrl** and related functions. The **setOdataUrl** function calls **getProjectGuid** and **getDocumentUrl** to initialize the global variables. In the [getProjectFieldAsync method](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the **removeAttr** method in the jQuery library, and then displays the URL of the **ProjectData** service. If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message. The SurfaceErrors.js file includes the **throwError** method.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="8392a-p140">Se você executar o Visual Studio no computador do Project Server, use a depuração **F5**, sem comentar o código após a linha que inicializa a variável global **_pwa**. Para ativar usando o método jQuery **ajax** ao depurar no computador do Project Server, defina o valor **localhost** para a URL PWA. Se você executar o Visual Studio em um computador remoto, a URL do **localhost** não será necessária. Antes de implantar o suplemento, comente o código.</span><span class="sxs-lookup"><span data-stu-id="8392a-p140">If you run Visual Studio on the Project Server computer, to use  **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable. To enable using the jQuery **ajax** method when debugging on the Project Server computer, you must set the **localhost** value for the PWA URL.If you run Visual Studio on a remote computer, the  **localhost** URL is not required. Before you deploy the add-in, comment out that code.</span></span>

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);
    
                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;
    
                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    
    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

3. <span data-ttu-id="8392a-p141">Adicione a função **retrieveOData** que relaciona valores da consulta REST e chama a função **ajax** no jQuery para obter os dados solicitados do serviço **ProjectData**. A variável **support.cors** habilita o CORS (compartilhamento de recursos entre origens) com a função **ajax**. Se a instrução **support.cors** estiver ausente ou definida como **false**, a função **ajax** retorna um erro **Sem transporte**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p141">Add the  **retrieveOData** function, which concatenates values for the REST query and then calls the **ajax** function in jQuery to get the requested data from the **ProjectData** service. The **support.cors** variable enables cross-origin resource sharing (CORS) with the **ajax** function. If the **support.cors** statement is missing or is set to **false**, the  **ajax** function returns a **No transport** error.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="8392a-p142">O seguinte código funciona com uma instalação no local do Project Server 2013. Para o Project Online, use o OAuth para autenticação baseada em token. Para saber mais, confira [Como lidar com limitações de política de mesma origem nos Suplementos do Office](../develop/addressing-same-origin-policy-limitations.md).</span><span class="sxs-lookup"><span data-stu-id="8392a-p142">The following code works with an on-premises installation of Project Server 2013. For Project Online, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="8392a-p143">Na chamada **ajax**, use o parâmetro _headers_ ou o parâmetro _beforeSend_. O parâmetro _complete_ é uma função anônima, por isso está no mesmo escopo das variáveis no **retrieveOData**. A função para o parâmetro _complete_ exibe os resultados no controle **odataText** e também chama o método **parseODataResult** para analisar e exibir a resposta JSON. O parâmetro _error_ especifica a função **getProjectDataErrorHandler** nomeada, que grava uma mensagem de erro para o controle **odataText** e também usa o método **throwError** para exibir uma mensagem pop-up.</span><span class="sxs-lookup"><span data-stu-id="8392a-p143">In the **ajax** call, you can use either the _headers_ parameter or the _beforeSend_ parameter. The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in **retrieveOData**. The function for the  _complete_ parameter displays results in the **odataText** control and also calls the **parseODataResult** method to parse and display the JSON response. The _error_ parameter specifies the named **getProjectDataErrorHandler** function, which writes an error message to the **odataText** control and also uses the **throwError** method to display a pop-up message.</span></span>

    ```js
    /****************************************************************
    * Functions to get and parse the Project Server reporting data.
    *****************************************************************/
    
    // Get data about all projects on Project Server, 
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();
    
        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project Online.
        $.support.cors = true;
    
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;
    
                // xhr.responseText is the result from an XmlHttpRequest, which 
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);
    
                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);
    
                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }
    
    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

4. <span data-ttu-id="8392a-p144">Adicione o método **parseODataResult**, que desserializa e processa a resposta JSON do serviço OData. O método **parseODataResult** calcula valores médios dos dados de trabalho e de custo com precisão de uma ou duas casas decimais, formata valores com a cor correta, adiciona uma unidade (**$**, **hrs** ou **%**) e finalmente exibe os valores nas células da tabela especificada.</span><span class="sxs-lookup"><span data-stu-id="8392a-p144">Add the **parseODataResult** method, which deserializes and processes the JSON response from the OData service. The **parseODataResult** method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**,  **hrs**, or  **%**), and then displays the values in specified table cells.</span></span>
    
   <span data-ttu-id="8392a-p145">Se o GUID do projeto ativo corresponde ao valor **ProjectId**, a variável **myProjectIndex** é definida para o índice do projeto. Se **myProjectIndex** indica que o projeto ativo é publicado no Project Server, o método **parseODataResult** formata e exibe os dados de custo e trabalhos para o projeto. Se o projeto ativo não for publicado, os valores para o projeto ativo serão exibidos como um **ND** azul.</span><span class="sxs-lookup"><span data-stu-id="8392a-p145">If the GUID of the active project matches the  **ProjectId** value, the **myProjectIndex** variable is set to the project index. If **myProjectIndex** indicates the active project is published on Project Server, the **parseODataResult** method formats and displays cost and work data for that project. If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete   
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;
        
        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);
        
        // Display averages in the table, with the correct units. 
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";
            
        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);
            
            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);
            
            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;
            
            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;
            
            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";
            
            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";
            
            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"
            
            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```


## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="8392a-248">Testar o aplicativo HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="8392a-248">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="8392a-p146">Para testar e depurar o suplemento **HelloProjectOData** com o Visual Studio 2015, o Project Professional 2013 deve estar instalado no computador de desenvolvimento. Para habilitar cenários de teste diferentes, certifique-se de poder escolher se o Project abre no caso de arquivos no computador local ou se ele se conecta ao Project Web App. Por exemplo, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="8392a-p146">To test and debug the  **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer. To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App. For example, do the following steps:</span></span>

1. <span data-ttu-id="8392a-252">Na guia **ARQUIVO** na faixa de opções, escolha a guia **Informações** no modo de exibição Backstage e escolha **Gerenciar Contas**.</span><span class="sxs-lookup"><span data-stu-id="8392a-252">On the  **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>
    
2. <span data-ttu-id="8392a-p147">Na caixa de diálogo **Contas do Project Web App**, a lista **Contas disponíveis** pode ter várias contas do Project Web App além da conta **Computador** local. Na seção **Ao iniciar**, selecione **Escolher uma conta**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p147">In the  **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account. In the **When starting** section, select **Choose an account**.</span></span>
    
3. <span data-ttu-id="8392a-255">Feche o Project para que o Visual Studio possa iniciá-lo na depuração do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8392a-255">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>
    
<span data-ttu-id="8392a-256">Os testes básicos devem incluir o seguinte:</span><span class="sxs-lookup"><span data-stu-id="8392a-256">Basic tests should include the following:</span></span>

- <span data-ttu-id="8392a-p148">Execute o suplemento no Visual Studio, e abra um projeto publicado do Project Web App que contém dados de custos e trabalho. Verifique se o suplemento exibe o ponto de extremidade **ProjectData** e se exibe corretamente os dados de custo e de trabalho na tabela. Você pode usar a saída no controle **odataText** para verificar a consulta REST e outras informações.</span><span class="sxs-lookup"><span data-stu-id="8392a-p148">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data. Verify that the add-in displays the  **ProjectData** endpoint and correctly displays the cost and work data in the table. You can use the output in the **odataText** control to check the REST query and other information.</span></span>
    
- <span data-ttu-id="8392a-p149">Execute o suplemento novamente escolhendo o perfil do computador local na caixa de diálogo **Login** quando o Project inicia. Abra um arquivo .mpp local e teste o suplemento. Verifique se o suplemento exibe uma mensagem de erro ao tentar acessar o ponto de extremidade **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p149">Run the add-in again, where you choose the local computer profile in the  **Login** dialog box when Project starts. Open a local .mpp file, and then test the add-in. Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>
    
- <span data-ttu-id="8392a-p150">Execute o suplemento novamente e crie um projeto com tarefas com dados de custo e de trabalho. Você pode salvar o projeto no Project Web App, mas não o publique. Verifique se o suplemento exibe dados do Project Server, mas **NA** para o projeto atual.</span><span class="sxs-lookup"><span data-stu-id="8392a-p150">Run the add-in again, where you create a project that has tasks with cost and work data. You can save the project to Project Web App, but don't publish it. Verify that the add-in displays data from Project Server, but  **NA** for the current project.</span></span>
    

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="8392a-p151">Procedimento 6. Para testar o suplemento</span><span class="sxs-lookup"><span data-stu-id="8392a-p151">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="8392a-p152">Execute o Project Professional 2013, conecte-se ao Project Web App e crie um projeto de teste. Atribua tarefas aos recursos locais ou a recursos da empresa, defina vários valores de porcentagem concluída em algumas tarefas e publique o projeto. Feche o projeto, o que permite que o Visual Studio inicie o Project para depurar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="8392a-p152">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>
    
2. <span data-ttu-id="8392a-p153">No Visual Studio, pressione **F5**. Faça logon no Project Web App e abra o projeto que você criou na etapa anterior. Você pode abrir o projeto no modo somente leitura ou no modo de edição.</span><span class="sxs-lookup"><span data-stu-id="8392a-p153">In Visual Studio, press  **F5**. Log on to Project Web App, and then open the project that you created in the previous step. You can open the project in read-only mode or in edit mode.</span></span>
    
3. <span data-ttu-id="8392a-p154">Na guia **PROJETO** da faixa de opções, na lista suspensa **Suplementos do Office**, selecione **Hello ProjectData** (confira a Figura 5). O botão **Comparar Todos os Projetos** deve estar desativado.</span><span class="sxs-lookup"><span data-stu-id="8392a-p154">On the  **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5). The **Compare All Projects** button should be disabled.</span></span>
    
    <span data-ttu-id="8392a-276">*Figura 5. Iniciando o suplemento HelloProjectOData*</span><span class="sxs-lookup"><span data-stu-id="8392a-276">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![Testando o aplicativo HelloProjectOData](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="8392a-p155">No painel de tarefas **Hello ProjectData**, selecione **Obter Ponto de Extremidade de ProjectData**. A linha **projectDataEndPoint** deve mostrar a URL do serviço **ProjectData** e o botão **Comparar Todos os Projetos** deve estar habilitado (confira a Figura 6).</span><span class="sxs-lookup"><span data-stu-id="8392a-p155">In the  **Hello ProjectData** task pane, select **Get ProjectData Endpoint**. The  **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>
    
5. <span data-ttu-id="8392a-p156">Selecione **Comparar Todos os Projetos**. O suplemento pode pausar enquanto recupera os dados do serviço **ProjectData** e deve exibir os valores médios e atuais formatados na tabela.</span><span class="sxs-lookup"><span data-stu-id="8392a-p156">Select  **Compare All Projects**. The add-in may pause while it retrieves data from the  **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>
    
    <span data-ttu-id="8392a-282">*Figura 6. Exibindo resultados da consulta REST*</span><span class="sxs-lookup"><span data-stu-id="8392a-282">*Figure 6. Viewing results of the REST query*</span></span>

    ![Exibindo resultados da consulta REST](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="8392a-p157">Examine a saída na caixa de texto. Ela deve mostrar o caminho do documento, a consulta REST, as informações de status e os resultados JSON das chamadas a **ajax** e **parseODataResult**. A saída ajuda a entender, a criar e a depurar o código no método **parseODataResult**, como `projCost += Number(res.d.results[i].ProjectCost);`.</span><span class="sxs-lookup"><span data-stu-id="8392a-p157">Examine output in the text box. It should show the document path, REST query, status information, and JSON results from the calls to  **ajax** and **parseODataResult**. The output helps to understand, create, and debug code in the  **parseODataResult** method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>
    
    <span data-ttu-id="8392a-287">Veja a seguir um exemplo de saída com quebras de linha e espaços adicionados ao texto para fins de esclarecimentos, para três projetos em uma instância do Project Web App:</span><span class="sxs-lookup"><span data-stu-id="8392a-287">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance:</span></span>

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost
    
    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200
    
    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

7. <span data-ttu-id="8392a-p158">Pare a depuração (pressione **Shift+F5**) e pressione **F5** novamente para executar uma nova instância do Project. Na caixa de diálogo **Login**, escolha o perfil local **Computador**, e não o Project Web App. Crie ou abra um arquivo .mpp de projeto local, abra o painel de tarefas **Hello ProjectData** e selecione **Obter Ponto de Extremidade de ProjectData**. O suplemento deve exibir um erro **Sem conexão!** (confira a Figura 7) e o botão **Comparar Todos os Projetos** deve permanecer desativado.</span><span class="sxs-lookup"><span data-stu-id="8392a-p158">Stop debugging (press  **Shift + F5**), and then press  **F5** again to run a new instance of Project. In the **Login** dialog box, choose the local **Computer** profile, not Project Web App. Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**. The add-in should show a  **No connection!** error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>
    
   <span data-ttu-id="8392a-293">*Figura 7. Uso do suplemento sem uma conexão do Project Web App*</span><span class="sxs-lookup"><span data-stu-id="8392a-293">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![Usando o aplicativo sem uma conexão do Project Web App](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="8392a-p159">Pare a depuração e pressione **F5** novamente. Faça logon no Project Web App e crie um projeto com dados de custo e de trabalho. Você pode salvar o projeto, mas não o publique.</span><span class="sxs-lookup"><span data-stu-id="8392a-p159">Stop debugging, and then press  **F5** again. Log on to Project Web App, and then create a project that contains cost and work data. You can save the project, but don't publish it.</span></span>
    
   <span data-ttu-id="8392a-298">No painel de tarefas **Hello ProjectData**, quando você seleciona **Comparar Todos os Projetos**, deve ver um **ND** nos campos da coluna **Atual** (confira a Figura 8).</span><span class="sxs-lookup"><span data-stu-id="8392a-298">In the  **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue  **NA** for fields in the **Current** column (see Figure 8).</span></span>
    
   <span data-ttu-id="8392a-299">*Figura 8. Comparação de um projeto não publicado com outros projetos*</span><span class="sxs-lookup"><span data-stu-id="8392a-299">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![Como comparar um projeto não publicado com outros](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="8392a-p160">Mesmo que seu suplemento tenha funcionado corretamente nos testes anteriores, há outros testes que devem ser executados. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="8392a-p160">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="8392a-p161">Abra um projeto do Project Web App que não tenha nenhum dado de custo ou de trabalho para as tarefas. Você deve ver valores zerados nos campos da coluna **Atual**.</span><span class="sxs-lookup"><span data-stu-id="8392a-p161">Open a project from Project Web App that has no cost or work data for the tasks. You should see values of zero in the fields in the  **Current** column.</span></span>
    
- <span data-ttu-id="8392a-305">Teste um projeto sem tarefas.</span><span class="sxs-lookup"><span data-stu-id="8392a-305">Test a project that has no tasks.</span></span>
    
- <span data-ttu-id="8392a-p162">Se você modificar o suplemento e publicá-lo, deve executar testes semelhantes novamente com o suplemento publicado. Para outras considerações, confira [Próximas etapas](#next-steps).</span><span class="sxs-lookup"><span data-stu-id="8392a-p162">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>
    

> [!NOTE]
> <span data-ttu-id="8392a-p163">Há limites para a quantidade de dados que pode ser retornada em uma consulta do serviço **ProjectData**. A quantidade de dados varia conforme a entidade. Por exemplo, o conjunto de entidades **Projects** tem um limite padrão de 100 projetos por consulta, mas o conjunto de entidades **Risks** tem um limite padrão de 200. Para uma instalação de produção, o código no exemplo **HelloProjectOData** deve ser modificado para habilitar consultas de mais de 100 projetos. Para saber mais, confira [Próximas etapas](#next-steps) e [Consultar feeds OData para dados de relatório do Project](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="8392a-p163">There are limits to the amount of data that can be returned in one query of the  **ProjectData** service; the amount of data varies by entity. For example, the **Projects** entity set has a default limit of 100 projects per query, but the **Risks** entity set has a default limit of 200. For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects. For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>


## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="8392a-312">Exemplo de código para o suplemento de HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="8392a-312">Example code for the HelloProjectOData add-in</span></span>


### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="8392a-313">Arquivo HelloProjectOData.html</span><span class="sxs-lookup"><span data-stu-id="8392a-313">HelloProjectOData.html file</span></span>

<span data-ttu-id="8392a-314">O código a seguir está no arquivo `Pages\HelloProjectOData.html` do projeto **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="8392a-314">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```


### <a name="helloprojectodatajs-file"></a><span data-ttu-id="8392a-315">Arquivo HelloProjectOData.js</span><span class="sxs-lookup"><span data-stu-id="8392a-315">HelloProjectOData.js file</span></span>

<span data-ttu-id="8392a-316">O código a seguir está no arquivo `Scripts\Office\HelloProjectOData.js` do projeto **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="8392a-316">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer, 
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

/****************************************************************
* Functions to get and parse the Project Server reporting data.
*****************************************************************/

// Get data about all projects on Project Server, 
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project Online.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete   
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a><span data-ttu-id="8392a-317">Arquivo App.css</span><span class="sxs-lookup"><span data-stu-id="8392a-317">App.css file</span></span>

<span data-ttu-id="8392a-318">O código a seguir está no arquivo `Content\App.css` do projeto **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="8392a-318">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/
 
body
{
    font-size: 11pt;
}
h1 
{
    font-size: 22pt;
}
h2 
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.row_leftCol
{
    width: 20px;
    font-size: small; 
    font-weight: bold; 
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="8392a-319">Arquivo SurfaceErrors.js</span><span class="sxs-lookup"><span data-stu-id="8392a-319">SurfaceErrors.js file</span></span>

<span data-ttu-id="8392a-320">Você pode copiar o código para o arquivo SurfaceErrors.js da seção _Programação Robusta_ de [Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="8392a-320">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>


## <a name="next-steps"></a><span data-ttu-id="8392a-321">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="8392a-321">Next steps</span></span>

<span data-ttu-id="8392a-p164">Se **HelloProjectOData** fosse um suplemento de produção a ser vendido no AppSource ou distribuído em um catálogo de suplementos do SharePoint, ele deveria ser projetado de forma diferente. Por exemplo, não haveria nenhuma saída de depuração em uma caixa de texto e provavelmente nenhum botão para obter o ponto de extremidade do **ProjectData**. Você também precisaria reescrever a função **retireveOData** para lidar com instâncias do Project Web App que tenham mais de 100 projetos.</span><span class="sxs-lookup"><span data-stu-id="8392a-p164">If  **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint add-in catalog, it would be designed differently. For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint. You would also have to rewrite the **retireveOData** function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="8392a-p165">O suplemento deveria conter mais verificações de erro, além de lógica para capturar e explicar ou mostrar casos extremos. Por exemplo, se uma instância do Project Web App tiver mil projetos com uma duração média de cinco dias e custo médio de US$ 2.400, e o projeto ativo for o único que tem uma duração de mais de 20 dias, a comparação de custo e trabalho poderá ficar desequilibrada. Isso poderia ser exibido com um gráfico de frequência. Você poderia adicionar opções para exibir a duração, comparar projetos de tamanhos semelhantes ou comparar projetos de um mesmo departamento ou de departamentos diferentes. Ou poderia adicionar uma forma de o usuário selecionar os campos a exibir em uma lista.</span><span class="sxs-lookup"><span data-stu-id="8392a-p165">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="8392a-p166">Para outras consultas do serviço **ProjectData**, há limites para o comprimento da cadeia de consulta, que afeta o número de etapas que uma consulta pode executar de um conjunto pai para um objeto em um conjunto filho. Por exemplo, uma consulta de duas etapas de **Projects** para **Tasks** para itens de tarefa funciona, mas uma consulta de três etapas, como **Projects** para **Tasks** para **Assignments** para itens de atribuição pode exceder o comprimento máximo de URL padrão. Para saber mais, confira [Consultar feeds OData para dados de relatório do Project](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="8392a-p166">For other queries of the  **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection. For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length. For more information, see [Querying OData feeds for Project reporting data](https://docs.microsoft.com/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

<span data-ttu-id="8392a-333">Se você modificar o suplemento **HelloProjectOData** para uso em produção, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="8392a-333">If you modify the  **HelloProjectOData** add-in for production use, do the following steps:</span></span>

- <span data-ttu-id="8392a-334">No arquivo HelloProjectOData.html, para obter melhor desempenho, mude a referência ao office.js do projeto local para a referência da CDN:</span><span class="sxs-lookup"><span data-stu-id="8392a-334">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>
    
    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="8392a-p167">Reescreva a função **retrieveOData** para habilitar consultas de mais de 100 projetos. Por exemplo, você pode obter o número de projetos com uma consulta `~/ProjectData/Projects()/$count` e para usar os operadores _$skip_ e _$top_ na consulta REST para dados de projeto. Execute várias consultas em sequência e tire a média dos dados de cada consulta. Cada consulta de dados de projeto teria a forma:</span><span class="sxs-lookup"><span data-stu-id="8392a-p167">Rewrite the  **retrieveOData** function to enable queries of more than 100 projects. For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data. Run multiple queries in a loop, and then average the data from each query. Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`
    
  <span data-ttu-id="8392a-p168">Para mais informaçãos, veja [Opções de consulta do sistema OData usando o ponto de extremidade REST](https://docs.microsoft.com/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). Você também pode usar o comando [Set-SPProjectOdataConfiguration](http://technet.microsoft.com/library/jj219516%28v=office.15%29.aspx) no Windows PowerShell para substituir o tamanho de página padrão para uma consulta do conjunto de entidades **Projetos** (ou de qualquer um dos 33 conjuntos de entidades). Veja [ProjectData - Referência do serviço OData do projeto](https://docs.microsoft.com/previous-versions/office/project-odata/jj163015(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="8392a-p168">For more information, see [OData System Query Options Using the REST Endpoint](https://docs.microsoft.com/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](http://technet.microsoft.com/library/jj219516%28v=office.15%29.aspx) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](https://docs.microsoft.com/previous-versions/office/project-odata/jj163015(v=office.15)).</span></span>
    
- <span data-ttu-id="8392a-342">Para implantar o suplemento, confira [Publicar seu suplemento do Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="8392a-342">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>
    

## <a name="see-also"></a><span data-ttu-id="8392a-343">Confira também</span><span class="sxs-lookup"><span data-stu-id="8392a-343">See also</span></span>

- [<span data-ttu-id="8392a-344">Suplementos do painel de tarefas para Project</span><span class="sxs-lookup"><span data-stu-id="8392a-344">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="8392a-345">Criar seu primeiro suplemento de painel de tarefas para o Project 2013 usando um editor de texto</span><span class="sxs-lookup"><span data-stu-id="8392a-345">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- <span data-ttu-id="8392a-346">[ProjectData - referência do serviço OData do Project](https://docs.microsoft.com/previous-versions/office/project-odata/jj163015(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="8392a-346">[ProjectData - Project OData service reference](https://docs.microsoft.com/previous-versions/office/project-odata/jj163015(v=office.15))</span></span> 
- [<span data-ttu-id="8392a-347">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8392a-347">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md) 
- [<span data-ttu-id="8392a-348">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="8392a-348">Publish your Office Add-in</span></span>](../publish/publish.md)
    
