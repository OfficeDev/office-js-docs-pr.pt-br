---
title: Crie o seu primeiro suplemento de painel de tarefas para o Microsoft Project usando um editor de texto
description: ''
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: 36e2688240ad348669e7d6845f371997cd3c3ec2
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/16/2019
ms.locfileid: "37524273"
---
# <a name="create-your-first-task-pane-add-in-for-microsoft-project-by-using-a-text-editor"></a><span data-ttu-id="127dc-102">Crie o seu primeiro suplemento de painel de tarefas para o Microsoft Project usando um editor de texto</span><span class="sxs-lookup"><span data-stu-id="127dc-102">Create your first task pane add-in for Microsoft Project by using a text editor</span></span>

<span data-ttu-id="127dc-103">Você pode criar um suplemento de painel de tarefas para o Project Standard 2013, o Project Professional 2013 ou versões posteriores usando o gerador Yeoman para suplementos do Office. Este artigo descreve como criar um suplemento simples que usa um manifesto XML que aponta para um arquivo HTML em um compartilhamento de arquivos.</span><span class="sxs-lookup"><span data-stu-id="127dc-103">You can create a task pane add-in for Project Standard 2013, Project Professional 2013, or later versions using the Yeoman generator for Office Add-ins. This article describes how to create a simple add-in that uses an XML manifest that points to an HTML file on a file share.</span></span> <span data-ttu-id="127dc-104">O suplemento de exemplo teste do Project OM testa algumas funções de JavaScript que usam o modelo de objeto para suplementos. Depois que você usar a **Central de confiabilidade** no projeto para registrar o compartilhamento de arquivo que contém o arquivo de manifesto, você pode abrir a tarefa do painel suplemento do **Projeto** na faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="127dc-104">The Project OM Test sample add-in tests some JavaScript functions that use the object model for add-ins. After you use the  **Trust Center** in Project to register the file share that contains the manifest file, you can open the task pane add-in from the **Project** tab on the ribbon.</span></span> <span data-ttu-id="127dc-105">(O código de exemplo deste artigo é baseado em um aplicativo de teste de Arvind Iyer, da Microsoft Corporation).</span><span class="sxs-lookup"><span data-stu-id="127dc-105">(The sample code in this article is based on a test application by Arvind Iyer, Microsoft Corporation.)</span></span>

<span data-ttu-id="127dc-106">O Project usa o mesmo esquema de manifesto de suplemento que outros clientes do Microsoft Office, e grande parte da mesma API Java.</span><span class="sxs-lookup"><span data-stu-id="127dc-106">Project uses the same add-in manifest schema that other Microsoft Office clients use, and much of the same JavaScript API.</span></span> <span data-ttu-id="127dc-107">O código completo para o suplemento que está descrito neste artigo está disponível no subdiretório `Samples\Apps` do download do SDK do Project 2013.</span><span class="sxs-lookup"><span data-stu-id="127dc-107">The complete code for the add-in that is described in this article is available in the  `Samples\Apps` subdirectory of the Project 2013 SDK download.</span></span>

<span data-ttu-id="127dc-108">O suplemento de exemplo Teste do Project OM pode obter o GUID de uma tarefa, as propriedades do aplicativo e o projeto ativo.</span><span class="sxs-lookup"><span data-stu-id="127dc-108">The Project OM Test sample add-in can get the GUID of a task and properties of the application and the active project.</span></span> <span data-ttu-id="127dc-109">Se o Project Professional 2013 abre um projeto que está em uma biblioteca do SharePoint, o suplemento pode mostrar a URL do projeto.</span><span class="sxs-lookup"><span data-stu-id="127dc-109">If Project Professional 2013 opens a project that is in a SharePoint library, the add-in can show the URL of the project.</span></span> 

<span data-ttu-id="127dc-p104">O [download do SDK do Project 2013](https://www.microsoft.com/download/details.aspx?id=30435%20) inclui o código-fonte completo. Ao extrair e instalar o SDK e exemplos que estão no arquivo Project2013SDK.msi, confira o `\Samples\Apps\Copy_to_AppManifests_FileShare`subdiretório do arquivo de manifesto e o `\Samples\Apps\Copy_to_AppSource_FileShare`subdiretório do código-fonte.</span><span class="sxs-lookup"><span data-stu-id="127dc-p104">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes the complete source code. When you extract and install the SDK and samples that are in the Project2013SDK.msi file, see the `\Samples\Apps\Copy_to_AppManifests_FileShare` subdirectory for the manifest file and the `\Samples\Apps\Copy_to_AppSource_FileShare` subdirectory for the source code.</span></span> 

<span data-ttu-id="127dc-112">O exemplo JSOMCall.html usa funções JavaScript nos arquivos office.js e project-15.js, que estão incluídos.</span><span class="sxs-lookup"><span data-stu-id="127dc-112">The JSOMCall.html sample uses JavaScript functions in the office.js file and project-15.js file, which are included.</span></span> <span data-ttu-id="127dc-113">Você pode usar os arquivos de depuração correspondentes (office.debug.js e project-15.debug.js) para examinar as funções.</span><span class="sxs-lookup"><span data-stu-id="127dc-113">You can use the corresponding debug files (office.debug.js and project-15.debug.js) to examine the functions.</span></span>

<span data-ttu-id="127dc-114">Para ver uma introdução sobre como usar o JavaScript em suplementos do Office, confira [Noções básicas sobre a API JavaScript para Office](../develop/understanding-the-javascript-api-for-office.md).</span><span class="sxs-lookup"><span data-stu-id="127dc-114">For an introduction to using JavaScript in Office Add-ins, see [Understanding the JavaScript API for Office](../develop/understanding-the-javascript-api-for-office.md).</span></span>

## <a name="procedure-1-to-create-the-add-in-manifest-file"></a><span data-ttu-id="127dc-p106">Procedimento 1. Para criar o arquivo de manifesto do suplemento</span><span class="sxs-lookup"><span data-stu-id="127dc-p106">Procedure 1. To create the add-in manifest file</span></span>

<span data-ttu-id="127dc-p107">Crie um arquivo XML em um diretório local. O arquivo XML inclui o elemento **OfficeApp** e elementos filhos, que estão descritos em [Manifesto XML de suplementos do Office](../develop/add-in-manifests.md). Por exemplo, crie um arquivo denominado JSOM_SimpleOMCalls.xml contendo o seguinte XML (altere o valor do GUID do elemento **Id**).</span><span class="sxs-lookup"><span data-stu-id="127dc-p107">Create an XML file in a local directory. The XML file includes the **OfficeApp** element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md). For example, create a file named JSOM_SimpleOMCalls.xml that contains the following XML (change the GUID value of the **Id** element).</span></span>

```XML
<?xml version="1.0" encoding="utf-8"?>
   <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
              xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
              xsi:type="TaskPaneApp">
     <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
     <Id>93A26520-9414-492F-994B-4983A1C7A607</Id>
     <Version>15.0</Version>
     <ProviderName>Microsoft</ProviderName>
     <DefaultLocale>en-us</DefaultLocale>
     <DisplayName DefaultValue="Project OM Test">
       <Override Locale="fr-fr" Value="Le Project OM Test"/>
     </DisplayName>
     <Description DefaultValue="Test the task pane add-in object model for Project - English (US)">
       <Override Locale="fr-fr" Value="Test the task pane add-in object model for Project - French (France)"/>
     </Description>
     <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
     <Hosts>
       <Host Name="Project"/>
       <Host Name="Workbook"/>
       <Host Name="Document"/>
     </Hosts>
    <DefaultSettings>
       <SourceLocation DefaultValue="\\ServerName\AppSource\JSOMCall.html">
         <Override Locale="fr-fr" Value="\\ServerName\AppSource\JSOMCall.html"/>
       </SourceLocation>
     </DefaultSettings>
     <Permissions>ReadWriteDocument</Permissions>
     <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
       <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
     </IconUrl>
     <AllowSnapshot>true</AllowSnapshot>
   </OfficeApp>
```

<span data-ttu-id="127dc-120">Para o Project, o elemento **OfficeApp** deve incluir o valor do atributo `xsi:type="TaskPaneApp"`.</span><span class="sxs-lookup"><span data-stu-id="127dc-120">For Project, the **OfficeApp** element must include the `xsi:type="TaskPaneApp"` attribute value.</span></span> <span data-ttu-id="127dc-121">O elemento **Id** é um GUID.</span><span class="sxs-lookup"><span data-stu-id="127dc-121">The **Id** element is a GUID.</span></span> <span data-ttu-id="127dc-122">O valor **SourceLocation** deve ser um caminho de compartilhamento de arquivos ou uma URL do SharePoint para o arquivo de origem HTML do suplemento ou o aplicativo web que é executado no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="127dc-122">The **SourceLocation** value must be a file share path or a SharePoint URL for the add-in HTML source file or the web application that runs in the task pane.</span></span> <span data-ttu-id="127dc-123">Confira [Suplementos do painel de tarefas para o Project](../project/project-add-ins.md) para acessar uma explicação dos outros elementos no arquivo do manifesto.</span><span class="sxs-lookup"><span data-stu-id="127dc-123">For an explanation of the other elements in manifest file, see [Task pane add-ins for Project](../project/project-add-ins.md).</span></span>

<span data-ttu-id="127dc-p109">O Procedimento 2 mostra como criar o arquivo HTML que o manifesto JSOM_SimpleOMCalls.xml especifica para o suplemento de teste do Project. Botões especificados no arquivo HTML chamam funções JavaScript relacionadas. Você pode adicionar funções JavaScript no arquivo HTML ou colocá-las em um arquivo .js separado.</span><span class="sxs-lookup"><span data-stu-id="127dc-p109">Procedure 2 shows how to create the HTML file that the JSOM_SimpleOMCalls.xml manifest specifies for the Project test add-in. Buttons that are specified in the HTML file call related JavaScript functions. You can add the JavaScript functions within the HTML file, or put them in a separate .js file.</span></span>

## <a name="procedure-2-to-create-the-source-files-for-the-project-om-test-add-in"></a><span data-ttu-id="127dc-p110">Procedimento 2. Para criar os arquivos de origem para o suplemento Teste de modelo de objeto do Project</span><span class="sxs-lookup"><span data-stu-id="127dc-p110">Procedure 2. To create the source files for the Project OM Test add-in</span></span>

1. <span data-ttu-id="127dc-129">Crie um arquivo HTML com um nome especificado pelo elemento **SourceLocation** no manifesto JSOM_SimpleOMCalls.xml.</span><span class="sxs-lookup"><span data-stu-id="127dc-129">Create an HTML file with a name that is specified by the **SourceLocation** element in the JSOM_SimpleOMCalls.xml manifest.</span></span> 

   <span data-ttu-id="127dc-130">Por exemplo, crie o arquivo JSOMCall.html no `C:\Project\AppSource` diretório.</span><span class="sxs-lookup"><span data-stu-id="127dc-130">For example, create theJSOMCall.html file in the `C:\Project\AppSource` directory.</span></span> <span data-ttu-id="127dc-131">Embora você possa usar um editor de texto simples para criar os arquivos de origem, é mais fácil usar uma ferramenta como o código do Visual Studio, que funciona com tipos específicos de documentos (como HTML e JavaScript) e tem outros auxílios de edição.</span><span class="sxs-lookup"><span data-stu-id="127dc-131">Although you can use a simple text editor to create the source files, it is easier to use a tool such as Visual Studio code, which works with specific document types (such as HTML and JavaScript) and has other editing aids.</span></span> <span data-ttu-id="127dc-132">Se você ainda não tiver feito o exemplo da Pesquisa do Bing descrito em [Suplementos de painel de tarefas para Project](../project/project-add-ins.md), o Procedimento 3 mostra como criar o `\\ServerName\AppSource` compartilhamento de arquivos que o manifesto especifica.</span><span class="sxs-lookup"><span data-stu-id="127dc-132">If you have not already done the Bing Search example that is described in [Task pane add-ins for Project](../project/project-add-ins.md), Procedure 3 shows how to create the `\\ServerName\AppSource` file share that the manifest specifies.</span></span>

   <span data-ttu-id="127dc-133">O arquivo JSOMCall.html usa o arquivo MicrosoftAjax.js comum para a funcionalidade AJAX e o arquivo Office.js para a funcionalidade de suplemento em aplicativos do Microsoft Office 2013.</span><span class="sxs-lookup"><span data-stu-id="127dc-133">The JSOMCall.html file uses the common MicrosoftAjax.js file for AJAX functionality and the Office.js file for the add-in functionality in Microsoft Office 2013 applications.</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
        <head>
            <title>Project OM Sample Code</title>
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <script type="text/javascript" src="MicrosoftAjax.js"></script>

            <!-- Use the CDN reference to office.js when deploying your add-in. -->
            <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
            <script type="text/javascript" src="office.js"></script>
            <script type="text/javascript" src="JSOM_Sample.js"></script>
        </head>
        <body>
            <div id="Common_JSOM_API">
                OBJECT MODEL TESTS
            </div>

            <textarea id="text" rows="6" cols="25">This is the text result.</textarea>
        </body>
    </html>
    ```

   <span data-ttu-id="127dc-134">O elemento **textarea** especifica uma caixa de texto que mostra os resultados das funções de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="127dc-134">The **textarea** element specifies a text box that shows results of the JavaScript functions.</span></span>

   > [!NOTE]
   > <span data-ttu-id="127dc-135">Para o exemplo de teste do Project funcionar, copie os seguintes arquivos de download de SDK do Project 2013 no mesmo diretório do arquivo JSOMCall.html: Office.js, Project 15.js e MicrosoftAjax.js.</span><span class="sxs-lookup"><span data-stu-id="127dc-135">For the Project OM Test sample to work, copy the following files from the Project 2013 SDK download to the same directory as the JSOMCall.html file: Office.js, Project-15.js, and MicrosoftAjax.js.</span></span>

   <span data-ttu-id="127dc-136">A etapa 2 adiciona o arquivo JSOM_Sample.js para funções específicas que o suplemento de amostra de Teste de modelo de objeto do Project utiliza.</span><span class="sxs-lookup"><span data-stu-id="127dc-136">Step 2 adds the JSOM_Sample.js file for specific functions that the Project OM Test sample add-in uses.</span></span> <span data-ttu-id="127dc-137">Nas etapas posteriores, você adicionará outros elementos HTML para botões que acionam funções de JavaScript.</span><span class="sxs-lookup"><span data-stu-id="127dc-137">In later steps, you will add other HTML elements for buttons that call JavaScript functions.</span></span>

2. <span data-ttu-id="127dc-138">Crie um arquivo JavaScript denominado JSOM_Sample.js na mesma pasta do arquivo de JSOMCall.html.</span><span class="sxs-lookup"><span data-stu-id="127dc-138">Create a JavaScript file named JSOM_Sample.js in the same directory as the JSOMCall.html file.</span></span> 

   <span data-ttu-id="127dc-p113">O código a seguir obtém as informações de contexto e documentação do aplicativo usando funções no arquivo Office.js. O objeto **text** é a ID do controle **textarea** no arquivo HTML.</span><span class="sxs-lookup"><span data-stu-id="127dc-p113">The following code gets the application context and document information by using functions in the Office.js file. The **text** object is the ID of the **textarea** control in the HTML file.</span></span>

   <span data-ttu-id="127dc-p114">A variável **\_projDoc** é inicializada com um objeto **ProjectDocument**. O código inclui algumas funções de tratamento de erros simples e a função **getContextValues** que obtém o contexto do aplicativo e as propriedades contextuais do documento do Project. Para saber mais sobre o modelo de objeto do JavaScript para o Project, confira [API do JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="127dc-p114">The **\_projDoc** variable is initialized with a **ProjectDocument** object. The code includes some simple error handling functions, and the **getContextValues** function that gets application context and project document context properties. For more information about the JavaScript object model for Project, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office).</span></span>

    ```js
    /*
    * JavaScript functions for the Project OM Test example app
    * in the Project 2013 SDK.
    */

    var _projDoc;
    var _app;
    var taskGuid;
    var resourceGuid;

    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            _projDoc = Office.context.document;
            _app = Office.context;
        });
    }

    function logError(errorText) {
        text.value = "Error in " + errorText;
    }

    function logEventError(erroneousEvent) {
        logError("event " + erroneousEvent);
    }

    function logMethodError(methodName, errorName, errorMessage) {
        logError(methodName + " method.\nError name: " + errorName + "\nMessage: " + errorMessage);
    }

    // . . . Add other JavaScript functions here.

    function getContextValues() {
        getDocumentUrl();
        getDocumentMode();
        getApplicationContentLanguage();
        getApplicationDisplayLanguage();
    }

    function getDocumentUrl() {
        text.value ="Document URL:\n" +_projDoc.url;
    }

    function getDocumentMode() {
        var docMode = _projDoc.mode;
        text.value = text.value + "\n\nDocument mode: " + docMode;
    }

    function getApplicationContentLanguage() {
        text.value = text.value + "\nApp language: " + _app.contentLanguage;
    }

    function getApplicationDisplayLanguage() {
        text.value = text.value + "\nDisplay language: " + _app.displayLanguage;
    }
    ```

   <span data-ttu-id="127dc-144">Confira as informações sobre as funções no arquivo Office.debug.js em [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="127dc-144">For information about the functions in the Office.debug.js file, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office).</span></span> <span data-ttu-id="127dc-145">Por exemplo, a função **getDocumentUrl** obtém o caminho de URL ou do arquivo do projeto aberto.</span><span class="sxs-lookup"><span data-stu-id="127dc-145">For example, the **getDocumentUrl** function gets the URL or file path of the open project.</span></span>

3. <span data-ttu-id="127dc-146">Adicione funções JavaScript que acionam funções assíncronas em Office.js e Project-15.js para acessar dados selecionados:</span><span class="sxs-lookup"><span data-stu-id="127dc-146">Add JavaScript functions that call asynchronous functions in Office.js and Project-15.js to get selected data:</span></span>

   - <span data-ttu-id="127dc-p116">Por exemplo, **getSelectedDataAsync** é uma função geral no Office.js que obtém texto não formatado para os dados selecionados. Para saber mais, confira [objeto AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="127dc-p116">For example, **getSelectedDataAsync** is a general function in Office.js that gets unformatted text for the selected data. For more information, see [AsyncResult object](/javascript/api/office/office.asyncresult).</span></span>

   - <span data-ttu-id="127dc-p117">A função **getSelectedTaskAsync** em Project-15.js obtém o GUID da tarefa selecionada. Da mesma forma, a função **getSelectedResourceAsync** obtém o GUID do recurso selecionado. Se você chamar essas funções quando uma tarefa ou um recurso não estiver selecionado, as funções mostrarão um erro indefinido.</span><span class="sxs-lookup"><span data-stu-id="127dc-p117">The **getSelectedTaskAsync** function in Project-15.js gets the GUID of the selected task. Similarly, the **getSelectedResourceAsync** function gets the GUID of the selected resource. If you call those functions when a task or a resource is not selected, the functions show an undefined error.</span></span>

   - <span data-ttu-id="127dc-p118">A função **getTaskAsync** obtém o nome da tarefa e os nomes dos recursos atribuídos. Se a tarefa estiver em uma lista de tarefas do SharePoint sincronizada, **getTaskAsync** obtém a ID de tarefa na lista do SharePoint. Caso contrário, a ID de tarefa do SharePoint é 0.</span><span class="sxs-lookup"><span data-stu-id="127dc-p118">The **getTaskAsync** function gets the task name and the names of the assigned resources. If the task is in a synchronized SharePoint task list, **getTaskAsync** gets the task ID in the SharePoint list; otherwise, the SharePoint task ID is 0.</span></span>

     > [!NOTE]
     > <span data-ttu-id="127dc-p119">Para fins de demonstração, o código de exemplo inclui um bug. Se **taskGuid** estiver indefinida, os erros da função **getTaskAsync** são desativados. Se você obtiver um  GUID de tarefas válido e depois selecionar uma tarefa diferente, a função **getTaskAsync** obterá dados para a tarefa mais recente que foi operada pela função **getSelectedTaskAsync**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p119">For demonstration purposes, the example code includes a bug. If  **taskGuid** is undefined, the **getTaskAsync** function errors off. If you get a valid task GUID and then select a different task, the **getTaskAsync** function gets data for the most recent task that was operated on by the **getSelectedTaskAsync** function.</span></span>
  
   - <span data-ttu-id="127dc-p120">**getTaskFields**, **getResourceFields** e **getProjectFields** são funções locais que chamam **getTaskFieldAsync**, **getResourceFieldAsync** ou **getProjectFieldAsync** várias vezes para obter campos especificados de uma tarefa ou um recurso. No arquivo project-15.debug.js, as enumerações **ProjectTaskFields** e **ProjectResourceFields** mostram quais campos têm suporte.</span><span class="sxs-lookup"><span data-stu-id="127dc-p120">**getTaskFields**, **getResourceFields**, and **getProjectFields** are local functions that call **getTaskFieldAsync**, **getResourceFieldAsync**, or **getProjectFieldAsync** multiple times to get specified fields of a task or a resource. In the project-15.debug.js file, the **ProjectTaskFields** enumeration and the **ProjectResourceFields** enumeration show which fields are supported.</span></span>

   - <span data-ttu-id="127dc-159">A função **getSelectedViewAsync** obtém o tipo de exibição (definido na enumeração **ProjectViewTypes** no projeto 15.debug.js) e o nome do modo de exibição.</span><span class="sxs-lookup"><span data-stu-id="127dc-159">The **getSelectedViewAsync** function gets the type of view (defined in the **ProjectViewTypes** enumeration in project-15.debug.js) and the name of the view.</span></span>

   - <span data-ttu-id="127dc-p121">Se o projeto é sincronizado com uma lista de tarefas do SharePoint, a função **getWSSUrlAsync** obtém a URL e o nome da lista de tarefas. Se o projeto não está sincronizado com uma lista de tarefas do SharePoint, a função **getWSSUrlAsync** falha.</span><span class="sxs-lookup"><span data-stu-id="127dc-p121">If the project is synchronized with a SharePoint tasks list, the  **getWSSUrlAsync** function gets the URL and the name of the tasks list. If the project is not synchronized with a SharePoint tasks list, the **getWSSUrlAsync** function errors off.</span></span>

     > [!NOTE]
     > <span data-ttu-id="127dc-162">Para obter a URL do SharePoint e o nome da lista de tarefas, recomendamos que você use a função **getProjectFieldAsync** com as constantes **WSSUrl** e **WSSList** na enumeração [ProjectProjectFields](/javascript/api/office/office.projectprojectfields).</span><span class="sxs-lookup"><span data-stu-id="127dc-162">To get the SharePoint URL and name of the tasks list, we recommend that you use the  **getProjectFieldAsync** function with the **WSSUrl** and **WSSList** constants in the [ProjectProjectFields](/javascript/api/office/office.projectprojectfields) enumeration.</span></span>

   <span data-ttu-id="127dc-p122">Cada uma das funções no código a seguir inclui uma função anônima que é especificada por `function (asyncResult)`, que é um retorno de chamada que obtém o resultado assíncrono. Em vez de funções anônimas, você poderia usar funções nomeadas, que podem ajudar na capacidade de manutenção de suplementos complexos.</span><span class="sxs-lookup"><span data-stu-id="127dc-p122">Each of the functions in the following code includes an anonymous function that is specified by  `function (asyncResult)`, which is a callback that gets the asynchronous result. Instead of anonymous functions, you could use named functions, which can help with maintainability of complex add-ins.</span></span>

    ```js
    // Get the data in the selected cells of the grid in the active view.
    function getSelectedDataAsync() {
        _projDoc.getSelectedDataAsync(
            Office.CoercionType.Text,
            { ValueFormat: "Formatted" },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded)
                    text.value = asyncResult.value;
                else
                    logMethodError("getSelectedDataAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        );
    }

    // Get the GUID of the selected task.
    function getSelectedTaskAsync() {
        _projDoc.getSelectedTaskAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                taskGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get the GUID of the selected resource.
    function getSelectedResourceAsync() {
        _projDoc.getSelectedResourceAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = asyncResult.value;
                resourceGuid = asyncResult.value;
            }
            else {
                logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                                   asyncResult.error.message);
            }
        });
    }

    // Get data for the specified task.
    function getTaskAsync() {
        if (taskGuid != undefined) {
            _projDoc.getTaskAsync(
                taskGuid,
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        logMethodError("getTaskAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                    } else {
                        var taskInfo = asyncResult.value;
                        var taskOutput = "Task name: " + taskInfo.taskName +
                                         "\nGUID: " + taskGuid +
                                         "\nWSS Id: " + taskInfo.wssTaskId +
                                         "\nResourceNames: " + taskInfo.resourceNames;
                        text.value = taskOutput;
                    }
                }
            );
        } else {
            text.value = 'Task GUID not valid:\n' + taskGuid;
        } 
    }

    // Get additional data for task fields.
    function getTaskFields() {
        text.value = "";

        _projDoc. getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Name: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.ID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "ID: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Start: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Duration,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Duration: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Priority,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Priority: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getTaskFieldAsync(taskGuid, Office.ProjectTaskFields.Notes,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Notes: "
                        + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getTaskFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        ); 
    }

    // Get data for the specified resource fields.
    function getResourceFields() {
        text.value = "";

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Name,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Resource name: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Cost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.StandardRate,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Standard Rate: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualCost,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Cost: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.ActualWork,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Actual Work: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );

        _projDoc.getResourceFieldAsync(resourceGuid, Office.ProjectResourceFields.Units,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Units: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getResourceFieldAsync", asyncResult.error.name,
                                   asyncResult.error.message);
                }
            }
        );
    }

    // Get the URL and list name of the synchronized SharePoint task list.
    // Recommended: use getProjectField instead.
    function getWSSUrlAsync() {
        _projDoc.getWSSUrlAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                text.value = "SharePoint URL:\n" + asyncResult.value.serverUrl
                    + "\nList name: " + asyncResult.value.listName;
            }
            else {
                logMethodError("getWSSUrlAsync", asyncResult.error.name, asyncResult.error.message);
            }
        });
    }

    // Get the type and name of the selected view.
    function getSelectedViewAsync() {
        _projDoc.getSelectedViewAsync(function (asyncResult) {
            text.value = "View type: " + asyncResult.value.viewType
                + "\nName: " + asyncResult.value.viewName;
        });
    }

    // Get information about the active project.
    function getProjectFields() {
        text.value = "";

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Project GUID: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Start,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nStart: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.Finish,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nFinish: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProject " + errorText);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencyDigits,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nCurrency digits: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbol,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "Currency symbol: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.CurrencySymbolPosition,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSymbol position: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nProject web app URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint URL:\n  " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );

        _projDoc.getProjectFieldAsync(Office.ProjectProjectFields.WSSList,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    text.value = text.value + "\nSharePoint list: " + asyncResult.value.fieldValue + "\n";
                }
                else {
                    logMethodError("getProjectFieldAsync", asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    ```

4. <span data-ttu-id="127dc-p123">Adicione retornos de chamada e funções de manipulador de eventos JavaScript para registrar a seleção de tarefas, a seleção de recursos, exibir os manipuladores de eventos de alteração de seleção e desfazer o registro dos manipuladores de eventos. A função **manageEventHandlerAsync** adiciona ou remove o manipulador de eventos específico, dependendo do parâmetro _operation_. A operação pode ser **addHandlerAsync** ou **removeHandlerAsync**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p123">Add JavaScript event handler callbacks and functions to register the task selection, resource selection, and view selection change event handlers and to unregister the event handlers. The **manageEventHandlerAsync** function adds or removes the specified event handler, depending on the _operation_ parameter. The operation can be **addHandlerAsync** or **removeHandlerAsync**.</span></span>

   <span data-ttu-id="127dc-168">As funções **manageTaskEventHandler**, **manageResourceEventHandler** e **manageViewEventHandler** podem adicionar ou remover um manipulador de eventos, como especificado pelo parâmetro _docMethod_.</span><span class="sxs-lookup"><span data-stu-id="127dc-168">The **manageTaskEventHandler**, **manageResourceEventHandler**, and **manageViewEventHandler** functions can add or remove an event handler, as specified by the _docMethod_ parameter.</span></span>

    ```js
    // Task selection changed event handler.
    function onTaskSelectionChanged(eventArgs) {
        text.value = "In task selection change event handler";
    }

    // Resource selection changed event handler.
    function onResourceSelectionChanged(eventArgs) {
        text.value = "In Resource selection changed event handler";
    }

    // View selection changed event handler.
    function onViewSelectionChanged(eventArgs) {
        text.value = "In View selection changed event handler";
    }

    // Add or remove the specified event handler.
    function manageEventHandlerAsync(eventType, handler, operation, onComplete) {
        _projDoc[operation]   //The operation is addHandlerAsync or removeHandlerAsync.
        (
            eventType,
            handler,
            function (asyncResult) {
                if (onComplete) {
                    onComplete(asyncResult, operation);
                } else {
                    var message = "Operation: " + operation;
                    message = message + "\nStatus: " + asyncResult.status + "\n";
                    text.value = message;
                }
            }
        );
    }

    // Write the asyncResult status from the manageEventHandlerAsync function (optional). 
    function onComplete(asyncResult, operation) {
        var message = "In onComplete function for " + operation;
        message = message + "\nStatus: " + asyncResult.status;
        text.value = message;
    }

    // Add or remove a task selection changed event handler.
    function manageTaskEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.TaskSelectionChanged,      // The task selection changed event.
            onTaskSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a resource selection changed event handler.
    function manageResourceEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ResourceSelectionChanged,  // The resource selection changed event.
            onResourceSelectionChanged,                 // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }

    // Add or remove a view selection changed event handler.
    function manageViewEventHandler(docMethod) {
        manageEventHandlerAsync(
            Office.EventType.ViewSelectionChanged,      // The view selection changed event.
            onViewSelectionChanged,                     // The event handler.
            docMethod,                // The Office.Document method to add or remove an event handler.
            onComplete                // Manages the successful asyncResult data (optional).
        );
    }
    ```

5. <span data-ttu-id="127dc-p124">Para o corpo do documento HTML, adicione botões que chamam funções JavaScript para teste. Por exemplo, no elemento **div** para a API JSOM comum, adicione um botão de entrada que chama a função geral **getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p124">For the body of the HTML document, add buttons that call the JavaScript functions for testing. For example, in the  **div** element for the common JSOM API, add an input button that calls the general **getSelectedDataAsync** function.</span></span>

    ```HTML
    <body>
        <div id="Common_JSOM_API">
        OBJECT MODEL TESTS
        <br /><br />
        <strong>General function:</strong>
        <br />
        <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
            value="getSelectedDataAsync" />
        </div>
        <!--  more code . . .  -->
    ```

6. <span data-ttu-id="127dc-171">Adicione uma seção **div** com botões para funções de tarefas específicas do projeto e para o evento **TaskSelectionChanged**.</span><span class="sxs-lookup"><span data-stu-id="127dc-171">Add a **div** section with buttons for project-specific task functions and for the **TaskSelectionChanged** event.</span></span>

    ```HTML
    <div id="ProjectSpecificTask">
      <br />
      <strong>Project-specific task methods:</strong><br />
      <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
      <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
      <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
      <strong>Task selection changed:</strong>
      <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    ```

7. <span data-ttu-id="127dc-172">Adicionar seções **div** com botões para os métodos de recursos e eventos, métodos de exibição e eventos, propriedades do projeto e propriedades do contexto</span><span class="sxs-lookup"><span data-stu-id="127dc-172">Add  **div** sections with buttons for the resource methods and events, view methods and events, project properties, and context properties</span></span>

    ```HTML
    <div id="ResourceMethods">
      <br />
      <strong>Resource methods:</strong>
      <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
      <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
      <strong>Resource selection changed:</strong>
      <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ViewMethods">
      <br />
      <strong>View method:</strong>
      <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
      <strong>View selection changed:</strong>
      <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
      <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
    </div>
    <div id="ProjectMethods">
      <br />
      <strong>Project properties:</strong>
      <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
    </div>
    <div id="ContextVariables">
      <br />
      <strong>Context properties:</strong>
      <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
    </div>
    ```

8. <span data-ttu-id="127dc-p125">Para formatar elementos de botão, adicione um elemento CSS **style**. Por exemplo, adicione o seguinte como um filho do elemento **head**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p125">To format the button elements, add a CSS  **style** element. For example, add the following as a child of the **head** element.</span></span>

    ```HTML
    <style type="text/css">
        .button-wide
        {
            width: 210px;
            margin-top: 2px;
        }
        .button-narrow
        {
            width: 80px;
            margin-top: 2px;
        }
    </style>
    ```

<span data-ttu-id="127dc-175">O Procedimento 3 mostra como instalar e usar os recursos do suplemento Teste de modelo de objeto do Project.</span><span class="sxs-lookup"><span data-stu-id="127dc-175">Procedure 3 shows how to install and use the Project OM Test add-in features.</span></span>

## <a name="procedure-3-to-install-and-use-the-project-om-test-add-in"></a><span data-ttu-id="127dc-p126">Procedimento 3. Para instalar e usar o suplemento Teste de modelo de objeto do Project</span><span class="sxs-lookup"><span data-stu-id="127dc-p126">Procedure 3. To install and use the Project OM Test add-in</span></span>

1. <span data-ttu-id="127dc-p127">Crie um compartilhamento de arquivos para o diretório que contém o manifesto JSOM_SimpleOMCalls.xml. Você pode criar o compartilhamento de arquivos no computador local ou em um computador remoto que esteja acessível na rede. Por exemplo, se o manifesto estiver no diretório `C:\Project\AppManifests` no computador local, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="127dc-p127">Create a file share for the directory that contains the JSOM_SimpleOMCalls.xml manifest. You can create the file share on the local computer or on a remote computer that is accessible on the network. For example, if the manifest is in the  `C:\Project\AppManifests` directory on the local computer, run the following command:</span></span>

    `Net share AppManifests=C:\Project\AppManifests`

2. <span data-ttu-id="127dc-p128">Crie um compartilhamento de arquivos para o diretório que contenha os arquivos HTML e JavaScript para o suplemento Teste de modelo de objeto do Project. Verifique se o caminho de compartilhamento do arquivo corresponde ao caminho especificado no manifesto JSOM_SimpleOMCalls.xml. Por exemplo, se os arquivos estão no diretório `C:\Project\AppSource` no computador local, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="127dc-p128">Create a file share for the directory that contains the HTML and JavaScript files for the Project OM Test add-in. Ensure the file share path matches the path that is specified in the JSOM_SimpleOMCalls.xml manifest. For example, if the files are in the  `C:\Project\AppSource` directory on the local computer, run the following command:</span></span>

    `net share AppSource=C:\Project\AppSource`

3. <span data-ttu-id="127dc-184">No Project, abra a caixa de diálogo **Opções do Project**, escolha **Central de Confiabilidade** e escolha **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="127dc-184">In Project, open the **Project Options** dialog box, choose **Trust Center**, and then choose  **Trust Center Settings**.</span></span>

   <span data-ttu-id="127dc-185">O procedimento para registrar um suplemento também está descrito em [Suplementos do painel de tarefas para o Project](../project/project-add-ins.md), com informações adicionais.</span><span class="sxs-lookup"><span data-stu-id="127dc-185">The procedure for registering an add-in is also described in [Task pane add-ins for Project](../project/project-add-ins.md), with additional information.</span></span>

4. <span data-ttu-id="127dc-186">Na caixa de diálogo **Central de Confiabilidade**, no painel esquerdo, escolha **Catálogos de Suplementos Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="127dc-186">In the **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>

5. <span data-ttu-id="127dc-p129">Se você já tiver adicionado o caminho `\\ServerName\AppManifests` para o suplemento Pesquisa do Bing, pule esta etapa. Caso contrário, no painel **Catálogos de Suplementos Confiáveis**, adicione o caminho `\\ServerName\AppManifests` na caixa de texto **URL do Catálogo**, escolha **Adicionar catálogo**, habilite o compartilhamento de rede como origem padrão (confira a Figura 1) e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p129">If you have already added the `\\ServerName\AppManifests` path for the Bing Search add-in, skip this step. Otherwise, in the **Trusted Add-in Catalogs** pane, add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add catalog**, enable the network share as a default source (see Figure 1), and then choose  **OK**.</span></span>

   <span data-ttu-id="127dc-189">*Figura 1. Adicionar um compartilhamento de arquivos de rede para manifestos de suplementos*</span><span class="sxs-lookup"><span data-stu-id="127dc-189">*Figure 1. Adding a network file share for add-in manifests*</span></span>

   ![Adicionando um compartilhamento de arquivos de rede para manifestos de aplicativos](../images/pj15-create-simple-agave-manage-catalogs.png)

6. <span data-ttu-id="127dc-p130">Depois de adicionar novos suplementos ou alterar o código-fonte, reinicie o Project. Na faixa de opções **PROJETO**, escolha o menu suspenso **Suplementos do Office** e escolha **Ver Tudo**. Na caixa de diálogo **Inserir Suplemento**, escolha **PASTA COMPARTILHADA** (confira a Figura 2), selecione **Teste de modelo de objeto do Project** e escolha **Inserir**. O suplemento Teste de modelo de objeto do Project inicia em um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="127dc-p130">After you add new add-ins, or change the source code, restart Project. On the  **PROJECT** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the  **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2), select **Project OM Test**, and then choose  **Insert**. The Project OM Test add-in starts in a task pane.</span></span>

   <span data-ttu-id="127dc-195">*Figura 2. Iniciando o suplemento do Teste de Modelo de Objeto do Project contido em um compartilhamento de arquivo*</span><span class="sxs-lookup"><span data-stu-id="127dc-195">*Figure 2. Starting the Project OM Test add-in that is on a file share*</span></span>

   ![Inserindo um aplicativo](../images/pj15-create-simple-agave-start-agave-app.png)

7. <span data-ttu-id="127dc-p131">No Project, crie e salve um projeto simples que tenha pelo menos duas tarefas. Por exemplo, crie tarefas chamadas T1 e T2 e um marco chamado M1, e defina as durações das tarefas e os predecessores de maneira semelhante à Figura 3. Escolha a guia **PROJETO** na faixa de opções, selecione a linha inteira para a tarefa T2 e escolha o botão **getSelectedDataAsync** no painel de tarefas. A Figura 3 mostra os dados que estão selecionados na caixa de texto do suplemento **Teste de modelo de objeto do Project**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p131">In Project, create and save a simple project that has at least two tasks. For example, create tasks named T1, T2, and a milestone named M1, and then set the task durations and predecessors to be similar to those in Figure 3. Choose the  **PROJECT** tab on the ribbon, select the entire row for task T2, and then choose the **getSelectedDataAsync** button in the task pane. Figure 3 shows the data that is selected in the text box of the **Project OM Test** add-in.</span></span>

   <span data-ttu-id="127dc-201">*Figura 3. Usando o suplemento do Teste de Modelo de Objeto do Project*</span><span class="sxs-lookup"><span data-stu-id="127dc-201">*Figure 3. Using the Project OM Test add-in*</span></span>

   ![Usando o aplicativo do Teste de Modelo de Objeto do Project](../images/pj15-create-simple-agave-project-om-test.png)

8. <span data-ttu-id="127dc-p132">Selecione a célula na coluna **Duração** da primeira tarefa e escolha o botão **getSelectedDataAsync** no suplemento **Teste de modelo de objeto do Project**. A função **getSelectedDataAsync** define o valor da caixa de texto para mostrar `2 days`.</span><span class="sxs-lookup"><span data-stu-id="127dc-p132">Select the cell in the  **Duration** column for the first task, and then choose the **getSelectedDataAsync** button in the **Project OM Test** add-in. The **getSelectedDataAsync** function sets the text box value to show `2 days`.</span></span> 

9. <span data-ttu-id="127dc-p133">Selecione as três células de **Duração** para todas as três tarefas. A função **getSelectedDataAsync** retorna valores de texto separados por ponto e vírgula para células selecionadas em linhas diferentes, por exemplo, `2 days;4 days;0 days`.</span><span class="sxs-lookup"><span data-stu-id="127dc-p133">Select the three  **Duration** cells for all three tasks. The **getSelectedDataAsync** function returns semicolon-separated text values for cells selected in different rows, for example, `2 days;4 days;0 days`.</span></span>

   <span data-ttu-id="127dc-p134">A função **getSelectedDataAsync** retorna valores de texto separados por vírgula para células selecionadas em uma linha. Por exemplo, na Figura 3, a linha inteira da tarefa T2 está selecionada. Quando você escolhe **getSelectedDataAsync**, a caixa de texto mostra o seguinte: `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span><span class="sxs-lookup"><span data-stu-id="127dc-p134">The  **getSelectedDataAsync** function returns comma-separated text values for cells selected within a row. For example in Figure 3, the entire row for task T2 is selected. When you choose **getSelectedDataAsync**, the text box shows the following:  `,Auto Scheduled,T2,4 days,Thu 6/14/12,Tue 6/19/12,1,,<NA>`</span></span>

   <span data-ttu-id="127dc-p135">As colunas **Indicadores** e **Nomes de Recursos** estão vazias, portanto, a matriz de texto mostra valores vazios para essas colunas. O valor `<NA>` é para a célula **Adicionar Nova Coluna**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p135">The  **Indicators** column and the **Resource Names** column are both empty, so the text array shows empty values for those columns. The `<NA>` value is for the **Add New Column** cell.</span></span>

10. <span data-ttu-id="127dc-p136">Selecione qualquer célula na linha da tarefa T2, ou a linha inteira da tarefa T2, e escolha **getSelectedTaskAsync**. A caixa de texto mostra o valor de tarefa do GUID, por exemplo, `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`. O Project armazena esse valor na variável global **taskGuid** do suplemento **Teste de modelo de objeto do Project**.</span><span class="sxs-lookup"><span data-stu-id="127dc-p136">Select any cell in the row for task T2, or the entire row for task T2, and then choose  **getSelectedTaskAsync**. The text box shows the task GUID value, for example,  `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`. Project stores that value in the global  **taskGuid** variable of the **Project OM Test** add-in.</span></span>

11. <span data-ttu-id="127dc-p137">Selecione **getTaskAsync**. Se a variável **taskGuid** contém o GUID para a tarefa T2, a caixa de texto exibe as informações da tarefa. O valor **ResourceNames** fica vazio.</span><span class="sxs-lookup"><span data-stu-id="127dc-p137">Select **getTaskAsync**. If the **taskGuid** variable contains the GUID for task T2, the text box displays the task information. The **ResourceNames** value is empty.</span></span>

    <span data-ttu-id="127dc-p138">Create two local resources R1 andR2, assign them to task T2 at 50% each, and choose  **getTaskAsync** again. The results in the text box include the resource information. If the task is in a synchronized SharePoint task list, the results also include the SharePoint task ID.</span><span class="sxs-lookup"><span data-stu-id="127dc-p138">Create two local resources R1 andR2, assign them to task T2 at 50% each, and choose  **getTaskAsync** again. The results in the text box include the resource information. If the task is in a synchronized SharePoint task list, the results also include the SharePoint task ID.</span></span>

    - <span data-ttu-id="127dc-221">Nome da tarefa: `T2`</span><span class="sxs-lookup"><span data-stu-id="127dc-221">Task name: `T2`</span></span>
    - <span data-ttu-id="127dc-222">GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span><span class="sxs-lookup"><span data-stu-id="127dc-222">GUID: `{25D3E03B-9A7D-E111-92FC-00155D3BA208}`</span></span>
    - <span data-ttu-id="127dc-223">Id da WSS: `0`</span><span class="sxs-lookup"><span data-stu-id="127dc-223">WSS Id: `0`</span></span>
    - <span data-ttu-id="127dc-224">ResourceNames: `R1[50%],R2[50%]`</span><span class="sxs-lookup"><span data-stu-id="127dc-224">ResourceNames: `R1[50%],R2[50%]`</span></span>

12. <span data-ttu-id="127dc-p139">Selecione o botão **Get Task Fields**. A função **getTaskFields** chama a função **getTaskfieldAsync** várias vezes para o nome da tarefa, o índice, a data de início, a duração, a prioridade e as anotações da tarefa.</span><span class="sxs-lookup"><span data-stu-id="127dc-p139">Select the **Get Task Fields** button. The **getTaskFields** function calls the **getTaskfieldAsync** function multiple times for the task name, index, start date, duration, priority, and task notes.</span></span>

    - <span data-ttu-id="127dc-227">Nome: `T2`</span><span class="sxs-lookup"><span data-stu-id="127dc-227">Name: `T2`</span></span>
    - <span data-ttu-id="127dc-228">ID: `2`</span><span class="sxs-lookup"><span data-stu-id="127dc-228">ID: `2`</span></span>
    - <span data-ttu-id="127dc-229">Início: `Thu 6/14/12`</span><span class="sxs-lookup"><span data-stu-id="127dc-229">Start: `Thu 6/14/12`</span></span>
    - <span data-ttu-id="127dc-230">Duração: `4d`</span><span class="sxs-lookup"><span data-stu-id="127dc-230">Duration: `4d`</span></span>
    - <span data-ttu-id="127dc-231">Prioridade: `500`</span><span class="sxs-lookup"><span data-stu-id="127dc-231">Priority: `500`</span></span>
    - <span data-ttu-id="127dc-232">Observações: essa é uma anotação de tarefa T2.</span><span class="sxs-lookup"><span data-stu-id="127dc-232">Notes: This is a note for task T2.</span></span> <span data-ttu-id="127dc-233">É apenas uma anotação de teste.</span><span class="sxs-lookup"><span data-stu-id="127dc-233">It is only a test note.</span></span> <span data-ttu-id="127dc-234">Se fosse uma anotação de verdade, teria algumas informações reais.</span><span class="sxs-lookup"><span data-stu-id="127dc-234">If it had been a real note, there would be some real information.</span></span>

13. <span data-ttu-id="127dc-p141">Selecione o botão **getWSSUrlAsync**. Se o projeto é um dos tipos a seguir, os resultados mostram a URL e o nome da lista de tarefas.</span><span class="sxs-lookup"><span data-stu-id="127dc-p141">Select the **getWSSUrlAsync** button. If the project is one of the following kinds, the results show the task list URL and name.</span></span>

    - <span data-ttu-id="127dc-237">Uma lista de tarefas do SharePoint importada no Project Server.</span><span class="sxs-lookup"><span data-stu-id="127dc-237">A SharePoint task list that was imported to Project Server.</span></span>
    - <span data-ttu-id="127dc-238">Uma lista de tarefas do SharePoint importada no Project Professional, depois salva novamente no SharePoint (sem usar o Project Server).</span><span class="sxs-lookup"><span data-stu-id="127dc-238">A SharePoint task list that was imported to Project Professional, and then saved back in SharePoint (not using Project Server).</span></span>

    > [!NOTE]
    > <span data-ttu-id="127dc-239">Se o Project Professional estiver instalado em um computador com Windows Server, para poder salvar o projeto de volta no SharePoint, use o **Gerenciador de Servidores** para adicionar o recurso **Experiência Desktop**.</span><span class="sxs-lookup"><span data-stu-id="127dc-239">If Project Professional is installed on a Windows Server computer, to be able to save the project back to SharePoint, you can use the  **Server Manager** to add the **Desktop Experience** feature.</span></span>

    <span data-ttu-id="127dc-240">Se o projeto for um projeto local, ou se você usar o Project Professional para abrir um projeto gerenciado pelo Project Server, o método **getWSSUrlAsync** mostrará um erro indefinido.</span><span class="sxs-lookup"><span data-stu-id="127dc-240">If the project is a local project, or if you use Project Professional to open a project that is managed by Project Server, the  **getWSSUrlAsync** method shows an undefined error.</span></span>

    - <span data-ttu-id="127dc-241">URL do SharePoint: `http://ServerName`</span><span class="sxs-lookup"><span data-stu-id="127dc-241">SharePoint URL: `http://ServerName`</span></span>
    - <span data-ttu-id="127dc-242">Nome da lista: `Test task list`</span><span class="sxs-lookup"><span data-stu-id="127dc-242">List name: `Test task list`</span></span>

14. <span data-ttu-id="127dc-p142">Selecione o botão **Adicionar** na seção **Evento TaskSelectionChanged**, que chama a função **manageTaskEventHandler** para registrar um evento alterado de seleção de tarefa e retorna `In onComplete function for addHandlerAsync Status: succeeded` na caixa de texto. Selecione uma tarefa diferente. A caixa de texto mostra `In task selection changed event handler`, que é o resultado da função de retorno de chamada para o evento de alteração de seleção de tarefa. Escolha o botão **Remover** para cancelar o registro do manipulador de eventos.</span><span class="sxs-lookup"><span data-stu-id="127dc-p142">Select the **Add** button in the **TaskSelectionChanged event** section, which calls the **manageTaskEventHandler** function to register a task selection changed event and returns `In onComplete function for addHandlerAsync Status: succeeded` in the text box. Select a different task; the text box shows `In task selection changed event handler`, which is the output of the callback function for the task selection changed event. Choose the  **Remove** button to unregister the event handler.</span></span>

15. <span data-ttu-id="127dc-p143">Para usar os métodos de recurso, primeiro selecione um modo de exibição, como **Folha de Recursos**, **Uso de Recursos** ou **Formulário de recursos** e selecione um recurso no modo de exibição. Escolha **getSelectedResourceAsync** para iniciar a variável **resourceGuid** e escolha **Get Resource Fields** a fim de chamar **getResourceFieldAsync** várias vezes para as propriedades do recurso. Você também pode adicionar ou remover o manipulador de eventos da alteração da seleção do recurso.</span><span class="sxs-lookup"><span data-stu-id="127dc-p143">To use the resource methods, first select a view such as  **Resource Sheet**,  **Resource Usage**, or  **Resource Form**, and then select a resource in that view. Choose  **getSelectedResourceAsync** to initialize the **resourceGuid** variable, and then choose **Get Resource Fields** to call **getResourceFieldAsync** multiple times for the resource properties. You can also add or remove the resource selection changed event handler.</span></span>

    - <span data-ttu-id="127dc-249">Nome do recurso: `R1`</span><span class="sxs-lookup"><span data-stu-id="127dc-249">Resource name: `R1`</span></span>
    - <span data-ttu-id="127dc-250">Custo: `$800.00`</span><span class="sxs-lookup"><span data-stu-id="127dc-250">Cost: `$800.00`</span></span>
    - <span data-ttu-id="127dc-251">Taxa padrão: `$50.00/h`</span><span class="sxs-lookup"><span data-stu-id="127dc-251">Standard Rate: `$50.00/h`</span></span>
    - <span data-ttu-id="127dc-252">Custo real: `$0.00`</span><span class="sxs-lookup"><span data-stu-id="127dc-252">Actual Cost: `$0.00`</span></span>
    - <span data-ttu-id="127dc-253">Trabalho real: `0h`</span><span class="sxs-lookup"><span data-stu-id="127dc-253">Actual Work: `0h`</span></span>
    - <span data-ttu-id="127dc-254">Unidades: `100%`</span><span class="sxs-lookup"><span data-stu-id="127dc-254">Units: `100%`</span></span>

16. <span data-ttu-id="127dc-p144">Selecione **getSelectedViewAsync** para exibir o tipo e o nome do modo de exibição ativo. Você também pode adicionar ou remover o manipulador de eventos da alteração da seleção de exibição. Por exemplo, se **Formulário de Recursos** é o modo de exibição ativo, a função **getSelectedViewAsync** mostra o seguinte na caixa de texto:</span><span class="sxs-lookup"><span data-stu-id="127dc-p144">Select **getSelectedViewAsync** to show the type and name of the active view. You can also add or remove the view selection changed event handler. For example, if **Resource Form** is the active view, the **getSelectedViewAsync** function shows the following in the text box:</span></span>

    - <span data-ttu-id="127dc-258">Tipo de exibição: `6`</span><span class="sxs-lookup"><span data-stu-id="127dc-258">View type: `6`</span></span>
    - <span data-ttu-id="127dc-259">Nome: `Resource Form`</span><span class="sxs-lookup"><span data-stu-id="127dc-259">Name: `Resource Form`</span></span>

17. <span data-ttu-id="127dc-p145">Selecione **Get Project Fields** para chamar a função **getProjectFieldAsync** várias vezes para propriedades diferentes do projeto ativo. Se o projeto é aberto do Project Web App, a função **getProjectFieldAsync** pode obter a URL da instância do Project Web App.</span><span class="sxs-lookup"><span data-stu-id="127dc-p145">Select **Get Project Fields** to call the **getProjectFieldAsync** function multiple times for different properties of the active project. If the project is opened from Project Web App, the **getProjectFieldAsync** function can get the URL of the Project Web App instance.</span></span>

    - <span data-ttu-id="127dc-262">GUID do projeto: `9845922E-DAB4-E111-8AF3-00155D3BA208`</span><span class="sxs-lookup"><span data-stu-id="127dc-262">Project GUID: `9845922E-DAB4-E111-8AF3-00155D3BA208`</span></span>
    - <span data-ttu-id="127dc-263">Início: `Tue 6/12/12`</span><span class="sxs-lookup"><span data-stu-id="127dc-263">Start: `Tue 6/12/12`</span></span>
    - <span data-ttu-id="127dc-264">Término: `Tue 6/19/12`</span><span class="sxs-lookup"><span data-stu-id="127dc-264">Finish: `Tue 6/19/12`</span></span>
    - <span data-ttu-id="127dc-265">Dígitos da moeda: `2`</span><span class="sxs-lookup"><span data-stu-id="127dc-265">Currency digits: `2`</span></span>
    - <span data-ttu-id="127dc-266">Símbolo da moeda: `$`</span><span class="sxs-lookup"><span data-stu-id="127dc-266">Currency symbol: `$`</span></span>
    - <span data-ttu-id="127dc-267">Posição do símbolo: `0`</span><span class="sxs-lookup"><span data-stu-id="127dc-267">Symbol position: `0`</span></span>
    - <span data-ttu-id="127dc-268">URL do Project Web App: `http://servername/pwa`</span><span class="sxs-lookup"><span data-stu-id="127dc-268">Project web app URL: `http://servername/pwa`</span></span>
  
18. <span data-ttu-id="127dc-p146">Selecione o botão **Get Context Values** para obter as propriedades do documento e o aplicativo no qual o suplemento está sendo executado, obtendo propriedades dos objetos **Office.Context.document** e **Office.context.application**. Por exemplo, se o arquivo Project1.mpp estiver na área de trabalho do computador local, a URL do documento será `C:\Users\UserAlias\Desktop\Project1.mpp`. Se o arquivo .mpp estiver em uma biblioteca do SharePoint, o valor será a URL do documento. Se você usar o Project Professional 2013 para abrir um projeto chamado Project1 do Project Web App, a URL do documento será `<>\Project1`.</span><span class="sxs-lookup"><span data-stu-id="127dc-p146">Select  the **Get Context Values** button get properties of the document and the application in which the add-in is running, by getting properties of the **Office.Context.document** object and the **Office.context.application** object. For example, if the Project1.mpp file is on the local computer desktop, the document URL is `C:\Users\UserAlias\Desktop\Project1.mpp`. If the .mpp file is in a SharePoint library, the value is the URL of the document. If you use Project Professional 2013 to open a project named Project1 from Project Web App, the document URL is  `<>\Project1`.</span></span>

    - <span data-ttu-id="127dc-273">URL do documento: `<>\Project1`</span><span class="sxs-lookup"><span data-stu-id="127dc-273">Document URL: `<>\Project1`</span></span>
    - <span data-ttu-id="127dc-274">Modo do documento: `readWrite`</span><span class="sxs-lookup"><span data-stu-id="127dc-274">Document mode: `readWrite`</span></span>
    - <span data-ttu-id="127dc-275">Idioma do aplicativo: `en-US`</span><span class="sxs-lookup"><span data-stu-id="127dc-275">App language: `en-US`</span></span>
    - <span data-ttu-id="127dc-276">Idioma de exibição: `en-US`</span><span class="sxs-lookup"><span data-stu-id="127dc-276">Display language: `en-US`</span></span>

19. <span data-ttu-id="127dc-p147">Você pode atualizar o suplemento após editar o código-fonte fechando e reiniciando o Project. Na faixa de opções **Projeto** a lista suspensa **Suplementos do Office** mantém a lista de suplementos usados recentemente.</span><span class="sxs-lookup"><span data-stu-id="127dc-p147">You can refresh the add-in after you edit the source code by closing and restarting Project. In the  **Project** ribbon, the **Office Add-ins** drop-down list maintains the list of recently used add-ins.</span></span>

## <a name="example"></a><span data-ttu-id="127dc-279">Exemplo</span><span class="sxs-lookup"><span data-stu-id="127dc-279">Example</span></span>

<span data-ttu-id="127dc-p148">O download do SDK do Project 2013 contém o código completo no arquivo JSOMCall.html, o arquivo JSOM_Sample.js e os arquivos Office.js, Office.debug.js, Project-15.js e Project-15.debug.js relacionados. Este é o código no arquivo JSOMCall.html.</span><span class="sxs-lookup"><span data-stu-id="127dc-p148">The Project 2013 SDK download contains the complete code in the JSOMCall.html file, the JSOM_Sample.js file, and the related Office.js, Office.debug.js, Project-15.js, and Project-15.debug.js files. Following is the code in the JSOMCall.html file.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <title>Project OM Sample Code</title>
        <meta http-equiv="X-UA-Compatible" content="IE=Edge"/>

        <script type="text/javascript" src="MicrosoftAjax.js"></script>

        <!-- Use the CDN reference to office.js when deploying your add-in. -->
        <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
        <script type="text/javascript" src="office.js"></script>
        <script type="text/javascript" src="JSOM_Sample.js"></script>

        <style type="text/css">
            .button-wide {
                width: 210px;
                margin-top: 2px;
            }
            .button-narrow 
            {
                width: 80px;
                margin-top: 2px;
            }
        </style>
    </head>

    <body>
        <div id="Common_JSOM_API">
            OBJECT MODEL TESTS
            <br /><br />
            <strong>General method:</strong>
            <br />
            <input id="Button5" class="button-wide" type="button" onclick="getSelectedDataAsync()" 
                value="getSelectedDataAsync" />
        </div>
        <div id="ProjectSpecificTask">
            <br />
            <strong>Project-specific task methods:</strong><br />
            <button class="button-wide" onclick="getSelectedTaskAsync()">getSelectedTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskAsync()">getTaskAsync</button><br />
            <button class="button-wide" onclick="getTaskFields()">Get Task Fields</button><br />
            <button class="button-wide" onclick="getWSSUrlAsync()">getWSSUrlAsync</button>
            <strong>Task selection changed:</strong>
            <button class="button-narrow" onclick="manageTaskEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageTaskEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ResourceMethods">
            <br />
            <strong>Resource methods:</strong>
            <button class="button-wide" onclick="getSelectedResourceAsync()">getSelectedResourceAsync</button><br />
            <button class="button-wide" onclick="getResourceFields()">Get Resource Fields</button><br />
            <strong>Resource selection changed:</strong>
            <button class="button-narrow" onclick="manageResourceEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageResourceEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ViewMethods">
            <br />
            <strong>View method:</strong>
            <button class="button-wide" onclick="getSelectedViewAsync()">getSelectedViewAsync</button><br />
            <strong>View selection changed:</strong>
            <button class="button-narrow" onclick="manageViewEventHandler('addHandlerAsync')">Add</button>
            <button class="button-narrow" onclick="manageViewEventHandler('removeHandlerAsync')">Remove</button>
        </div>
        <div id="ProjectMethods">
            <br />
            <strong>Project properties:</strong>
            <button class="button-wide" onclick="getProjectFields()">Get Project Fields</button><br />
        </div>
        <div id="ContextVariables">
            <br />
            <strong>Context properties:</strong>
            <button class="button-wide" onclick="getContextValues()">Get Context Values</button>
        </div>
        <br />
        <textarea id="text" rows="10" cols="25">This is the text result.</textarea>
    </body>
</html
```

## <a name="robust-programming"></a><span data-ttu-id="127dc-282">Programação robusta</span><span class="sxs-lookup"><span data-stu-id="127dc-282">Robust programming</span></span>

<span data-ttu-id="127dc-p149">O suplemento **Teste de modelo de objeto do Project** é um exemplo que mostra o uso de algumas funções JavaScript do Project 2013 nos arquivos Project-15.js e Office.js. O exemplo é somente para teste e não inclui verificações de erro robustas. Por exemplo, se você não selecionar um recurso e executar a função **getSelectedResourceAsync**, a variável **resourceGuid** não inicia e as chamadas para **getResourceFieldAsync** retornam um erro. Para um suplemento de produção, você deve verificar se há erros específicos e ignorar os resultados, ocultar funcionalidades que não se aplicam ou notificar o usuário para escolher um modo de exibição e fazer uma seleção válida antes de usar uma função.</span><span class="sxs-lookup"><span data-stu-id="127dc-p149">The  **Project OM Test** add-in is an example that shows the use of some JavaScript functions for Project 2013 in the Project-15.js and Office.js files. The example is for testing only and does not include robust error checks. For example, if you do not select a resource and run the **getSelectedResourceAsync** function, the **resourceGuid** variable is not initialized, and calls to **getResourceFieldAsync** return an error. For a production add-in, you should check for specific errors and ignore the results, hide functionality that does not apply, or notify the user to choose a view and make a valid selection before using a function.</span></span>

<span data-ttu-id="127dc-287">Para obter um exemplo simples, a saída de erro no código a seguir inclui a variável **actionMessage** que especifica a ação a tomar para evitar erros na função **getSelectedResourceAsync**.</span><span class="sxs-lookup"><span data-stu-id="127dc-287">For a simple example, the error output in the following code includes the  **actionMessage** variable that specifies the action to take to avoid an error in the **getSelectedResourceAsync** function.</span></span>

```js
function logError(errorText) {
    text.value = "Error in " + errorText;
}

function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);
}

// Get the GUID of the selected resource.
function getSelectedResourceAsync() {
    _projDoc.getSelectedResourceAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
            text.value = asyncResult.value;
            resourceGuid = asyncResult.value;
        }
        else {
            var actionMessage = "Select a resource before running the getSelectedResourceAsync method.";
            logMethodError("getSelectedResourceAsync", asyncResult.error.name,
                               asyncResult.error.message, actionMessage);
        }
    });
}
```

<span data-ttu-id="127dc-288">O exemplo **HelloProject_OData** no download do SDK do Project 2013 inclui o arquivo SurfaceErrors.js que usa a biblioteca JQuery para exibir uma mensagem de erro pop-up.</span><span class="sxs-lookup"><span data-stu-id="127dc-288">The **HelloProject_OData** sample in the Project 2013 SDK download includes the SurfaceErrors.js file that uses the JQuery library to display a pop-up error message.</span></span> <span data-ttu-id="127dc-289">A Figura 4 mostra a mensagem de erro em uma notificação do sistema.</span><span class="sxs-lookup"><span data-stu-id="127dc-289">Figure 4 shows the error message in a "toast" notification.</span></span>

<span data-ttu-id="127dc-290">O código a seguir no arquivo SurfaceErrors.js inclui a função **throwError** que cria um objeto **Toast**.</span><span class="sxs-lookup"><span data-stu-id="127dc-290">The following code in the SurfaceErrors.js file includes the  **throwError** function that creates a **Toast** object.</span></span>

```js
/*
 * Show error messages in a "toast" notification.
 */

// Throws a custom defined error.
function throwError(errTitle, errMessage) {
    try {
        // Define and throw a custom error.
        var customError = { name: errTitle, message: errMessage }
        throw customError;
    }
    catch (err) {
        // Catch the error and display it to the user.
        Toast.showToast(err.name, err.message);
    }
}

// Add a dynamically-created div "toast" for displaying errors to the user.
var Toast = {

    Toast: "divToast",
    Close: "btnClose",
    Notice: "lblNotice",
    Output: "lblOutput",

    // Show the toast with the specified information.
    showToast: function (title, message) {

        if (document.getElementById(this.Toast) == null) {
            this.createToast();
        }

        document.getElementById(this.Notice).innerText = title;
        document.getElementById(this.Output).innerText = message;

        $("#" + this.Toast).hide();
        $("#" + this.Toast).show("slow");
    },

    // Create the display for the toast.
    createToast: function () {
        var divToast;
        var lblClose;
        var btnClose;
        var divOutput;
        var lblOutput;
        var lblNotice;

        // Create the container div.
        divToast = document.createElement("div");
        var toastStyle = "background-color:rgba(220, 220, 128, 0.80);" +
            "position:absolute;" +
            "bottom:0px;" +
            "width:90%;" +
            "text-align:center;" +
            "font-size:11pt;";
        divToast.setAttribute("style", toastStyle);
        divToast.setAttribute("id", this.Toast);

        // Create the close button.
        lblClose = document.createElement("div");
        lblClose.setAttribute("id", this.Close);
        var btnStyle = "text-align:right;" +
            "padding-right:10px;" +
            "font-size:10pt;" +
            "cursor:default";
        lblClose.setAttribute("style", btnStyle);
        lblClose.appendChild(document.createTextNode("CLOSE "));

        btnClose = document.createElement("span");
        btnClose.setAttribute("style", "cursor:pointer;");
        btnClose.setAttribute("onclick", "Toast.close()");
        btnClose.innerText = "X";
        lblClose.appendChild(btnClose);

        // Create the div to contain the toast title and message.
        divOutput = document.createElement("div");
        divOutput.setAttribute("id", "divOutput");
        var outputStyle = "margin-top:0px;";
        divOutput.setAttribute("style", outputStyle);

        lblNotice = document.createElement("span");
        lblNotice.setAttribute("id", this.Notice);
        var labelStyle = "font-weight:bold;margin-top:0px;";
        lblNotice.setAttribute("style", labelStyle);

        lblOutput = document.createElement("span");
        lblOutput.setAttribute("id", this.Output);

        // Add the child nodes to the toast div.
        divOutput.appendChild(lblNotice);
        divOutput.appendChild(document.createElement("br"));
        divOutput.appendChild(lblOutput);
        divToast.appendChild(lblClose);
        divToast.appendChild(divOutput);

        // Add the toast div to the document body.
        document.body.appendChild(divToast);
    },

    // Close the toast.
    close: function () {
        $("#" + this.Toast).hide("slow");
    }
}
```

<span data-ttu-id="127dc-291">Para usar a função **throwError**, inclua a biblioteca JQuery e o script SurfaceErrors.js no arquivo JSOMCall.html e adicione uma chamada para **throwError** em outras funções JavaScript, como **logMethodError**.</span><span class="sxs-lookup"><span data-stu-id="127dc-291">To use the  **throwError** function, include the JQuery library and the SurfaceErrors.js script in the JSOMCall.html file, and then add a call to **throwError** in other JavaScript functions such as **logMethodError**.</span></span>

> [!NOTE]
> <span data-ttu-id="127dc-p151">Antes de implantar o suplemento, mude a referência office.js e a referência jQuery para a referência CDN (rede de distribuição de conteúdo). A referência CDN fornece a versão mais recente e melhora o desempenho.</span><span class="sxs-lookup"><span data-stu-id="127dc-p151">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

```HTML
<!DOCTYPE html>
<html>
<head>
    <title>Project OM Sample Code</title>
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <script type="text/javascript" src="MicrosoftAjax.js"></script>

    <!-- Use the CDN reference to Office.js and jQuery when deploying your add-in. -->
    <!-- <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> -->
    <script type="text/javascript" src="office.js"></script>
    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jQuery/jquery-1.9.0.min.js"></script>

    <script type="text/javascript" src="JSOM_Sample.js"></script>
    <script type="text/javascript" src="SurfaceErrors.js"></script>

    <!-- . . . INVALID USE OF SYMBOLS . . . -->
</head>

```

<br/>

```js
function logMethodError(methodName, errorName, errorMessage, actionMessage) {
    logError(methodName + " method.\nError name: " + errorName
        + "\nMessage: " + errorMessage
        + "\n\nAction: " + actionMessage);

    throwError(methodName + " error", actionMessage);
}
```

<br/>

<span data-ttu-id="127dc-294">*Figura 4. Funções no arquivo SurfaceErrors.js podem mostrar uma notificação "toast"*</span><span class="sxs-lookup"><span data-stu-id="127dc-294">*Figure 4. Functions in the SurfaceErrors.js file can show a "toast" notification*</span></span>

![Usando as rotinas do SurfaceError para mostrar um erro](../images/pj15-create-simple-agave-surface-error.png)


## <a name="see-also"></a><span data-ttu-id="127dc-296">Confira também</span><span class="sxs-lookup"><span data-stu-id="127dc-296">See also</span></span>

- [<span data-ttu-id="127dc-297">Suplementos do painel de tarefas para Project</span><span class="sxs-lookup"><span data-stu-id="127dc-297">Task pane add-ins for Project</span></span>](../project/project-add-ins.md)
- [<span data-ttu-id="127dc-298">Noções básicas da API JavaScript para suplementos</span><span class="sxs-lookup"><span data-stu-id="127dc-298">Understanding the JavaScript API for add-ins</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="127dc-299">API JavaScript para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="127dc-299">JavaScript API for Office Add-ins</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="127dc-300">Referência de esquema para manifestos de suplementos do Office (versão 1.1)</span><span class="sxs-lookup"><span data-stu-id="127dc-300">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="127dc-301">Download do SDK do Project 2013</span><span class="sxs-lookup"><span data-stu-id="127dc-301">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
