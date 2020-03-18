---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: Saiba como converter um projeto de suplemento do Office no Visual Studio para usar TypeScript.
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 1dbb3503a521f1a7c3e71764a50f02708b667a11
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719039"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="fbba3-103">Converter um projeto de Suplemento do Office no Visual Studio para TypeScript</span><span class="sxs-lookup"><span data-stu-id="fbba3-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="fbba3-104">Você pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="fbba3-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="fbba3-105">Este artigo descreve o processo de conversão de um suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="fbba3-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="fbba3-106">Você pode usar o mesmo processo para converter outros tipos de projetos de suplementos do Office de JavaScript para TypeScript no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="fbba3-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="fbba3-107">Para criar um projeto de suplementos TypeScript do Office sem usar o Visual Studio, siga as instruções na seção "Gerador do Yeoman" de um [início rápido em 5 minutos](../index.md) e escolha `TypeScript` quando for solicitado pelo [Gerador de suplementos do Office do Yeoman](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="fbba3-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="fbba3-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="fbba3-108">Prerequisites</span></span>

- <span data-ttu-id="fbba3-109">[Visual Studio 2019](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada</span><span class="sxs-lookup"><span data-stu-id="fbba3-109">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="fbba3-110">Se você já instalou o Visual Studio 2019, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada.</span><span class="sxs-lookup"><span data-stu-id="fbba3-110">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="fbba3-111">Se essa carga de trabalho ainda não estiver instalada, use o instalador do Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="fbba3-111">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="fbba3-112">TypeScript SDK versão 2.3 ou posterior (para o Visual Studio 2019)</span><span class="sxs-lookup"><span data-stu-id="fbba3-112">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="fbba3-113">No [Instalador do Visual Studio](/visualstudio/install/modify-visual-studio), selecione a guia **Componentes individuais** e role a tela para baixo até a seção **SDKs, bibliotecas e estruturas**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-113">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="fbba3-114">Nessa seção, verifique se pelo menos um dos componentes do **SDK do TypeScript** (versão 2.3 ou posterior) está selecionado.</span><span class="sxs-lookup"><span data-stu-id="fbba3-114">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="fbba3-115">Se nenhum dos componentes do **SDK do TypeScript** estiver selecionado, selecione a versão mais recente do SDK disponível e, em seguida, escolha o botão **Modificar** para [instalar esse componente individual](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="fbba3-115">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="fbba3-116">Excel 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="fbba3-116">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="fbba3-117">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="fbba3-117">Create the add-in project</span></span>

1. <span data-ttu-id="fbba3-118">No Visual Studio, escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-118">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="fbba3-119">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-119">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="fbba3-120">Escolha \*\*suplemento do Excel Web \*\*, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-120">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="fbba3-121">Nomeie seu projeto e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-121">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="fbba3-122">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="fbba3-122">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="fbba3-p105">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="fbba3-p105">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="fbba3-125">Converter o projeto do suplemento para TypeScript</span><span class="sxs-lookup"><span data-stu-id="fbba3-125">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="fbba3-126">Localize o arquivo **Home.js** e o renomeie para **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-126">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="fbba3-127">Localize o arquivo **./Functions/FunctionFile.js** e renomeie-o para **FunctionFile.ts**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-127">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="fbba3-128">Localize o arquivo **./Scripts/MessageBanner.js** e renomeie-o para **MessageBanner.ts**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-128">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="fbba3-129">Na guia **Ferramentas**, escolha **Gerenciador de Pacotes NuGet** e, em seguida, selecione **Gerenciar Pacotes do NuGet para Solução...**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-129">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="fbba3-130">Com a guia **Navegar** selecionada, insira **office-js.TypeScript.DefinitelyTyped** na caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="fbba3-130">With the **Browse** tab selected, enter **office-js.TypeScript.DefinitelyTyped** into the search box.</span></span> <span data-ttu-id="fbba3-131">Instalar ou atualizar esse pacote se ele já estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="fbba3-131">Install or update this package if it is already installed.</span></span> <span data-ttu-id="fbba3-132">Isso adicionará as definições de tipo de TypeScript da biblioteca do Office.js ao seu projeto.</span><span class="sxs-lookup"><span data-stu-id="fbba3-132">This will add the TypeScript type definitions for the Office.js library to your project.</span></span>

6. <span data-ttu-id="fbba3-133">Na mesma caixa de pesquisa, digite **jquery.TypeScript.DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-133">In the same search box, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="fbba3-134">Instalar ou atualizar esse pacote se ele já estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="fbba3-134">Install or update this package if it is already installed.</span></span> <span data-ttu-id="fbba3-135">Isso adicionará as definições do TypeScript jQuery ao seu projeto.</span><span class="sxs-lookup"><span data-stu-id="fbba3-135">This will add the jQuery TypeScript definitions into your project.</span></span> <span data-ttu-id="fbba3-136">Os pacotes do jQuery e do Office.js agora serão exibidos em um novo arquivo gerado pelo Visual Studio, chamado **packages.config**.</span><span class="sxs-lookup"><span data-stu-id="fbba3-136">The packages for both jQuery and Office.js will now appear in a new file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="fbba3-p108">Em seu projeto em TypeScript, você pode ter uma combinação de arquivos TypeScript e JavaScript e seu projeto irá compilar. Isso ocorre porque o TypeScript é um superconjunto tipado do JavaScript que compila o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="fbba3-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

7. <span data-ttu-id="fbba3-139">Em **Home.ts**, localize a linha `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` e substitua-a pelo seguinte:</span><span class="sxs-lookup"><span data-stu-id="fbba3-139">In **Home.ts**, find the line `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` and replace it with the following:</span></span>

    ```TypeScript
    if(!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
    ```

    > [!NOTE]
    > <span data-ttu-id="fbba3-140">Atualmente, para que o projeto seja compilado com êxito depois de convertido para TypeScript, você deve especificar o número do conjunto de requisitos como um valor numérico, conforme mostrado no trecho de código anterior.</span><span class="sxs-lookup"><span data-stu-id="fbba3-140">Currently, for the project to compile successfully after it's converted to TypeScript, you must specify the requirement set number as a numeric value as shown in the previous code snippet.</span></span> <span data-ttu-id="fbba3-141">Infelizmente, isso significa que você não poderá usar `isSetSupported` para testar o suporte ao conjunto de requisitos `1.10`, pois o valor numérico `1.10` retorna `1.1` em tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="fbba3-141">Unfortunately this means you'll be unable to use `isSetSupported` to test for requirement set `1.10` support, as the numeric value `1.10` evaluates to `1.1` at runtime.</span></span> 
    > 
    > <span data-ttu-id="fbba3-142">Esse problema ocorre devido ao pacote NuGet **office-js.TypeScript.DefinitelyTyped** se encontrar desatualizado, o que significa que o seu projeto não tem acesso às definições TypeScript mais recentes para o Office.js.</span><span class="sxs-lookup"><span data-stu-id="fbba3-142">This problem is due to the **office-js.TypeScript.DefinitelyTyped** NuGet package currently being outdated, which means your project doesn't have access to the latest TypeScript definitions for Office.js.</span></span> <span data-ttu-id="fbba3-143">Esse problema está sendo solucionado e este artigo será atualizado quando o problema for resolvido.</span><span class="sxs-lookup"><span data-stu-id="fbba3-143">This issue is being addressed and this article will be updated when the issue is resolved.</span></span>

8. <span data-ttu-id="fbba3-144">Em **Home.ts**, localize a linha `Office.initialize = function (reason) {` e adicione uma linha imediatamente depois para fazer polyfill do `window.Promise` global, como mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="fbba3-144">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

9. <span data-ttu-id="fbba3-145">Em **Home.ts**, localize a função `displaySelectedCells`, substitua a função inteira pelo código a seguir e, em seguida, salve o arquivo:</span><span class="sxs-lookup"><span data-stu-id="fbba3-145">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

    ```TypeScript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }
    ```

10. <span data-ttu-id="fbba3-146">Em **./Scripts/MessageBanner.ts**, localize a linha `_onResize(null);` e substitua-a pelo seguinte:</span><span class="sxs-lookup"><span data-stu-id="fbba3-146">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="fbba3-147">Executar o projeto do suplemento convertido</span><span class="sxs-lookup"><span data-stu-id="fbba3-147">Run the converted add-in project</span></span>

1. <span data-ttu-id="fbba3-p111">No Visual Studio, pressione **F5** ou clique no botão **Iniciar** para iniciar o Excel com o botão do suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento estará hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="fbba3-p111">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="fbba3-150">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="fbba3-150">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="fbba3-151">Na planilha, selecione as nove células que contêm números.</span><span class="sxs-lookup"><span data-stu-id="fbba3-151">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="fbba3-152">Pressione o botão **Realçar** no painel de tarefas para realçar a célula no intervalo selecionado com o maior valor.</span><span class="sxs-lookup"><span data-stu-id="fbba3-152">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="fbba3-153">Arquivo de código Home.ts</span><span class="sxs-lookup"><span data-stu-id="fbba3-153">Home.ts code file</span></span>

<span data-ttu-id="fbba3-p112">Para sua referência o trecho de código a seguir mostra o conteúdo do arquivo **Home.ts** após a aplicação das alterações descritas anteriormente. Esse código contém o número mínimo de alterações necessárias para que seu suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="fbba3-p112">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(highlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function highlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```

## <a name="see-also"></a><span data-ttu-id="fbba3-156">Confira também</span><span class="sxs-lookup"><span data-stu-id="fbba3-156">See also</span></span>

- [<span data-ttu-id="fbba3-157">Discussão de implementação do Promise no StackOverflow</span><span class="sxs-lookup"><span data-stu-id="fbba3-157">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="fbba3-158">Exemplos de Suplementos do Office no GitHub</span><span class="sxs-lookup"><span data-stu-id="fbba3-158">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
