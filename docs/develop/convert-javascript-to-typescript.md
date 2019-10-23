---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: ''
ms.date: 10/11/2019
localization_priority: Priority
ms.openlocfilehash: 0a828a3f11a1fcaf71e277bdb667f866ea4ae06a
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626800"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="837f1-102">Converter um projeto de Suplemento do Office no Visual Studio para TypeScript</span><span class="sxs-lookup"><span data-stu-id="837f1-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="837f1-103">Você pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="837f1-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="837f1-104">Este artigo descreve o processo de conversão de um suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="837f1-104">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="837f1-105">Você pode usar o mesmo processo para converter outros tipos de projetos de suplementos do Office de JavaScript para TypeScript no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="837f1-105">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="837f1-106">Para criar um projeto de suplementos TypeScript do Office sem usar o Visual Studio, siga as instruções na seção "Gerador do Yeoman" de um [início rápido em 5 minutos](../index.md) e escolha `TypeScript` quando for solicitado pelo [Gerador de suplementos do Office do Yeoman](https://github.com/officedev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="837f1-106">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/officedev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="837f1-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="837f1-107">Prerequisites</span></span>

- <span data-ttu-id="837f1-108">[Visual Studio 2019](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada</span><span class="sxs-lookup"><span data-stu-id="837f1-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="837f1-109">Se você já instalou o Visual Studio 2019, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada.</span><span class="sxs-lookup"><span data-stu-id="837f1-109">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="837f1-110">Se essa carga de trabalho ainda não estiver instalada, use o instalador do Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="837f1-110">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="837f1-111">TypeScript SDK versão 2.3 ou posterior (para o Visual Studio 2019)</span><span class="sxs-lookup"><span data-stu-id="837f1-111">TypeScript SDK version 2.3 or later (for Visual Studio 2017)</span></span>

    > [!TIP]
    > <span data-ttu-id="837f1-112">No [Instalador do Visual Studio](/visualstudio/install/modify-visual-studio), selecione a guia **Componentes individuais** e role a tela para baixo até a seção **SDKs, bibliotecas e estruturas**.</span><span class="sxs-lookup"><span data-stu-id="837f1-112">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="837f1-113">Nessa seção, verifique se pelo menos um dos componentes do **SDK do TypeScript** (versão 2.3 ou posterior) está selecionado.</span><span class="sxs-lookup"><span data-stu-id="837f1-113">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="837f1-114">Se nenhum dos componentes do **SDK do TypeScript** estiver selecionado, selecione a versão mais recente do SDK disponível e, em seguida, escolha o botão **Modificar** para [instalar esse componente individual](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="837f1-114">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="837f1-115">Excel 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="837f1-115">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="837f1-116">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="837f1-116">Create the add-in project</span></span>

1. <span data-ttu-id="837f1-117">No Visual Studio, escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="837f1-117">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="837f1-118">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="837f1-118">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="837f1-119">Escolha \*\*Suplemento Excel Web \*\*, em seguida, selecione **Próximo**.</span><span class="sxs-lookup"><span data-stu-id="837f1-119">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="837f1-120">Nomeie seu projeto e selecione **Criar**.</span><span class="sxs-lookup"><span data-stu-id="837f1-120">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="837f1-121">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="837f1-121">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="837f1-p105">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="837f1-p105">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="837f1-124">Converter o projeto do suplemento para TypeScript</span><span class="sxs-lookup"><span data-stu-id="837f1-124">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="837f1-125">Localize o arquivo **Home.js** e o renomeie para **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="837f1-125">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="837f1-126">Na guia **Ferramentas**, escolha **Gerenciador de Pacotes NuGet** e, em seguida, selecione **Gerenciar Pacotes do NuGet para Solução...**.</span><span class="sxs-lookup"><span data-stu-id="837f1-126">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

3. <span data-ttu-id="837f1-127">Com a guia **Navegar** selecionada, insira **office-js.TypeScript.DefinitelyTyped** na caixa de pesquisa.</span><span class="sxs-lookup"><span data-stu-id="837f1-127">With the **Browse** tab selected, enter **office-js.TypeScript.DefinitelyTyped** into the search box.</span></span> <span data-ttu-id="837f1-128">Instalar ou atualizar esse pacote se ele já estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="837f1-128">Install or update this package if it is already installed.</span></span> <span data-ttu-id="837f1-129">Isso adicionará as definições de tipo de TypeScript da biblioteca do Office.js ao seu projeto.</span><span class="sxs-lookup"><span data-stu-id="837f1-129">This will add the TypeScript type definitions for the Office.js library to your project.</span></span>

4. <span data-ttu-id="837f1-130">Na mesma caixa de pesquisa, digite **jquery.TypeScript.DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="837f1-130">In the same search box, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="837f1-131">Instalar ou atualizar esse pacote se ele já estiver instalado.</span><span class="sxs-lookup"><span data-stu-id="837f1-131">Install or update this package if it is already installed.</span></span> <span data-ttu-id="837f1-132">Isso adicionará as definições do TypeScript jQuery ao seu projeto.</span><span class="sxs-lookup"><span data-stu-id="837f1-132">This will add the jQuery TypeScript definitions into your project.</span></span> <span data-ttu-id="837f1-133">Os pacotes do jQuery e do Office.js agora serão exibidos em um novo arquivo gerado pelo Visual Studio, chamado **packages.config**.</span><span class="sxs-lookup"><span data-stu-id="837f1-133">The packages for both jQuery and Office.js will now appear in a new file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="837f1-p108">Em seu projeto em TypeScript, você pode ter uma combinação de arquivos TypeScript e JavaScript e seu projeto irá compilar. Isso ocorre porque o TypeScript é um superconjunto tipado do JavaScript que compila o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="837f1-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

5. <span data-ttu-id="837f1-136">Abra o arquivo **Home.ts** e adicione a seguinte declaração à parte superior do arquivo:</span><span class="sxs-lookup"><span data-stu-id="837f1-136">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```TypeScript
    declare var fabric: any;
    ```

6. <span data-ttu-id="837f1-137">Em **Home.ts**, exclua a linha `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` e substitua pela seguinte:</span><span class="sxs-lookup"><span data-stu-id="837f1-137">In **Home.ts**, delete the line `if(!Office.context.requirements.isSetSupported('ExcelApi', '1.1') {` and replace with the following:</span></span>

    ```TypeScript
    if(!Office.context.requirements.isSetSupported('ExcelApi', 1.1) {
    ```

7. <span data-ttu-id="837f1-138">No arquivo **Home.ts** localize a linha `Office.initialize = function (reason) {` e adicione uma linha imediatamente após ela para usar o polyfill na `window.Promise` global, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="837f1-138">In the **Home.ts** file, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

8. <span data-ttu-id="837f1-139">No arquivo **Home.ts**, localize a função `displaySelectedCells`, substitua toda a função pelo código a seguir e salve o arquivo:</span><span class="sxs-lookup"><span data-stu-id="837f1-139">In the **Home.ts** file, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="837f1-140">Executar o projeto do suplemento convertido</span><span class="sxs-lookup"><span data-stu-id="837f1-140">Run the converted add-in project</span></span>

1. <span data-ttu-id="837f1-p109">No Visual Studio, pressione **F5** ou clique no botão **Iniciar** para iniciar o Excel com o botão do suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento estará hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="837f1-p109">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="837f1-143">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="837f1-143">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="837f1-144">Na planilha, selecione as nove células que contêm números.</span><span class="sxs-lookup"><span data-stu-id="837f1-144">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="837f1-145">Pressione o botão **Realçar** no painel de tarefas para realçar a célula no intervalo selecionado com o maior valor.</span><span class="sxs-lookup"><span data-stu-id="837f1-145">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="837f1-146">Arquivo de código Home.ts</span><span class="sxs-lookup"><span data-stu-id="837f1-146">Home.ts code file</span></span>

<span data-ttu-id="837f1-p110">Para sua referência o trecho de código a seguir mostra o conteúdo do arquivo **Home.ts** após a aplicação das alterações descritas anteriormente. Esse código contém o número mínimo de alterações necessárias para que seu suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="837f1-p110">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```typescript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;

        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
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

## <a name="see-also"></a><span data-ttu-id="837f1-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="837f1-149">See also</span></span>

* [<span data-ttu-id="837f1-150">Discussão de implementação do Promise no StackOverflow</span><span class="sxs-lookup"><span data-stu-id="837f1-150">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="837f1-151">Exemplos de Suplementos do Office no GitHub</span><span class="sxs-lookup"><span data-stu-id="837f1-151">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
