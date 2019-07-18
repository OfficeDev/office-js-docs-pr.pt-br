---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: ''
ms.date: 07/17/2019
localization_priority: Priority
ms.openlocfilehash: 7c51479c1a5d1df5d9b0622dbae4fe9f01ad0c2c
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771303"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="8f0ab-102">Converter um projeto de Suplemento do Office no Visual Studio para TypeScript</span><span class="sxs-lookup"><span data-stu-id="8f0ab-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="8f0ab-103">Você pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="8f0ab-104">Este artigo descreve o processo de conversão de um suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-104">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="8f0ab-105">Você pode usar o mesmo processo para converter outros tipos de projetos de suplementos do Office de JavaScript para TypeScript no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-105">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="8f0ab-106">Para criar um projeto de suplementos TypeScript do Office sem usar o Visual Studio, siga as instruções na seção "Gerador do Yeoman" de um [início rápido em 5 minutos](../index.md) e escolha `TypeScript` quando for solicitado pelo [Gerador de suplementos do Office do Yeoman](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="8f0ab-106">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quick start](../index.md) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="8f0ab-107">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="8f0ab-107">Prerequisites</span></span>

- <span data-ttu-id="8f0ab-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada</span><span class="sxs-lookup"><span data-stu-id="8f0ab-108">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="8f0ab-109">Se você já instalou o Visual Studio 2017, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-109">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="8f0ab-110">Se essa carga de trabalho ainda não estiver instalada, use o instalador do Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="8f0ab-110">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).</span></span>

- <span data-ttu-id="8f0ab-111">SDK do TypeScript versão 2.3 ou posterior (para Visual Studio 2017)</span><span class="sxs-lookup"><span data-stu-id="8f0ab-111">TypeScript SDK version 2.3 or later (for Visual Studio 2017)</span></span>

    > [!TIP]
    > <span data-ttu-id="8f0ab-112">No [Instalador do Visual Studio](/visualstudio/install/modify-visual-studio), selecione a guia **Componentes individuais** e role a tela para baixo até a seção **SDKs, bibliotecas e estruturas**.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-112">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="8f0ab-113">Nessa seção, verifique se pelo menos um dos componentes do **SDK do TypeScript** (versão 2.3 ou posterior) está selecionado.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-113">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="8f0ab-114">Se nenhum dos componentes do **SDK do TypeScript** estiver selecionado, selecione a versão mais recente do SDK disponível e, em seguida, escolha o botão **Modificar** para [instalar esse componente individual](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="8f0ab-114">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components).</span></span> 

- <span data-ttu-id="8f0ab-115">Excel 2016 ou posterior</span><span class="sxs-lookup"><span data-stu-id="8f0ab-115">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="8f0ab-116">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="8f0ab-116">Create the add-in project</span></span>

1. <span data-ttu-id="8f0ab-117">Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-117">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="8f0ab-118">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-118">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="8f0ab-119">Dê um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-119">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="8f0ab-120">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-120">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="8f0ab-p104">O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="8f0ab-123">Converter o projeto do suplemento para TypeScript</span><span class="sxs-lookup"><span data-stu-id="8f0ab-123">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="8f0ab-124">No **Gerenciador de Soluções**, renomeie o arquivo **Home.js** como **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-124">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="8f0ab-p105">Em seu projeto em TypeScript, você pode ter uma combinação de arquivos TypeScript e JavaScript e seu projeto irá compilar. Isso ocorre porque o TypeScript é um superconjunto tipado do JavaScript que compila o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="8f0ab-127">Selecione **Sim** para confirmar que você deseja alterar a extensão do nome de arquivo.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-127">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="8f0ab-128">Crie um novo arquivo chamado **Office.d.ts** na raiz do projeto de aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-128">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="8f0ab-p106">No navegador, abra o [arquivo de definições de tipo para o Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copie o conteúdo do arquivo para a área de transferência.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p106">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="8f0ab-131">No Visual Studio, abra o arquivo **Office.d.ts**, cole o conteúdo de sua área de transferência de arquivo e salve-o.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-131">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="8f0ab-132">Crie um novo arquivo chamado **jQuery.d.ts** na raiz do projeto de aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-132">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="8f0ab-p107">No navegador, abra o [arquivo de definições de tipos para jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts). Copie o conteúdo do arquivo para a área de transferência.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p107">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts). Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="8f0ab-135">No Visual Studio, abra o arquivo **jQuery.d.ts**, cole o conteúdo de sua área de transferência nesse arquivo e salve-o.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-135">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="8f0ab-136">No Visual Studio, crie um novo arquivo chamado **tsconfig.json** na raiz do projeto de aplicativo web.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-136">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="8f0ab-137">Abra o arquivo **tsconfig.json**, adicione o conteúdo a seguir no arquivo e salve-o:</span><span class="sxs-lookup"><span data-stu-id="8f0ab-137">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```json
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ],
            "sourceMap": true
        }
    }
    ```

11. <span data-ttu-id="8f0ab-138">Abra o arquivo **Home.ts** e adicione a seguinte declaração à parte superior do arquivo:</span><span class="sxs-lookup"><span data-stu-id="8f0ab-138">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```typescript
    declare var fabric: any;
    ```

12. <span data-ttu-id="8f0ab-139">No arquivo **Home.ts** localize a linha `Office.initialize = function (reason) {` e adicione uma linha imediatamente após ela para metaabastecimento global `window.Promise`, conforme mostrado aqui:</span><span class="sxs-lookup"><span data-stu-id="8f0ab-139">In the **Home.ts** file, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```typescript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

13. <span data-ttu-id="8f0ab-140">No arquivo **Home.ts**, altere **'1.1'** para **1.1** (ou seja, remova as aspas) na seguinte linha:</span><span class="sxs-lookup"><span data-stu-id="8f0ab-140">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line:</span></span>

    ```typescript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

14. <span data-ttu-id="8f0ab-141">No arquivo **Home.ts**, localize a função `displaySelectedCells`, substitua toda a função pelo código a seguir e salve o arquivo:</span><span class="sxs-lookup"><span data-stu-id="8f0ab-141">In the **Home.ts** file, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

    ```typescript
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
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="8f0ab-142">Executar o projeto do suplemento convertido</span><span class="sxs-lookup"><span data-stu-id="8f0ab-142">Run the converted add-in project</span></span>

1. <span data-ttu-id="8f0ab-p108">No Visual Studio, pressione **F5** ou clique no botão **Iniciar** para iniciar o Excel com o botão do suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento estará hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p108">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="8f0ab-145">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-145">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="8f0ab-146">Na planilha, selecione as nove células que contêm números.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-146">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="8f0ab-147">Pressione o botão **Realçar** no painel de tarefas para realçar a célula no intervalo selecionado com o maior valor.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-147">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="8f0ab-148">Arquivo de código Home.ts</span><span class="sxs-lookup"><span data-stu-id="8f0ab-148">Home.ts code file</span></span>

<span data-ttu-id="8f0ab-p109">Para sua referência o trecho de código a seguir mostra o conteúdo do arquivo **Home.ts** após a aplicação das alterações descritas anteriormente. Esse código contém o número mínimo de alterações necessárias para que seu suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="8f0ab-p109">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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
            $('#highlight-button').click(hightlightHighestValue);
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

    function hightlightHighestValue() {
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

## <a name="see-also"></a><span data-ttu-id="8f0ab-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="8f0ab-151">See also</span></span>

* [<span data-ttu-id="8f0ab-152">Discussão de implementação do Promise no StackOverflow</span><span class="sxs-lookup"><span data-stu-id="8f0ab-152">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="8f0ab-153">Exemplos de Suplementos do Office no GitHub</span><span class="sxs-lookup"><span data-stu-id="8f0ab-153">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
