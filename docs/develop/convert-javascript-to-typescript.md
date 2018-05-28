---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 05e845b9d085b64b0534d28053dcd5ca3c7b403e
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/25/2018
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="2c6f4-102">Converter um projeto de Suplemento do Office no Visual Studio para TypeScript</span><span class="sxs-lookup"><span data-stu-id="2c6f4-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="2c6f4-103">Voc? pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="2c6f4-104">Use o Visual Studio para criar projeto suplemento, evite ter que criar desde o in?cio o projeto de suplemento do Office no TypeScript.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="2c6f4-105">Este artigo mostra como criar um suplemento do Excel usando o Visual Studio e depois converter o projeto do suplemento do JavaScript para o TypeScript.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="2c6f4-106">Voc? pode usar o mesmo processo para converter outros tipos de projetos de JavaScript para Suplementos do Office para o TypeScript no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="2c6f4-107">Para criar um projeto TypeScript do Suplemento do Office sem usar o Visual Studio, siga as instru??es na se??o "Qualquer editor" de qualquer [In?cio r?pido de 5 minutos](../index.yml) e escolha `TypeScript` quando solicitado pelo [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="2c6f4-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2c6f4-108">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="2c6f4-108">Prerequisites</span></span>

- <span data-ttu-id="2c6f4-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada</span><span class="sxs-lookup"><span data-stu-id="2c6f4-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="2c6f4-110">Se voc? j? instalou o Visual Studio 2017, [use o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="2c6f4-111">TypeScript 2.3 para Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="2c6f4-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="2c6f4-112">O TypeScript deve ser instalado por padr?o com o Visual Studio 2017, mas voc? pode [usar o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para confirmar se ele foi instalado.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="2c6f4-113">No Instalador do Visual Studio, selecione a guia **Componentes individuais** e verifique se a op??o**TypeScript 2.3 SDK** est? selecionada em **SDKs, bibliotecas e estruturas**.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="2c6f4-114">Excel 2016</span><span class="sxs-lookup"><span data-stu-id="2c6f4-114">Excel 2016</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="2c6f4-115">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="2c6f4-115">Create the add-in project</span></span>

1. <span data-ttu-id="2c6f4-116">Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo**  >  **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="2c6f4-117">Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a op??o **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="2c6f4-118">D? um nome ao projeto e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="2c6f4-119">Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="2c6f4-p104">O Visual Studio cria uma solu??o, e os dois projetos dele s?o exibidos no **Gerenciador de Solu??es**. O arquivo **Home.html** ? aberto no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="2c6f4-122">Converter o projeto do suplemento para TypeScript</span><span class="sxs-lookup"><span data-stu-id="2c6f4-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="2c6f4-123">No **Gerenciador de Solu??es**, renomeie o arquivo **Home.js** como **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="2c6f4-p105">Em seu projeto em TypeScript, voc? pode ter uma combina??o de arquivos TypeScript e JavaScript e seu projeto ir? compilar. Isso ocorre porque o TypeScript ? um superconjunto tipado do JavaScript que compila o JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="2c6f4-126">Selecione **Sim** para confirmar que voc? deseja alterar a extens?o do nome de arquivo.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="2c6f4-127">Crie um novo arquivo chamado **Office.d.ts** na raiz do projeto de aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="2c6f4-128">No navegador, abra o [arquivo de defini??es de tipo para o Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="2c6f4-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="2c6f4-129">Copie o conte?do do arquivo para a ?rea de transfer?ncia.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="2c6f4-130">No Visual Studio, abra o arquivo **Office.d.ts**, cole o conte?do de sua ?rea de transfer?ncia de arquivo e salve-o.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="2c6f4-131">Crie um novo arquivo chamado **jQuery.d.ts** na raiz do projeto de aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="2c6f4-132">No navegador, abra o [arquivo de defini??es de tipos para jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="2c6f4-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="2c6f4-133">Copie o conte?do do arquivo para a ?rea de transfer?ncia.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="2c6f4-134">No Visual Studio, abra o arquivo **jQuery.d.ts**, cole o conte?do de sua ?rea de transfer?ncia nesse arquivo e salve-o.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="2c6f4-135">No Visual Studio, crie um novo arquivo chamado **tsconfig.json** na raiz do projeto de aplicativo web.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="2c6f4-136">Abra o arquivo **tsconfig.json**, adicione o conte?do a seguir no arquivo e salve-o:</span><span class="sxs-lookup"><span data-stu-id="2c6f4-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="2c6f4-137">Abra o arquivo **Home.ts** e adicione a seguinte declara??o ? parte superior do arquivo:</span><span class="sxs-lookup"><span data-stu-id="2c6f4-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="2c6f4-138">No arquivo **Home.ts**, altere **'1.1'** para **1.1** (ou seja, remova as aspas) na seguinte linha e salve o arquivo:</span><span class="sxs-lookup"><span data-stu-id="2c6f4-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="2c6f4-139">Executar o projeto do suplemento convertido</span><span class="sxs-lookup"><span data-stu-id="2c6f4-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="2c6f4-p108">No Visual Studio, pressione F5 ou clique no bot?o **Iniciar** para iniciar o Excel com o bot?o do suplemento **Mostrar painel de tarefas** exibido na faixa de op??es. O suplemento ser? hospedado localmente no IIS.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="2c6f4-142">No Excel, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="2c6f4-143">Na planilha, selecione as nove c?lulas que cont?m n?meros.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="2c6f4-144">Pressione o bot?o **Real?ar** no painel de tarefas para real?ar a c?lula no intervalo selecionado com o maior valor.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="2c6f4-145">Arquivo de c?digo Home.ts</span><span class="sxs-lookup"><span data-stu-id="2c6f4-145">Home.ts code file</span></span>

<span data-ttu-id="2c6f4-146">Para sua refer?ncia o trecho de c?digo a seguir mostra o conte?do do arquivo **Home.ts** ap?s a aplica??o das altera??es descritas anteriormente.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="2c6f4-147">Esse c?digo cont?m o n?mero m?nimo de altera??es necess?rias para que seu suplemento seja executado.</span><span class="sxs-lookup"><span data-stu-id="2c6f4-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
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

## <a name="see-also"></a><span data-ttu-id="2c6f4-148">Confira tamb?m</span><span class="sxs-lookup"><span data-stu-id="2c6f4-148">See also</span></span>

* [<span data-ttu-id="2c6f4-149">Discuss?o de implementa??o do Promise no StackOverflow</span><span class="sxs-lookup"><span data-stu-id="2c6f4-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="2c6f4-150">Exemplos de Suplementos do Office no GitHub</span><span class="sxs-lookup"><span data-stu-id="2c6f4-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
