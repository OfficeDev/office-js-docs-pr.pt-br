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
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Converter um projeto de Suplemento do Office no Visual Studio para TypeScript

Voc? pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript. Use o Visual Studio para criar projeto suplemento, evite ter que criar desde o in?cio o projeto de suplemento do Office no TypeScript. 

Este artigo mostra como criar um suplemento do Excel usando o Visual Studio e depois converter o projeto do suplemento do JavaScript para o TypeScript. Voc? pode usar o mesmo processo para converter outros tipos de projetos de JavaScript para Suplementos do Office para o TypeScript no Visual Studio.

> [!NOTE]
> Para criar um projeto TypeScript do Suplemento do Office sem usar o Visual Studio, siga as instru??es na se??o "Qualquer editor" de qualquer [In?cio r?pido de 5 minutos](../index.yml) e escolha `TypeScript` quando solicitado pelo [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office).

## <a name="prerequisites"></a>Pr?-requisitos

- [Visual Studio 2017](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada

    > [!NOTE]
    > Se voc? j? instalou o Visual Studio 2017, [use o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada. 

- TypeScript 2.3 para Visual Studio 2017

    > [!NOTE]
    > O TypeScript deve ser instalado por padr?o com o Visual Studio 2017, mas voc? pode [usar o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para confirmar se ele foi instalado. No Instalador do Visual Studio, selecione a guia **Componentes individuais** e verifique se a op??o**TypeScript 2.3 SDK** est? selecionada em **SDKs, bibliotecas e estruturas**.

- Excel 2016

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo**  >  **Projeto**.

2. Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a op??o **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto. 

3. D? um nome ao projeto e escolha **OK**.

4. Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.

5. O Visual Studio cria uma solu??o, e os dois projetos dele s?o exibidos no **Gerenciador de Solu??es**. O arquivo **Home.html** ? aberto no Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Converter o projeto do suplemento para TypeScript

1. No **Gerenciador de Solu??es**, renomeie o arquivo **Home.js** como **Home.ts**.

    > [!NOTE]
    > Em seu projeto em TypeScript, voc? pode ter uma combina??o de arquivos TypeScript e JavaScript e seu projeto ir? compilar. Isso ocorre porque o TypeScript ? um superconjunto tipado do JavaScript que compila o JavaScript. 

2. Selecione **Sim** para confirmar que voc? deseja alterar a extens?o do nome de arquivo.

3. Crie um novo arquivo chamado **Office.d.ts** na raiz do projeto de aplicativo Web.

4. No navegador, abra o [arquivo de defini??es de tipo para o Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copie o conte?do do arquivo para a ?rea de transfer?ncia.

5. No Visual Studio, abra o arquivo **Office.d.ts**, cole o conte?do de sua ?rea de transfer?ncia de arquivo e salve-o.

6. Crie um novo arquivo chamado **jQuery.d.ts** na raiz do projeto de aplicativo Web.

7. No navegador, abra o [arquivo de defini??es de tipos para jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts). Copie o conte?do do arquivo para a ?rea de transfer?ncia.

8. No Visual Studio, abra o arquivo **jQuery.d.ts**, cole o conte?do de sua ?rea de transfer?ncia nesse arquivo e salve-o.

9. No Visual Studio, crie um novo arquivo chamado **tsconfig.json** na raiz do projeto de aplicativo web.

10. Abra o arquivo **tsconfig.json**, adicione o conte?do a seguir no arquivo e salve-o:

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. Abra o arquivo **Home.ts** e adicione a seguinte declara??o ? parte superior do arquivo:

    ```javascript
    declare var fabric: any;
    ```

12. No arquivo **Home.ts**, altere **'1.1'** para **1.1** (ou seja, remova as aspas) na seguinte linha e salve o arquivo:

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a>Executar o projeto do suplemento convertido

1. No Visual Studio, pressione F5 ou clique no bot?o **Iniciar** para iniciar o Excel com o bot?o do suplemento **Mostrar painel de tarefas** exibido na faixa de op??es. O suplemento ser? hospedado localmente no IIS.

2. No Excel, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.

3. Na planilha, selecione as nove c?lulas que cont?m n?meros.

4. Pressione o bot?o **Real?ar** no painel de tarefas para real?ar a c?lula no intervalo selecionado com o maior valor.

## <a name="homets-code-file"></a>Arquivo de c?digo Home.ts

Para sua refer?ncia o trecho de c?digo a seguir mostra o conte?do do arquivo **Home.ts** ap?s a aplica??o das altera??es descritas anteriormente. Esse c?digo cont?m o n?mero m?nimo de altera??es necess?rias para que seu suplemento seja executado.

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

## <a name="see-also"></a>Confira tamb?m

* [Discuss?o de implementa??o do Promise no StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Exemplos de Suplementos do Office no GitHub](https://github.com/officedev)
