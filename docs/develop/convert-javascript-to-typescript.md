---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 9b3916dc61fadb3b6d9bf61e43cb22bdc7ff68c8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448767"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Converter um projeto de Suplemento do Office no Visual Studio para TypeScript

Você pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript. Este artigo descreve o processo de conversão de um suplemento do Excel. Você pode usar o mesmo processo para converter outros tipos de projetos de suplementos do Office de JavaScript para TypeScript no Visual Studio.

> [!NOTE]
> Para criar um projeto de suplementos TypeScript do Office sem usar o Visual Studio, siga as instruções na seção "Qualquer editor" de um [início rápido em 5 minutos](../index.yml) e escolha `TypeScript` quando for solicitado pelo [Gerador de suplementos do Office do Yeoman](https://github.com/officedev/generator-office).

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio 2017](https://www.visualstudio.com/vs/) com a carga de trabalho de **desenvolvimento do Office/SharePoint** instalada

    > [!TIP]
    > Se você já instalou o Visual Studio 2017, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Office/SharePoint** seja instalada. Se essa carga de trabalho ainda não estiver instalada, use o instalador do Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-workloads).

- SDK do TypeScript versão 2.3 ou posterior (para Visual Studio 2017)

    > [!TIP]
    > No [Instalador do Visual Studio](/visualstudio/install/modify-visual-studio), selecione a guia **Componentes individuais** e role a tela para baixo até a seção **SDKs, bibliotecas e estruturas**. Nessa seção, verifique se pelo menos um dos componentes do **SDK do TypeScript** (versão 2.3 ou posterior) está selecionado. Se nenhum dos componentes do **SDK do TypeScript** estiver selecionado, selecione a versão mais recente do SDK disponível e, em seguida, escolha o botão **Modificar** para [instalar esse componente individual](/visualstudio/install/modify-visual-studio?view=vs-2017#modify-individual-components). 

- Excel 2016 ou posterior

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. Na barra de menus do Visual Studio, selecione **Arquivo** > **Novo**  >  **Projeto**.

2. Na lista de tipos de projeto em **Visual C#** ou em **Visual Basic**, expanda a opção **Office/SharePoint**, escolha **Suplementos** e depois **Suplemento da Web do Outlook** como o tipo de projeto. 

3. Dê um nome ao projeto e escolha **OK**.

4. Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel**e clique em **Concluir** para criar o projeto.

5. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Converter o projeto do suplemento para TypeScript

1. No **Gerenciador de Soluções**, renomeie o arquivo **Home.js** como **Home.ts**.

    > [!NOTE]
    > Em seu projeto em TypeScript, você pode ter uma combinação de arquivos TypeScript e JavaScript e seu projeto irá compilar. Isso ocorre porque o TypeScript é um superconjunto tipado do JavaScript que compila o JavaScript. 

2. Selecione **Sim** para confirmar que você deseja alterar a extensão do nome de arquivo.

3. Crie um novo arquivo chamado **Office.d.ts** na raiz do projeto de aplicativo Web.

4. No navegador, abra o [arquivo de definições de tipo para o Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Copie o conteúdo do arquivo para a área de transferência.

5. No Visual Studio, abra o arquivo **Office.d.ts**, cole o conteúdo de sua área de transferência de arquivo e salve-o.

6. Crie um novo arquivo chamado **jQuery.d.ts** na raiz do projeto de aplicativo Web.

7. No navegador, abra o [arquivo de definições de tipos para jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/misc.d.ts). Copie o conteúdo do arquivo para a área de transferência.

8. No Visual Studio, abra o arquivo **jQuery.d.ts**, cole o conteúdo de sua área de transferência nesse arquivo e salve-o.

9. No Visual Studio, crie um novo arquivo chamado **tsconfig.json** na raiz do projeto de aplicativo web.

10. Abra o arquivo **tsconfig.json**, adicione o conteúdo a seguir no arquivo e salve-o:

    ```json
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. Abra o arquivo **Home.ts** e adicione a seguinte declaração à parte superior do arquivo:

    ```typescript
    declare var fabric: any;
    ```

12. No arquivo **Home.ts**, altere **'1.1'** para **1.1** (ou seja, remova as aspas) na seguinte linha:

    ```typescript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

13. No arquivo **Home.ts**, localize a função `displaySelectedCells`, substitua toda a função pelo código a seguir e salve o arquivo:

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

## <a name="run-the-converted-add-in-project"></a>Executar o projeto do suplemento convertido

1. No Visual Studio, pressione **F5** ou clique no botão **Iniciar** para iniciar o Excel com o botão do suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento estará hospedado localmente no IIS.

2. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

3. Na planilha, selecione as nove células que contêm números.

4. Pressione o botão **Realçar** no painel de tarefas para realçar a célula no intervalo selecionado com o maior valor.

## <a name="homets-code-file"></a>Arquivo de código Home.ts

Para sua referência o trecho de código a seguir mostra o conteúdo do arquivo **Home.ts** após a aplicação das alterações descritas anteriormente. Esse código contém o número mínimo de alterações necessárias para que seu suplemento seja executado.

```typescript
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

## <a name="see-also"></a>Confira também

* [Discussão de implementação do Promise no StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Exemplos de Suplementos do Office no GitHub](https://github.com/officedev)
