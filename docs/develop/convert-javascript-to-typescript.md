---
title: Converter um projeto de Suplemento do Office no Visual Studio para TypeScript
description: Saiba como converter um projeto de suplemento do Office no Visual Studio para usar o TypeScript.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: daa81c3785484083aa49516b04491acad1404884
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889349"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Converter um projeto de Suplemento do Office no Visual Studio para TypeScript

Você pode usar o modelo de Suplemento do Office no Visual Studio para criar um suplemento que usa JavaScript e depois converter esse projeto de suplemento para o TypeScript. Este artigo descreve o processo de conversão de um suplemento do Excel. Você pode usar o mesmo processo para converter outros tipos de projetos de suplementos do Office de JavaScript para TypeScript no Visual Studio.

> [!IMPORTANT]
> Este artigo descreve as etapas  mínimas necessárias para garantir que, quando você pressionar F5, o código será transcompilado para JavaScript, que será então carregado automaticamente no Office. No entanto, o código não é muito "TypeScripty". Por exemplo, as variáveis são declaradas `var` com a palavra-chave `let` em vez de ou `const` não são declaradas com um tipo especificado. Para aproveitar ao máximo a forte digitação do TypeScript, considere fazer outras alterações no código.

> [!NOTE]
> Para criar um projeto de suplementos TypeScript do Office sem usar o Visual Studio, siga as instruções na seção "Gerador do Yeoman" de um [início rápido em 5 minutos](../index.yml) e escolha `TypeScript` quando for solicitado pelo [Gerador de suplementos do Office do Yeoman](yeoman-generator-overview.md).

## <a name="prerequisites"></a>Pré-requisitos

- [Visual Studio 2019 ou posterior com a](https://www.visualstudio.com/vs/) carga de trabalho de **desenvolvimento do Office/SharePoint** instalada

    > [!TIP]
    > Se você já instalou o Visual Studio, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho **de desenvolvimento do Office/SharePoint** esteja instalada. Se essa carga de trabalho ainda não estiver instalada, use o instalador do Visual Studio para [instalá-la](/visualstudio/install/modify-visual-studio#modify-workloads).

- SDK do TypeScript versão 2.3 ou posterior.

    > [!TIP]
    > No [Instalador do Visual Studio](/visualstudio/install/modify-visual-studio), selecione a guia **Componentes individuais** e role a tela para baixo até a seção **SDKs, bibliotecas e estruturas**. Nessa seção, verifique se pelo menos um dos componentes do **SDK do TypeScript** (versão 2.3 ou posterior) está selecionado. Se nenhum dos componentes do **SDK do TypeScript** estiver selecionado, selecione a versão mais recente disponível do SDK e escolha  Modificar para [instalar esse componente individual](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components).

- Excel 2016 ou posterior.

## <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. No Visual Studio, escolha **Criar um novo projeto**.

1. Usando a caixa de pesquisa, insira **suplemento**. Escolha **suplemento do Excel Web**, em seguida, selecione **Próximo**.

1. Nomeie seu projeto e selecione **Criar**.

1. Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel** e clique em **Concluir** para criar o projeto.

1. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.

## <a name="convert-the-add-in-project-to-typescript"></a>Converter o projeto do suplemento para TypeScript

1. Localize o arquivo **Home.js** e o renomeie para **Home.ts**.

1. Localize o arquivo **./Functions/FunctionFile.js** e renomeie-o para **FunctionFile.ts**.

1. Localize o arquivo **./Scripts/MessageBanner.js** e renomeie-o para **MessageBanner.ts**.

1. Na guia **Ferramentas**, escolha **Gerenciador de Pacotes NuGet** e, em seguida, selecione **Gerenciar Pacotes do NuGet para Solução...**.

1. Com a **guia** Procurar selecionada, insira **jquery. TypeScript.DefinitelyTyped**. Instale esse pacote ou atualize-o se ele já estiver instalado. Isso garantirá que as definições do TypeScript do jQuery sejam incluídas em seu projeto. Os pacotes para jQuery aparecem em um arquivo gerado pelo Visual Studio, chamado **packages.config**.

    > [!NOTE]
    > Em seu projeto em TypeScript, você pode ter uma combinação de arquivos TypeScript e JavaScript e seu projeto irá compilar. Isso ocorre porque o TypeScript é um superconjunto tipado do JavaScript que compila o JavaScript.

1. Em **Home.ts**, localize a linha `Office.initialize = function (reason) {` e adicione uma linha imediatamente depois dela para polifilar o global `window.Promise`, conforme mostrado aqui.

    ```TypeScript
    Office.initialize = function (reason) {
        // Add the following line.
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

1. Em **Home.ts**, localize a função `displaySelectedCells` , substitua toda a função pelo código a seguir e salve o arquivo.

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

1. Em **./Scripts/MessageBanner.ts**, localize a linha `_onResize(null);` e substitua-a pelo seguinte:

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a>Executar o projeto do suplemento convertido

1. No Visual Studio, pressione **F5** ou clique no botão **Iniciar** para iniciar o Excel com o botão do suplemento **Mostrar painel de tarefas** exibido na faixa de opções. O suplemento estará hospedado localmente no IIS.

1. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

1. Na planilha, selecione as nove células que contêm números.

1. Pressione o botão **Realçar** no painel de tarefas para realçar a célula no intervalo selecionado com o maior valor.

## <a name="homets-code-file"></a>Arquivo de código Home.ts

Para sua referência o trecho de código a seguir mostra o conteúdo do arquivo **Home.ts** após a aplicação das alterações descritas anteriormente. Esse código contém o número mínimo de alterações necessárias para que seu suplemento seja executado.

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

            // If you're using Excel 2013, use fallback logic.
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

## <a name="see-also"></a>Confira também

- [Discussão de implementação do Promise no StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [Exemplos de Suplementos do Office no GitHub](https://github.com/OfficeDev/Office-Add-in-samples)
