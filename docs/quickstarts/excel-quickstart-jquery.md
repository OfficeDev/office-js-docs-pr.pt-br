---
title: Crie seu primeiro suplemento do painel de tarefas do Excel
description: Saiba como criar um Suplemento do Excel simples usando a API JS do Office.
ms.date: 07/12/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: f4bd32e9e38e72b8cfb42db9476f6c180b151909
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797516"
---
# <a name="build-an-excel-task-pane-add-in"></a>Criar um suplemento do painel de tarefas do Excel

Neste artigo, você passará pelo processo de criação de um suplemento do painel de tarefas do Excel.

## <a name="create-the-add-in"></a>Criar o suplemento

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generator"></a>[Gerador do Yeoman](#tab/yeomangenerator)

### <a name="prerequisites"></a>Pré-requisitos

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Escolha o tipo de projeto:** `Office Add-in Task Pane project`
- **Escolha o tipo de script:** `Javascript`
- **Qual será o nome do suplemento?** `My Office Add-in`
- **Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?** `Excel`

![O Gerador Yeoman para a interface de linha de comando de Suplementos do Office.](../images/yo-office-excel.png)

Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a>Explore o projeto

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a>Experimente

1. Navegue até a pasta raiz do projeto.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![O menu da Página Inicial do Excel, com o botão Mostrar Painel de Tarefas destacado.](../images/excel-quickstart-addin-3b.png)

1. Selecione um intervalo de células na planilha.

1. Na parte inferior do painel de tarefas, escolha o link **Executar** para definir a cor do intervalo selecionado como amarelo.

    ![O painel de tarefas do suplemento aberto no Excel, com o botão Executar destacado no painel de tarefas do suplemento.](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel! Em seguida, saiba mais sobre os recursos de um suplemento do Excel e crie um suplemento mais complexo seguindo o [tutorial do suplemento Excel](../tutorials/excel-tutorial.md).

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>Pré-requisitos

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Criar o projeto do suplemento

1. No Visual Studio, escolha **Criar um novo projeto**.

1. Usando a caixa de pesquisa, insira **suplemento**. Escolha **suplemento do Excel Web**, em seguida, selecione **Próximo**.

1. Nomeie seu projeto **ExcelWebAddIn1** e selecione **Criar**.

1. Na janela **Criar Suplemento do Office**, escolha **Adicionar novas funcionalidades para o Excel** e clique em **Concluir** para criar o projeto.

1. O Visual Studio cria uma solução, e os dois projetos dele são exibidos no **Gerenciador de Soluções**. O arquivo **Home.html** é aberto no Visual Studio.

### <a name="explore-the-visual-studio-solution"></a>Explorar a solução do Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>Atualizar o código

1. **Home.html** especifica o HTML que será renderizado no painel de tarefas do suplemento. Em **Home.html**, substitua o elemento `<body>` pela marcação a seguir e salve o arquivo.

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

1. Abra o arquivo **Home.js** na raiz do projeto do aplicativo Web. Este arquivo especifica o script do suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

    [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        async function setColor() {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                await context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

1. Abra o arquivo **Home.css** na raiz do projeto do aplicativo Web. Este arquivo especifica os estilos personalizados para o suplemento. Substitua todo o conteúdo pelo código a seguir e salve o arquivo.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px;
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto;
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Atualizar o manifesto

1. No **Gerenciador de Soluções**, vá para o projeto de suplemento **ExcelWebAddIn1** e abra o diretório **ExcelWebAddIn1Manifest**. Este diretório contém seu arquivo de manifesto, **ExcelWebAddIn1.xml**. O arquivo de manifesto XML define as configurações e recursos do suplemento. Consulte a seção anterior [Explorar a solução Visual Studio](#explore-the-visual-studio-solution) para obter mais informações sobre os dois projetos criados por sua solução Visual Studio.

1. O elemento `ProviderName` tem um valor de espaço reservado. Substitua-o com seu nome.

1. O atributo `DefaultValue` do elemento `DisplayName` tem um espaço reservado. Substitua-o pelo **suplementos do My Office**.

1. O atributo `DefaultValue` do elemento `Description` tem um espaço reservado. Substitua-o com **um suplemento do painel de tarefas do Excel**.

1. Salve o arquivo.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a>Experimente

1. Use o Visual Studio, teste o suplemento recém-criado do Excel pressionando **F5** ou escolha o botão **Iniciar** para iniciar o Excel com o botão do suplemento **Exibir painel de tarefas** exibido na faixa de opções. O suplemento será hospedado localmente no IIS. Se for solicitado que você confie em um certificado, faça-o para permitir que o suplemento se conecte ao seu aplicativo do Office.

1. No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.

    ![O menu da Página Inicial do Excel, com o botão Mostrar Painel de Tarefas destacado.](../images/excel-quickstart-addin-2a.png)

1. Selecione um intervalo de células na planilha.

1. No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.

    ![O painel de tarefas do suplemento é aberto no Excel.](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a>Próximas etapas

Parabéns, você criou com êxito um suplemento do painel de tarefas do Excel! Em seguida, saiba mais sobre como [desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).

---

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Modelo de objeto JavaScript do Excel em Suplementos do Office](../excel/excel-add-ins-core-concepts.md)
- [Exemplos de código do suplemento do Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Referência da API JavaScript do Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Usando o Visual Studio Code para publicar](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
