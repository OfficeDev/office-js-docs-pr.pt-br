---
title: Criar um suplemento do Excel usando o Angular
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a7600b5e53fa9c7361f2dd0c117d97937ec954c7
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742370"
---
# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="56552-102">Criar um suplemento do Excel usando o Angular</span><span class="sxs-lookup"><span data-stu-id="56552-102">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="56552-103">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o Angular e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="56552-103">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="56552-104">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="56552-104">Prerequisites</span></span>

- [<span data-ttu-id="56552-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="56552-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="56552-106">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="56552-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="56552-107">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="56552-107">Create the web app</span></span>

1. <span data-ttu-id="56552-108">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="56552-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="56552-109">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="56552-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="56552-110">**Escolha o tipo de projeto:** `Office Add-in project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="56552-110">**Choose a project type:** `Office Add-in project using Angular framework`</span></span>
    - <span data-ttu-id="56552-111">**Escolha o tipo de script:** `Typescript`</span><span class="sxs-lookup"><span data-stu-id="56552-111">**Choose a script type:** `Typescript`</span></span>
    - <span data-ttu-id="56552-112">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="56552-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="56552-113">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="56552-113">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-excel-angular.png)
    
    <span data-ttu-id="56552-115">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="56552-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="56552-116">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="56552-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="56552-117">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="56552-117">Update the code</span></span>

1. <span data-ttu-id="56552-118">Em seu editor de código, abra o arquivo **app.css**, inclua os seguintes estilos no final do arquivo e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="56552-118">In your code editor, open the file **app.css**, add the following styles to the end of the file, and save the file.</span></span>

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
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="56552-119">Abra o arquivo **src/app/app.component.html**, substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="56552-119">Open the file **src/app/app.component.html**, replace the entire contents with the following code, and save the file.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>{{welcomeMessage}}</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <br />
            <div role="button" class="ms-Button" (click)="setColor()">
                <span class="ms-Button-label">Set color</span>
                <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
            </div>
        </div>
    </div>
    ```

3. <span data-ttu-id="56552-120">Abra o arquivo **src/app/app.component.ts**, substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="56552-120">Open the file **src/app/app.component.ts**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import { Component } from '@angular/core';
    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    const template = require('./app.component.html');

    @Component({
        selector: 'app-home',
        template
    })
    export default class AppComponent {
        welcomeMessage = 'Welcome';

        async setColor() {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="56552-121">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="56552-121">Update the manifest</span></span>

1. <span data-ttu-id="56552-122">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="56552-122">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="56552-123">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="56552-123">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="56552-124">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="56552-124">Replace it with your name.</span></span>

3. <span data-ttu-id="56552-125">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="56552-125">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="56552-126">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="56552-126">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="56552-127">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="56552-127">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="56552-128">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="56552-128">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="56552-129">Experimente</span><span class="sxs-lookup"><span data-stu-id="56552-129">Try it out</span></span>

1. <span data-ttu-id="56552-130">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="56552-130">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="56552-131">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="56552-131">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="56552-132">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="56552-132">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="56552-133">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="56552-133">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="56552-134">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="56552-134">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="56552-136">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="56552-136">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="56552-137">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="56552-137">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="56552-139">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="56552-139">Next steps</span></span>

<span data-ttu-id="56552-p104">Você criou com êxito um suplemento do Excel usando o Angular!, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="56552-p104">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="56552-142">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="56552-142">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="56552-143">Confira também</span><span class="sxs-lookup"><span data-stu-id="56552-143">See also</span></span>

* [<span data-ttu-id="56552-144">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="56552-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="56552-145">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="56552-145">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="56552-146">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="56552-146">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="56552-147">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="56552-147">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

