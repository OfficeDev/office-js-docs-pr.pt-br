# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="dd81f-101">Criar um suplemento do Excel usando o React</span><span class="sxs-lookup"><span data-stu-id="dd81f-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="dd81f-102">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="dd81f-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dd81f-103">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="dd81f-103">Prerequisites</span></span>

- [<span data-ttu-id="dd81f-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="dd81f-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="dd81f-105">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="dd81f-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="dd81f-106">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="dd81f-106">Create the web app</span></span>

1. <span data-ttu-id="dd81f-107">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="dd81f-107">Use the Yeoman generator to create an Outlook add-in project.</span></span> <span data-ttu-id="dd81f-108">Execute o comando a seguir e responda aos prompts da seguinte forma:</span><span class="sxs-lookup"><span data-stu-id="dd81f-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="dd81f-109">**Escolha o tipo de projeto:** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="dd81f-109">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="dd81f-110">**Qual será o nome do suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="dd81f-110">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="dd81f-111">**Você gostaria de proporcionar suporte para qual aplicativo cliente do Office?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="dd81f-111">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="dd81f-113">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="dd81f-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="dd81f-114">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="dd81f-114">Navigate to the root folder of the project in the Terminal app, and from Terminal run:</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="dd81f-115">Atualizar o código</span><span class="sxs-lookup"><span data-stu-id="dd81f-115">Update the code</span></span>

1. <span data-ttu-id="dd81f-116">Em seu editor de código, abra o arquivo **src/styles.less**, inclua os seguintes estilos no final do arquivo e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dd81f-116">In your code editor, open the file **src/styles.less**, add the following styles to the end of the file, and save the file.</span></span>

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

2. <span data-ttu-id="dd81f-117">O modelo de projeto criado pelo gerador Yeoman de Suplementos do Office inclui um componente que não é necessário para este início rápido.</span><span class="sxs-lookup"><span data-stu-id="dd81f-117">The project template that the Office Add-ins Yeoman generator created includes a React component that is not needed for this quick start.</span></span> <span data-ttu-id="dd81f-118">Exclua o arquivo **src/components/HeroList.tsx**.</span><span class="sxs-lookup"><span data-stu-id="dd81f-118">Delete the file **src/components/HeroList.tsx**.</span></span>

3. <span data-ttu-id="dd81f-119">Abra o arquivo **src/components/Header.tsx**, substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dd81f-119">Open the file **src/components/Header.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. <span data-ttu-id="dd81f-120">Crie um novo componente React chamado **Content.tsx** na pasta **src/components**, adicione o seguinte código e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dd81f-120">Create a new React component named **Content.tsx** in the **src/components** folder, add the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. <span data-ttu-id="dd81f-121">Abra o arquivo **src/components/App.tsx**, substitua todo o conteúdo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dd81f-121">Open the file **src/components/App.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
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

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="dd81f-122">Atualizar o manifesto</span><span class="sxs-lookup"><span data-stu-id="dd81f-122">Update the manifest</span></span>

1. <span data-ttu-id="dd81f-123">Abra o arquivo **manifest.xml** para definir as configurações e os recursos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="dd81f-123">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="dd81f-124">O elemento `ProviderName` tem um valor de espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="dd81f-124">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="dd81f-125">Substitua-o com seu nome.</span><span class="sxs-lookup"><span data-stu-id="dd81f-125">Replace it with your name.</span></span>

3. <span data-ttu-id="dd81f-126">O atributo `DefaultValue` do elemento `Description` tem um espaço reservado.</span><span class="sxs-lookup"><span data-stu-id="dd81f-126">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="dd81f-127">Substitua-o com **um suplemento do painel de tarefas do Excel**.</span><span class="sxs-lookup"><span data-stu-id="dd81f-127">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="dd81f-128">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="dd81f-128">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="dd81f-129">Iniciar o servidor de desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="dd81f-129">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="dd81f-130">Experimente</span><span class="sxs-lookup"><span data-stu-id="dd81f-130">Try it out</span></span>

1. <span data-ttu-id="dd81f-131">Siga as instruções para a plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="dd81f-131">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="dd81f-132">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="dd81f-132">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="dd81f-133">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="dd81f-133">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="dd81f-134">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="dd81f-134">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="dd81f-135">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="dd81f-135">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="dd81f-137">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="dd81f-137">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="dd81f-138">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="dd81f-138">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="dd81f-140">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="dd81f-140">Next steps</span></span>

<span data-ttu-id="dd81f-p105">Você criou com êxito um suplemento do Excel usando o React, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="dd81f-p105">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="dd81f-143">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dd81f-143">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="dd81f-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="dd81f-144">See also</span></span>

* [<span data-ttu-id="dd81f-145">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dd81f-145">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="dd81f-146">Conceitos fundamentais de programação com a API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dd81f-146">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="dd81f-147">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="dd81f-147">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="dd81f-148">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="dd81f-148">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
