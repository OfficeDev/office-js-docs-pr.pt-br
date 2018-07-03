# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="f65b9-101">Criar um suplemento do Excel usando o React</span><span class="sxs-lookup"><span data-stu-id="f65b9-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="f65b9-102">Neste artigo, você passará pelo processo de criação de um suplemento do Excel usando o React e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="f65b9-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="f65b9-103">Ambiente</span><span class="sxs-lookup"><span data-stu-id="f65b9-103">Environment</span></span>

- <span data-ttu-id="f65b9-104">**Área de Trabalho do Office**: Verifique se você tem a última versão do Office instalada.</span><span class="sxs-lookup"><span data-stu-id="f65b9-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="f65b9-105">Comandos de suplemento precisam da compilação 16.0.6769.0000 ou superior (**16.0.6868.0000** recomendada).</span><span class="sxs-lookup"><span data-stu-id="f65b9-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="f65b9-106">Saiba como [Instalar a última versão dos aplicativos do Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="f65b9-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="f65b9-107">**Office Online**: Não há configuração adicional.</span><span class="sxs-lookup"><span data-stu-id="f65b9-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="f65b9-108">Observe que o suporte para comandos no Office Online para contas de trabalho/escola está em versão prévia.</span><span class="sxs-lookup"><span data-stu-id="f65b9-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f65b9-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="f65b9-109">Prerequisites</span></span>

- <span data-ttu-id="f65b9-110">Instale globalmente [Criar aplicativo do React](https://github.com/facebookincubator/create-react-app).</span><span class="sxs-lookup"><span data-stu-id="f65b9-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="f65b9-111">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="f65b9-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="f65b9-112">Gerar um novo aplicativo do React</span><span class="sxs-lookup"><span data-stu-id="f65b9-112">Generate a new React app</span></span>

<span data-ttu-id="f65b9-113">Use Criar aplicativo do React para gerar seu aplicativo do React.</span><span class="sxs-lookup"><span data-stu-id="f65b9-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="f65b9-114">No terminal, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="f65b9-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="f65b9-115">Gerar o arquivo de manifesto e realizar sideload do suplemento</span><span class="sxs-lookup"><span data-stu-id="f65b9-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="f65b9-116">Cada suplemento requer um arquivo de manifesto para definir os recursos e configurações.</span><span class="sxs-lookup"><span data-stu-id="f65b9-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="f65b9-117">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="f65b9-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="f65b9-118">Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f65b9-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="f65b9-119">Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:</span><span class="sxs-lookup"><span data-stu-id="f65b9-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="f65b9-120">**Escolha um tipo de projeto:** `Manifest`</span><span class="sxs-lookup"><span data-stu-id="f65b9-120">**Choose a project type:** `Manifest`</span></span>
    - <span data-ttu-id="f65b9-121">**Como deseja nomear seu suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="f65b9-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="f65b9-122">**Para qual aplicativo cliente do Office você deseja suporte?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="f65b9-122">**Which Office client application would you like to support?:** `Excel`</span></span>


    <span data-ttu-id="f65b9-123">Depois de concluir o assistente, um arquivo de manifesto e um arquivo de recurso estarão disponíveis para você criar seu projeto.</span><span class="sxs-lookup"><span data-stu-id="f65b9-123">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>
    
    ![Gerador do Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="f65b9-125">Se for solicitada a substituição de **package.json**, responda **Não** (não substituir).</span><span class="sxs-lookup"><span data-stu-id="f65b9-125">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="f65b9-126">Siga as instruções da plataforma que você usará para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="f65b9-126">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="f65b9-127">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="f65b9-127">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="f65b9-128">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="f65b9-128">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="f65b9-129">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="f65b9-129">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="f65b9-130">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="f65b9-130">Update the app</span></span>

1. <span data-ttu-id="f65b9-131">Abra **public/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f65b9-131">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="f65b9-132">Abra **src/index.js**, substitua `ReactDOM.render(<App />, document.getElementById('root'));` pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f65b9-132">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="f65b9-133">Abra **src/App.js**, substitua o conteúdo do arquivo pelo código a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f65b9-133">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. <span data-ttu-id="f65b9-134">Abra **src/App.css**, substitua o conteúdo do arquivo pelo código de CSS a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="f65b9-134">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

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

## <a name="try-it-out"></a><span data-ttu-id="f65b9-135">Experimente</span><span class="sxs-lookup"><span data-stu-id="f65b9-135">Try it out</span></span>

1. <span data-ttu-id="f65b9-136">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f65b9-136">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="f65b9-137">Windows:</span><span class="sxs-lookup"><span data-stu-id="f65b9-137">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="f65b9-138">macOS:</span><span class="sxs-lookup"><span data-stu-id="f65b9-138">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="f65b9-p105">Uma nova janela de navegador será aberta contendo o suplemento. Feche esta janela.</span><span class="sxs-lookup"><span data-stu-id="f65b9-p105">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="f65b9-141">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f65b9-141">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="f65b9-143">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="f65b9-143">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="f65b9-144">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="f65b9-144">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="f65b9-146">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f65b9-146">Next steps</span></span>

<span data-ttu-id="f65b9-p106">Você criou com êxito um suplemento do Excel usando o React, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="f65b9-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f65b9-149">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="f65b9-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="f65b9-150">Veja também</span><span class="sxs-lookup"><span data-stu-id="f65b9-150">See also</span></span>

* [<span data-ttu-id="f65b9-151">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="f65b9-151">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="f65b9-152">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f65b9-152">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="f65b9-153">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="f65b9-153">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="f65b9-154">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="f65b9-154">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
