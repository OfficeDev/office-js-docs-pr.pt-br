# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="b80cc-101">Criar um suplemento do Excel usando o React</span><span class="sxs-lookup"><span data-stu-id="b80cc-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="b80cc-102">Neste artigo, voc? passar? pelo processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="b80cc-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="b80cc-103">Ambiente</span><span class="sxs-lookup"><span data-stu-id="b80cc-103">Environment</span></span>

- <span data-ttu-id="b80cc-104">**?rea de Trabalho do Office**: Verifique se voc? tem a ?ltima vers?o do Office instalada.</span><span class="sxs-lookup"><span data-stu-id="b80cc-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="b80cc-105">Comandos de suplemento precisam da compila??o 16.0.6769.0000 ou superior (**16.0.6868.0000** recomendada).</span><span class="sxs-lookup"><span data-stu-id="b80cc-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="b80cc-106">Saiba como [Instalar a ?ltima vers?o dos aplicativos do Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="b80cc-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="b80cc-107">**Office Online**: N?o h? configura??o adicional.</span><span class="sxs-lookup"><span data-stu-id="b80cc-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="b80cc-108">Observe que o suporte para comandos no Office Online para contas de trabalho/escola est? em vers?o pr?via.</span><span class="sxs-lookup"><span data-stu-id="b80cc-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b80cc-109">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="b80cc-109">Prerequisites</span></span>

- <span data-ttu-id="b80cc-110">Instale globalmente [Criar aplicativo do React](https://github.com/facebookincubator/create-react-app).</span><span class="sxs-lookup"><span data-stu-id="b80cc-110">Install [Create React App](https://github.com/facebookincubator/create-react-app) globally.</span></span>

    ```bash
    npm install -g create-react-app
    ```

- <span data-ttu-id="b80cc-111">Instale a ?ltima vers?o do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="b80cc-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a><span data-ttu-id="b80cc-112">Gerar um novo aplicativo do React</span><span class="sxs-lookup"><span data-stu-id="b80cc-112">Generate a new React app</span></span>

<span data-ttu-id="b80cc-113">Use Criar aplicativo do React para gerar seu aplicativo do React.</span><span class="sxs-lookup"><span data-stu-id="b80cc-113">Use Create React App to generate your React app.</span></span> <span data-ttu-id="b80cc-114">No terminal, execute o seguinte comando:</span><span class="sxs-lookup"><span data-stu-id="b80cc-114">From the terminal, run the following command:</span></span>

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a><span data-ttu-id="b80cc-115">Gerar o arquivo de manifesto e realizar sideload do suplemento</span><span class="sxs-lookup"><span data-stu-id="b80cc-115">Generate the manifest file and sideload the add-in</span></span>

<span data-ttu-id="b80cc-116">Cada suplemento requer um arquivo de manifesto para definir os recursos e configura??es.</span><span class="sxs-lookup"><span data-stu-id="b80cc-116">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="b80cc-117">Navegue at? a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="b80cc-117">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="b80cc-118">Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b80cc-118">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="b80cc-119">Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela:</span><span class="sxs-lookup"><span data-stu-id="b80cc-119">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="b80cc-120">**Gostaria de criar uma nova subpasta para o seu projeto?:** `No`</span><span class="sxs-lookup"><span data-stu-id="b80cc-120">**Would you like to create a new subfolder for your project?:** `No`</span></span>
    - <span data-ttu-id="b80cc-121">**Como deseja nomear seu suplemento?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="b80cc-121">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="b80cc-122">**Para qual aplicativo cliente do Office voc? deseja suporte?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="b80cc-122">**Which Office client application would you like to support?:** `Excel`</span></span>
    - <span data-ttu-id="b80cc-123">**Gostaria de criar um novo suplemento?:** `No`</span><span class="sxs-lookup"><span data-stu-id="b80cc-123">**Would you like to create a new add-in?:** `No`</span></span>

    <span data-ttu-id="b80cc-p105">O gerador perguntar? se voc? deseja abrir **resource.html**. N?o ? necess?rio abri-lo para este tutorial, mas fique ? vontade em fazer isso se tiver curiosidade. Escolha Sim ou N?o para concluir o assistente e deixar o gerador fazer seu trabalho.</span><span class="sxs-lookup"><span data-stu-id="b80cc-p105">The generator will then ask you if you want to open **resource.html**. It isn't necessary to open it for this tutorial, but feel free to open it if you're curious! Choose yes or no to complete the wizard and allow the generator to do its work.</span></span>

    ![Gerador do Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="b80cc-128">Se for solicitada a substitui??o de **package.json**, responda **N?o** (n?o substituir).</span><span class="sxs-lookup"><span data-stu-id="b80cc-128">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

3. <span data-ttu-id="b80cc-129">Siga as instru??es da plataforma que voc? usar? para executar o suplemento e realizar sideload do suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="b80cc-129">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="b80cc-130">Windows: [Realizar sideload de Suplementos do Office no Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="b80cc-130">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="b80cc-131">Excel Online: [Realizar sideload dos Suplementos do Office no Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="b80cc-131">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="b80cc-132">iPad e Mac: [Realizar sideload dos Suplementos do Office no iPad e Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="b80cc-132">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

## <a name="update-the-app"></a><span data-ttu-id="b80cc-133">Atualizar o aplicativo</span><span class="sxs-lookup"><span data-stu-id="b80cc-133">Update the app</span></span>

1. <span data-ttu-id="b80cc-134">Abra **public/index.html**, adicione a marca `<script>` a seguir imediatamente antes da marca `</head>` e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b80cc-134">Open **public/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. <span data-ttu-id="b80cc-135">Abra **src/index.js**, substitua `ReactDOM.render(<App />, document.getElementById('root'));` pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b80cc-135">Open **src/index.js**, replace `ReactDOM.render(<App />, document.getElementById('root'));` with the following code, and save the file.</span></span> 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. <span data-ttu-id="b80cc-136">Abra **src/App.js**, substitua o conte?do do arquivo pelo c?digo a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b80cc-136">Open **src/App.js**, replace file contents with the following code, and save the file.</span></span> 

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

4. <span data-ttu-id="b80cc-137">Abra **src/App.css**, substitua o conte?do do arquivo pelo c?digo de CSS a seguir e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b80cc-137">Open **src/App.css**, replace file contents with the following CSS code, and save the file.</span></span> 

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

## <a name="try-it-out"></a><span data-ttu-id="b80cc-138">Experimente</span><span class="sxs-lookup"><span data-stu-id="b80cc-138">Try it out</span></span>

1. <span data-ttu-id="b80cc-139">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="b80cc-139">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="b80cc-140">Windows:</span><span class="sxs-lookup"><span data-stu-id="b80cc-140">Windows:</span></span>
    ```bash
    set HTTPS=true&&npm start
    ```

    <span data-ttu-id="b80cc-141">macOS:</span><span class="sxs-lookup"><span data-stu-id="b80cc-141">macOS:</span></span>
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > <span data-ttu-id="b80cc-p106">Uma nova janela de navegador ser? aberta contendo o suplemento. Feche esta janela.</span><span class="sxs-lookup"><span data-stu-id="b80cc-p106">A browser window will open with the add-in in it. Close this window.</span></span>

2. <span data-ttu-id="b80cc-144">No Excel, escolha a guia **P?gina Inicial** e o bot?o **Mostrar Painel de Tarefas** na faixa de op??es para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b80cc-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Bot?o do Suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="b80cc-146">Selecione um intervalo de c?lulas na planilha.</span><span class="sxs-lookup"><span data-stu-id="b80cc-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="b80cc-147">No painel de tarefas, escolha o bot?o **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="b80cc-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="b80cc-149">Pr?ximas etapas</span><span class="sxs-lookup"><span data-stu-id="b80cc-149">Next steps</span></span>

<span data-ttu-id="b80cc-p107">Voc? criou com ?xito um suplemento do Excel usando o React, parab?ns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="b80cc-p107">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b80cc-152">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="b80cc-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="b80cc-153">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="b80cc-153">See also</span></span>

* [<span data-ttu-id="b80cc-154">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="b80cc-154">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="b80cc-155">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b80cc-155">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="b80cc-156">Exemplos de c?digo do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="b80cc-156">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="b80cc-157">Refer?ncia da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="b80cc-157">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
