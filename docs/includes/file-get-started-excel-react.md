# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="166a0-101">Criar um suplemento do Excel usando o React</span><span class="sxs-lookup"><span data-stu-id="166a0-101">Build an Excel add-in using React</span></span>

<span data-ttu-id="166a0-102">Neste artigo, você passará pelo processo de criar um suplemento do Excel usando o React e a API JavaScript do Excel.</span><span class="sxs-lookup"><span data-stu-id="166a0-102">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="environment"></a><span data-ttu-id="166a0-103">Ambiente</span><span class="sxs-lookup"><span data-stu-id="166a0-103">Environment</span></span>

- <span data-ttu-id="166a0-104">**Área de Trabalho do Office**: Verifique se você tem a última versão do Office instalada.</span><span class="sxs-lookup"><span data-stu-id="166a0-104">**Office Desktop**: Ensure that you have the latest version of Office installed.</span></span> <span data-ttu-id="166a0-105">Comandos de suplemento precisam da compilação 16.0.6769.0000 ou superior (**16.0.6868.0000** recomendada).</span><span class="sxs-lookup"><span data-stu-id="166a0-105">Add-in commands require build 16.0.6769.0000 or higher (**16.0.6868.0000** recommended).</span></span> <span data-ttu-id="166a0-106">Saiba como [Instalar a última versão dos aplicativos do Office](http://aka.ms/latestoffice).</span><span class="sxs-lookup"><span data-stu-id="166a0-106">Learn how to [Install the latest version of Office applications](http://aka.ms/latestoffice).</span></span> 
 
- <span data-ttu-id="166a0-107">**Office Online**: Não há configuração adicional.</span><span class="sxs-lookup"><span data-stu-id="166a0-107">**Office Online**: There is no additional setup.</span></span> <span data-ttu-id="166a0-108">Observe que o suporte para comandos no Office Online para contas de trabalho/escola está em versão prévia.</span><span class="sxs-lookup"><span data-stu-id="166a0-108">Please note that support for commands in Office Online for work/school accounts is in preview.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="166a0-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="166a0-109">Prerequisites</span></span>

- [<span data-ttu-id="166a0-110">Node.js</span><span class="sxs-lookup"><span data-stu-id="166a0-110">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="166a0-111">Instale a última versão do [Yeoman](https://github.com/yeoman/yo) e o [gerador do Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office) globalmente.</span><span class="sxs-lookup"><span data-stu-id="166a0-111">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="166a0-112">Criar o aplicativo Web</span><span class="sxs-lookup"><span data-stu-id="166a0-112">Create the web app</span></span>

1. <span data-ttu-id="166a0-113">Crie uma pasta na sua unidade local e nomeie-a como **my-addin**.</span><span class="sxs-lookup"><span data-stu-id="166a0-113">Create a folder on your local drive and name it **my-addin**.</span></span> <span data-ttu-id="166a0-114">Esse é o local em que você criará os arquivos para seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="166a0-114">This is where you'll create the files for your app.</span></span>

2. <span data-ttu-id="166a0-115">Navegue até a pasta do seu aplicativo.</span><span class="sxs-lookup"><span data-stu-id="166a0-115">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

3. <span data-ttu-id="166a0-116">Use o gerador do Yeoman para gerar o arquivo de manifesto para o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="166a0-116">Use the Yeoman generator to generate the manifest file for your add-in.</span></span> <span data-ttu-id="166a0-117">Execute o comando a seguir e responda aos prompts, conforme mostrado na seguinte captura de tela.</span><span class="sxs-lookup"><span data-stu-id="166a0-117">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="166a0-118">**Escolha um tipo de projeto:** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="166a0-118">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="166a0-119">**Como deseja nomear seu suplemento?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="166a0-119">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="166a0-120">**Para qual aplicativo cliente do Office você deseja suporte?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="166a0-120">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Gerador do Yeoman](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="166a0-122">Depois de concluir o assistente, o gerador criará o projeto e instalará os componentes do nó de suporte.</span><span class="sxs-lookup"><span data-stu-id="166a0-122">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

4.  <span data-ttu-id="166a0-123">Abra **src/components/App.tsx**, procure o comentário "Atualizar a cor de preenchimento", altere a cor de preenchimento de "amarelo" para "azul" e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="166a0-123">Open **src/components/App.tsx**, search for the comment "Update the fill color," then change the fill color from 'yellow' to 'blue', and save the file.</span></span> 

    ```js
    range.format.fill.color = 'blue'

    ```

5. <span data-ttu-id="166a0-124">No bloco `return` da função `render` em **src/components/App.tsx**, atualize `<Herolist>` para o código abaixo e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="166a0-124">In the `return` block of the `render` function within **src/components/App.tsx**, update the `<Herolist>` to the code below, and save the file.</span></span> 

    ```js
      <HeroList message='Discover what My Office Add-in can do for you today!' items={this.state.listItems}>
        <p className='ms-font-l'>Choose the button below to set the color of the selected range to blue. <b>Set color</b>.</p>
        <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
    </HeroList>
    ```

6. <span data-ttu-id="166a0-125">Execute as etapas em [Adicionar certificados autoassinados como certificado raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) para confiar no certificado do sistema operacional do seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="166a0-125">Carry out the steps in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

7. <span data-ttu-id="166a0-126">Carregue seu suplemento para que ele apareça no Excel.</span><span class="sxs-lookup"><span data-stu-id="166a0-126">Sideload your add-in so it will appear in Excel.</span></span> <span data-ttu-id="166a0-127">No terminal, execute o comando a seguir:</span><span class="sxs-lookup"><span data-stu-id="166a0-127">In the terminal run the following command:</span></span> 
    
    ```bash
    npm run sideload
    ```

## <a name="try-it-out"></a><span data-ttu-id="166a0-128">Experimente</span><span class="sxs-lookup"><span data-stu-id="166a0-128">Try it out</span></span>

1. <span data-ttu-id="166a0-129">No terminal, execute o comando a seguir para iniciar o servidor de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="166a0-129">From the terminal, run the following command to start the dev server.</span></span>

    <span data-ttu-id="166a0-130">Windows:</span><span class="sxs-lookup"><span data-stu-id="166a0-130">Windows:</span></span>
    ```bash
    npm start
    ```

2. <span data-ttu-id="166a0-131">No Excel, escolha a guia **Página Inicial** e o botão **Mostrar Painel de Tarefas** na faixa de opções para abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="166a0-131">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Botão do Suplemento do Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="166a0-133">Selecione um intervalo de células na planilha.</span><span class="sxs-lookup"><span data-stu-id="166a0-133">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="166a0-134">No painel de tarefas, escolha o botão **Definir cor** para definir a cor do intervalo selecionado como verde.</span><span class="sxs-lookup"><span data-stu-id="166a0-134">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Suplemento do Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="166a0-136">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="166a0-136">Next steps</span></span>

<span data-ttu-id="166a0-p106">Você criou com êxito um suplemento do Excel usando o React, parabéns! Agora, saiba mais sobre os recursos dos suplementos do Excel e crie um mais complexo, acompanhando o tutorial de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="166a0-p106">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="166a0-139">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="166a0-139">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="166a0-140">Veja também</span><span class="sxs-lookup"><span data-stu-id="166a0-140">See also</span></span>

* [<span data-ttu-id="166a0-141">Tutorial de suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="166a0-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="166a0-142">Principais conceitos da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="166a0-142">Excel JavaScript API core concepts</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="166a0-143">Exemplos de código do suplemento do Excel</span><span class="sxs-lookup"><span data-stu-id="166a0-143">Excel add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="166a0-144">Referência da API JavaScript do Excel</span><span class="sxs-lookup"><span data-stu-id="166a0-144">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
