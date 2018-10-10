# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="7ac4d-101">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="7ac4d-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="7ac4d-102">Introdução</span><span class="sxs-lookup"><span data-stu-id="7ac4d-102">Introduction</span></span>

<span data-ttu-id="7ac4d-p101">As funções personalizadas permitem adicionar novas funções ao Excel, definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que executam tarefas simples, como cálculos personalizados ou tarefas mais complexas, como a transmissão de dados em tempo real da Web para uma planilha.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="7ac4d-106">Neste tutorial, você irá:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="7ac4d-107">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="7ac4d-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="7ac4d-108">Usar uma função personalizada pré-criada para executar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="7ac4d-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="7ac4d-109">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="7ac4d-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="7ac4d-110">Criar uma função personalizada que transmite dados em tempo real da Web</span><span class="sxs-lookup"><span data-stu-id="7ac4d-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="7ac4d-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="7ac4d-111">Prerequisites</span></span>

* [<span data-ttu-id="7ac4d-112">Node.js e npm</span><span class="sxs-lookup"><span data-stu-id="7ac4d-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="7ac4d-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="7ac4d-p102">A versão mais recente do [Yeoman](http://yeoman.io/) e o [gerador Yo Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando via prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="7ac4d-116">Excel para Windows (número de build 10827 ou posterior) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="7ac4d-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* [<span data-ttu-id="7ac4d-117">Ingressar no programa Office Insider</span><span class="sxs-lookup"><span data-stu-id="7ac4d-117">Join the Office Insider program</span></span>](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="7ac4d-118">Criar um projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="7ac4d-118">Create a custom functions project</span></span>

<span data-ttu-id="7ac4d-119">Este tutorial começa usando o gerador Yo Office para criar os arquivos que você precisa para seu projeto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="7ac4d-120">Execute o comando a seguir e responda aos prompts da forma a seguir.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="7ac4d-121">Escolha um tipo de projeto: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="7ac4d-121">Choose a project type  </span></span>
    * <span data-ttu-id="7ac4d-122">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="7ac4d-122">Choose a script type  </span></span>
    * <span data-ttu-id="7ac4d-123">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="7ac4d-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![O Yo Office busca prompts de funções personalizadas](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="7ac4d-125">Depois de concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="7ac4d-126">Os arquivos de projeto vêm do repositório [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) do GitHub.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="7ac4d-127">Navegue até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="7ac4d-128">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-128">Start the local web server.</span></span>

    * <span data-ttu-id="7ac4d-129">Se for usar o Excel para Windows para testar suas funções personalizadas, execute o comando a seguir para iniciar o servidor Web local, inicie o Excel e faça o sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="7ac4d-130">Se for usar o Excel Online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="7ac4d-131">Experimentar uma função personalizada pré-criada</span><span class="sxs-lookup"><span data-stu-id="7ac4d-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="7ac4d-132">O projeto de funções personalizadas que você criou usando o gerador Yo Office contém algumas funções personalizadas pré-criadas, definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="7ac4d-133">O arquivo **manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="7ac4d-134">Antes de poder usar qualquer uma das funções personalizadas pré-criadas, é preciso registrar o suplemento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="7ac4d-135">Faça isso concluindo as etapas para a plataforma que você usará neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="7ac4d-136">Se for usar o Excel para Windows para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="7ac4d-137">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="7ac4d-138">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="7ac4d-139">![Inserir faixa de opções no Excel para Windows com o suplemento Excel Custom Functions realçado  na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="7ac4d-140">Se for usar o Excel Online para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="7ac4d-141">No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="7ac4d-142">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="7ac4d-143">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="7ac4d-144">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="7ac4d-145">Nesse momento, as funções personalizadas pré-criadas em seu projeto são carregadas e ficam disponíveis no Excel.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="7ac4d-146">Experimente a função personalizada `ADD` concluindo as seguintes etapas no Excel:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="7ac4d-147">Em uma célula, digite **=CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="7ac4d-148">Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="7ac4d-149">Execute a função `CONTOSO.ADD`, com os números `10` e `200` como parâmetros de entrada, especificando o valor a seguir na célula e pressionando Enter:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="7ac4d-150">A função personalizada `ADD` calcula a soma de dois números especificados como parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="7ac4d-151">Digitar `=CONTOSO.ADD(10,200)` deve produzir o resultado **210** na célula depois que você pressionar Enter.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="7ac4d-152">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="7ac4d-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="7ac4d-153">E se você precisasse de uma função que pudesse solicitar o preço de uma ação a uma API e exibir o resultado na célula de uma planilha?</span><span class="sxs-lookup"><span data-stu-id="7ac4d-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="7ac4d-154">Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da Web de maneira assíncrona.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="7ac4d-155">Conclua as seguintes etapas para criar uma função personalizada denominada `stockPrice`, que aceita um registrador de cotações (por exemplo, **MSFT**) e retorna o preço dessa ação.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="7ac4d-156">Essa função personalizada usa a API comercial da IEX, a qual é gratuita e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="7ac4d-157">No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="7ac4d-158">Adicione o código a seguir a **customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-158">Add the following code to **home.js** and save the file.</span></span>

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. <span data-ttu-id="7ac4d-159">Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="7ac4d-160">No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="7ac4d-161">Adicione o objeto a seguir à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="7ac4d-162">Esse JSON descreve a função `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-162">This JSON describes the `stockPrice` function.</span></span>

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. <span data-ttu-id="7ac4d-163">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="7ac4d-164">Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="7ac4d-165">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="7ac4d-166">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="7ac4d-167">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="7ac4d-168">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="7ac4d-169">![Inserir faixa de opções no Excel para Windows com o suplemento Excel Custom Functions realçado  na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="7ac4d-170">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="7ac4d-171">No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="7ac4d-172">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="7ac4d-173">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="7ac4d-174">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="7ac4d-175">Agora vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-175">Now, let's try out the new function.</span></span> <span data-ttu-id="7ac4d-176">Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="7ac4d-177">Você verá que o resultado da célula **B1** é o preço de estoque atual de um compartilhamento de ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="7ac4d-178">Criar uma função personalizada assíncrona de fluxo contínuo</span><span class="sxs-lookup"><span data-stu-id="7ac4d-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="7ac4d-179">A função `stockPrice` que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços de ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="7ac4d-180">Vamos criar uma função personalizada que faça o fluxo de dados de uma API para obter atualizações em tempo real do preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="7ac4d-181">Conclua as seguintes etapas para criar uma função personalizada denominada `stockPriceStream` que solicita o preço especificado de ações a cada 1.000 milissegundos (desde que a solicitação anterior tenha sido concluída).</span><span class="sxs-lookup"><span data-stu-id="7ac4d-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="7ac4d-182">Enquanto a solicitação inicial estiver em andamento, talvez você veja o valor de espaço reservado **#GETTING_DATA** na célula na qual a função está sendo chamada.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="7ac4d-183">Quando um valor é retornado pela função, **#GETTING_DATA** será substituído pelo valor na célula.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="7ac4d-184">No projeto **cotação de ações** que o gerador Yo Office criou, adicione código a seguir para **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. <span data-ttu-id="7ac4d-185">Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="7ac4d-186">No projeto **cotação de ações** que o gerador Yo Office criou, adicione o seguinte objeto à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="7ac4d-187">Esse JSON descreve a função `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="7ac4d-188">Para qualquer função de fluxo contínuo, as propriedades `stream` e `cancelable` devem ser definidas como `true` no objeto `options`, como mostrado neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. <span data-ttu-id="7ac4d-189">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="7ac4d-190">Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="7ac4d-191">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="7ac4d-192">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="7ac4d-193">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="7ac4d-194">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="7ac4d-195">![Inserir faixa de opções no Excel para Windows com o suplemento Excel Custom Functions realçado  na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="7ac4d-196">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="7ac4d-197">No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="7ac4d-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="7ac4d-198">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="7ac4d-199">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="7ac4d-200">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="7ac4d-201">Agora vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-201">Now, let's try out the new function.</span></span> <span data-ttu-id="7ac4d-202">Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="7ac4d-203">Se o mercado de ações estiver aberto, o resultado na célula **C1** deve ser constantemente atualizado para refletir o preço em tempo real para um compartilhamento de ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7ac4d-204">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="7ac4d-204">Next steps</span></span>

<span data-ttu-id="7ac4d-205">Neste tutorial, você criou um novo projeto de funções personalizadas, testou uma função pré-criada, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que faz o fluxo de dados em tempo real na Web.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="7ac4d-206">Para saber mais sobre as funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="7ac4d-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="7ac4d-207">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="7ac4d-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="7ac4d-208">Informações legais</span><span class="sxs-lookup"><span data-stu-id="7ac4d-208">Legal Information</span></span>

<span data-ttu-id="7ac4d-209">Dados fornecidos gratuitamente pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="7ac4d-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="7ac4d-210">Exibir os [Termos de uso da IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="7ac4d-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="7ac4d-211">O uso da API da IEX pela Microsoft neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="7ac4d-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
