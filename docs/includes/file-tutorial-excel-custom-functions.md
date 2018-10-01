# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="a37f7-101">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="a37f7-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="a37f7-102">Introdução</span><span class="sxs-lookup"><span data-stu-id="a37f7-102">Introduction</span></span>

<span data-ttu-id="a37f7-103">Funções personalizadas permitem que você adicione novas funções ao Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="a37f7-103">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="a37f7-104">Os usuários no Excel podem acessar funções personalizadas tal como fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="a37f7-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="a37f7-105">É possível criar funções personalizadas que executem tarefas simples, como cálculos personalizados ou as tarefas mais complexas, como o fluxo contínuo de dados em tempo real da Web a uma planilha.</span><span class="sxs-lookup"><span data-stu-id="a37f7-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="a37f7-106">Neste tutorial, você irá:</span><span class="sxs-lookup"><span data-stu-id="a37f7-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="a37f7-107">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="a37f7-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="a37f7-108">Usar uma função personalizada pré-criada para executar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="a37f7-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="a37f7-109">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="a37f7-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="a37f7-110">Criar uma função personalizada que transmite dados em tempo real da Web</span><span class="sxs-lookup"><span data-stu-id="a37f7-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="a37f7-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="a37f7-111">Prerequisites</span></span>

* [<span data-ttu-id="a37f7-112">Node.js e npm</span><span class="sxs-lookup"><span data-stu-id="a37f7-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="a37f7-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="a37f7-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="a37f7-114">A última versão do [Yeoman](http://yeoman.io/) e o [gerador Yo Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="a37f7-114">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="a37f7-115">Para instalar essas ferramentas globalmente, execute o comando a seguir via prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="a37f7-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="a37f7-116">Excel para Windows (número da versão 10827 ou posterior) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="a37f7-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* [<span data-ttu-id="a37f7-117">Ingressar no programa Office Insider</span><span class="sxs-lookup"><span data-stu-id="a37f7-117">Join the Office Insider program</span></span>](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="a37f7-118">Criar um projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="a37f7-118">Create a custom functions project</span></span>

<span data-ttu-id="a37f7-119">Este tutorial começa usando o gerador Yo Office para criar os arquivos que você precisa para seu projeto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="a37f7-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="a37f7-120">Execute o comando a seguir e responda aos prompts da forma a seguir.</span><span class="sxs-lookup"><span data-stu-id="a37f7-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="a37f7-121">Escolha um tipo de projeto: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="a37f7-121">Choose a project type  </span></span>
    * <span data-ttu-id="a37f7-122">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="a37f7-122">Choose a script type  </span></span>
    * <span data-ttu-id="a37f7-123">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="a37f7-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![O Yo Office busca prompts de funções personalizadas](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="a37f7-125">Depois de concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes do nós de suporte.</span><span class="sxs-lookup"><span data-stu-id="a37f7-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="a37f7-126">Navegue até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="a37f7-126">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="a37f7-127">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="a37f7-127">Start the local web server.</span></span>

    * <span data-ttu-id="a37f7-128">Se for usar o Excel para Windows para testar suas funções personalizadas, execute o comando a seguir para iniciar o servidor Web local, inicie o Excel e faça o sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="a37f7-128">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="a37f7-129">Se for usar o Excel Online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local:</span><span class="sxs-lookup"><span data-stu-id="a37f7-129">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="a37f7-130">Experimentar uma função personalizada pré-criada</span><span class="sxs-lookup"><span data-stu-id="a37f7-130">Try out a prebuilt custom function</span></span>

<span data-ttu-id="a37f7-131">O projeto de funções personalizadas que você criou usando o gerador Yo Office contém algumas funções personalizadas pré-criadas, definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-131">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="a37f7-132">O arquivo **manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="a37f7-132">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="a37f7-133">Antes de poder usar qualquer uma das funções personalizadas pré-criadas, é preciso registrar o suplemento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="a37f7-133">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="a37f7-134">Faça isso concluindo as etapas para a plataforma que você usará neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="a37f7-134">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="a37f7-135">Se for usar o Excel para Windows para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="a37f7-135">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="a37f7-136">No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-136">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="a37f7-137">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-137">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="a37f7-138">![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-138">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="a37f7-139">Se for usar o Excel Online para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="a37f7-139">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="a37f7-140">No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-140">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="a37f7-141">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-141">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="a37f7-142">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="a37f7-142">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="a37f7-143">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-143">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="a37f7-144">Nesse momento, as funções personalizadas pré-criadas em seu projeto são carregadas e ficam disponíveis no Excel.</span><span class="sxs-lookup"><span data-stu-id="a37f7-144">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="a37f7-145">Experimente a função personalizada `ADD` concluindo as seguintes etapas no Excel:</span><span class="sxs-lookup"><span data-stu-id="a37f7-145">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="a37f7-146">Em uma célula, digite **=CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-146">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="a37f7-147">Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="a37f7-147">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="a37f7-148">Execute a função `CONTOSO.ADD`, com os números `10` e `200` como parâmetros de entrada, especificando o valor a seguir na célula e pressionando Enter:</span><span class="sxs-lookup"><span data-stu-id="a37f7-148">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="a37f7-149">A função personalizada `ADD` calcula a soma de dois números especificados como parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="a37f7-149">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="a37f7-150">Digitar `=CONTOSO.ADD(10,200)` deve produzir o resultado **210** na célula depois que você pressionar Enter.</span><span class="sxs-lookup"><span data-stu-id="a37f7-150">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="a37f7-151">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="a37f7-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="a37f7-152">E se você precisasse de uma função que pudesse solicitar o preço de uma ação a uma API e exibir o resultado na célula de uma planilha?</span><span class="sxs-lookup"><span data-stu-id="a37f7-152">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="a37f7-153">Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da Web de maneira assíncrona.</span><span class="sxs-lookup"><span data-stu-id="a37f7-153">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="a37f7-154">Conclua as seguintes etapas para criar uma função personalizada denominada `stockPrice`, que aceita um registrador de cotações (por exemplo, **MSFT**) e retorna o preço dessa ação.</span><span class="sxs-lookup"><span data-stu-id="a37f7-154">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="a37f7-155">Essa função personalizada usa a API comercial da IEX, a qual é gratuita e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="a37f7-155">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="a37f7-156">No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="a37f7-156">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="a37f7-157">Adicione o código a seguir a **customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-157">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="a37f7-158">Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam.</span><span class="sxs-lookup"><span data-stu-id="a37f7-158">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="a37f7-159">No projeto **cotação de ações** que o gerador Yo Office criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="a37f7-159">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="a37f7-160">Adicione o objeto a seguir à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-160">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="a37f7-161">Esse JSON descreve a função `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="a37f7-161">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="a37f7-162">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="a37f7-162">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="a37f7-163">Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="a37f7-163">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="a37f7-164">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="a37f7-164">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="a37f7-165">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="a37f7-165">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="a37f7-166">No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-166">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="a37f7-167">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-167">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="a37f7-168">![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-168">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="a37f7-169">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="a37f7-169">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="a37f7-170">No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-170">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="a37f7-171">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-171">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="a37f7-172">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="a37f7-172">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="a37f7-173">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-173">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="a37f7-174">Agora vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="a37f7-174">Now, let's try out the new function.</span></span> <span data-ttu-id="a37f7-175">Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="a37f7-175">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="a37f7-176">Você verá que o resultado da célula **B1** é o preço de estoque atual de um compartilhamento de ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="a37f7-176">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="a37f7-177">Criar uma função personalizada assíncrona de fluxo contínuo</span><span class="sxs-lookup"><span data-stu-id="a37f7-177">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="a37f7-178">A função `stockPrice` que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços de ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="a37f7-178">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="a37f7-179">Vamos criar uma função personalizada que faça o fluxo de dados de uma API para obter atualizações em tempo real do preço de uma ação.</span><span class="sxs-lookup"><span data-stu-id="a37f7-179">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="a37f7-180">Conclua as seguintes etapas para criar uma função personalizada denominada `stockPriceStream` que solicita o preço especificado de ações a cada 1.000 milissegundos (desde que a solicitação anterior tenha sido concluída).</span><span class="sxs-lookup"><span data-stu-id="a37f7-180">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="a37f7-181">Enquanto a solicitação inicial estiver em andamento, talvez você veja o valor de espaço reservado **#GETTING_DATA** na célula na qual a função está sendo chamada.</span><span class="sxs-lookup"><span data-stu-id="a37f7-181">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="a37f7-182">Quando um valor é retornado pela função, **#GETTING_DATA** será substituído pelo valor na célula.</span><span class="sxs-lookup"><span data-stu-id="a37f7-182">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="a37f7-183">No projeto **cotação de ações** que o gerador Yo Office criou, adicione código a seguir para **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-183">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="a37f7-184">Antes que o Excel possa disponibilizar essa nova função para os usuários finais, você deve especificar metadados que a descrevam.</span><span class="sxs-lookup"><span data-stu-id="a37f7-184">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="a37f7-185">No projeto **cotação de ações** que o gerador Yo Office criou, adicione o seguinte objeto à matriz `functions` do arquivo **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-185">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="a37f7-186">Esse JSON descreve a função `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="a37f7-186">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="a37f7-187">Para qualquer função de fluxo contínuo, as propriedades `stream` e `cancelable` devem ser definidas como `true` no objeto `options`, como mostrado neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="a37f7-187">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="a37f7-188">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais.</span><span class="sxs-lookup"><span data-stu-id="a37f7-188">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="a37f7-189">Conclua as etapas a seguir para a plataforma que você estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="a37f7-189">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="a37f7-190">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="a37f7-190">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="a37f7-191">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="a37f7-191">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="a37f7-192">No Excel, escolha a guia **Inserir**, depois escolha a seta para baixo localizada à direita de **Meus suplementos**.  ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-192">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="a37f7-193">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** depois selecione o suplemento **Funções personalizados do Excel** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="a37f7-193">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="a37f7-194">![Inserir a faixa de opções no Excel para Windows com o suplemento Funções personalizadas do Excel destacado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-194">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="a37f7-195">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="a37f7-195">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="a37f7-196">No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="a37f7-196">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="a37f7-197">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-197">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="a37f7-198">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="a37f7-198">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="a37f7-199">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="a37f7-199">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="a37f7-200">Agora vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="a37f7-200">Now, let's try out the new function.</span></span> <span data-ttu-id="a37f7-201">Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="a37f7-201">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="a37f7-202">Se o mercado de ações estiver aberto, o resultado na célula **C1** deve ser constantemente atualizado para refletir o preço em tempo real para um compartilhamento de ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="a37f7-202">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a37f7-203">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="a37f7-203">Next steps</span></span>

<span data-ttu-id="a37f7-204">Neste tutorial, você criou um novo projeto de funções personalizadas, testou uma função pré-criada, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que faz o fluxo de dados em tempo real na Web.</span><span class="sxs-lookup"><span data-stu-id="a37f7-204">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="a37f7-205">Para saber mais sobre as funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="a37f7-205">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="a37f7-206">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="a37f7-206">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="a37f7-207">Informações legais</span><span class="sxs-lookup"><span data-stu-id="a37f7-207">Legal Information</span></span>

<span data-ttu-id="a37f7-208">Dados fornecidos gratuitamente pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="a37f7-208">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="a37f7-209">Exibir os [Termos de uso da IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="a37f7-209">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="a37f7-210">O uso da API da IEX pela Microsoft neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="a37f7-210">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
