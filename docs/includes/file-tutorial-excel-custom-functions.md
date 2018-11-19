# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="21501-101">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="21501-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="21501-102">Introdução</span><span class="sxs-lookup"><span data-stu-id="21501-102">Introduction</span></span>

<span data-ttu-id="21501-103">Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="21501-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="21501-104">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="21501-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="21501-105">Você pode criar funções personalizadas que realizam tarefas simples como cálculos personalizados ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="21501-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="21501-106">Neste tutorial, você vai:</span><span class="sxs-lookup"><span data-stu-id="21501-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="21501-107">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="21501-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="21501-108">Usar uma função personalizada predefinida para realizar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="21501-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="21501-109">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="21501-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="21501-110">Criar uma função personalizada que transmite os dados da web em tempo real</span><span class="sxs-lookup"><span data-stu-id="21501-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="21501-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="21501-111">Prerequisites</span></span>

* [<span data-ttu-id="21501-112">Node e npm</span><span class="sxs-lookup"><span data-stu-id="21501-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="21501-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="21501-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="21501-114">A versão mais recente da [Yeoman](http://yeoman.io/) e do [gerador Yo Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="21501-114">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="21501-115">Para instalar essas ferramentas globalmente, execute o seguinte comando para instalar o SDK:</span><span class="sxs-lookup"><span data-stu-id="21501-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="21501-116">Excel para Windows (versão 1810 ou posterior) ou o Excel Online</span><span class="sxs-lookup"><span data-stu-id="21501-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="21501-117">Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")</span><span class="sxs-lookup"><span data-stu-id="21501-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="21501-118">Criar um projeto com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="21501-118">Create a custom functions project</span></span>

<span data-ttu-id="21501-119">Você vai começar este tutorial usando o gerador Yo Office para criar os arquivos necessários para seu projeto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="21501-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="21501-120">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="21501-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="21501-121">Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="21501-121">Choose a project type  </span></span>
    * <span data-ttu-id="21501-122">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="21501-122">Choose a script type  </span></span>
    * <span data-ttu-id="21501-123">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="21501-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Yo bash Office solicita funções personalizadas](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="21501-125">Depois que você concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="21501-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="21501-126">Os arquivos do project são provenientes de [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span><span class="sxs-lookup"><span data-stu-id="21501-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="21501-127">Navegue até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="21501-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="21501-128">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="21501-128">Start the local web server.</span></span>

    * <span data-ttu-id="21501-129">Se estiver usando o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web, inicie o Excel e carregue o suplemento:</span><span class="sxs-lookup"><span data-stu-id="21501-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="21501-130">Se estiver usando o Excel Online para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web:</span><span class="sxs-lookup"><span data-stu-id="21501-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="21501-131">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="21501-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="21501-132">O projeto de funções personalizadas criado usando o gerador Yo Office contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="21501-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="21501-133">O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="21501-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="21501-134">Antes de usar as funções personalizadas predefinidas, você deverá registrar o suplemento de funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="21501-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="21501-135">Para fazer isso, conclua as etapas para a plataforma que usará neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="21501-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="21501-136">Se estiver usando o Excel para Windows para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="21501-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="21501-137">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="21501-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="21501-138">Na lista de suplementos disponíveis, localize a seção**suplementos do desenvolvedor** e selecione o suplemento**funções do Excel personalizado** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="21501-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="21501-139">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="21501-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="21501-140">Se estiver usando o Excel Online para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="21501-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="21501-141">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="21501-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="21501-142">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="21501-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="21501-143">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Office Yo.</span><span class="sxs-lookup"><span data-stu-id="21501-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="21501-144">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="21501-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="21501-145">Neste ponto, as funções personalizadas predefinidas do projeto são carregadas e estão disponíveis no Excel.</span><span class="sxs-lookup"><span data-stu-id="21501-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="21501-146">Experimentar a `ADD` função personalizada preenchendo os seguintes etapas no Excel:</span><span class="sxs-lookup"><span data-stu-id="21501-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="21501-147">Em uma célula, digite **= CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="21501-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="21501-148">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="21501-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="21501-149">Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o seguinte valor na célula e pressionando enter:</span><span class="sxs-lookup"><span data-stu-id="21501-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="21501-150">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="21501-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="21501-151">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="21501-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="21501-152">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="21501-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="21501-153">E se você precisasse de uma função que pode solicitar uma API de preço de uma ação e exibir o resultado na célula de uma planilha?</span><span class="sxs-lookup"><span data-stu-id="21501-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="21501-154">Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da web de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="21501-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="21501-155">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPrice` que aceita cotação da bolsa (por exemplo, **MSFT**) e retorna o preço dessa ação.</span><span class="sxs-lookup"><span data-stu-id="21501-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="21501-156">Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="21501-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="21501-157">No projeto**cotações** que o gerador do Office Yo criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="21501-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="21501-158">Adicione o código a seguir a **customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="21501-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="21501-159">Antes que o Excel possa fazer esta nova função nova disponível para usuários finais, você deve especificar os metadados que descreve essa função.</span><span class="sxs-lookup"><span data-stu-id="21501-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="21501-160">No projeto**cotações** que o gerador do Office Yo criou, localize o arquivo **config/customfunctions.json** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="21501-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="21501-161">Adicionar o objeto de seguir ao arquivo `functions` matriz na **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="21501-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="21501-162">Este JSON descreve a `stockPrice` função.</span><span class="sxs-lookup"><span data-stu-id="21501-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="21501-163">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="21501-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="21501-164">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="21501-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="21501-165">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="21501-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="21501-166">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="21501-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="21501-167">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="21501-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="21501-168">Na lista de suplementos disponíveis, localize a seção**suplementos do desenvolvedor** e selecione o suplemento**funções do Excel personalizado** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="21501-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="21501-169">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="21501-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="21501-170">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="21501-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="21501-171">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="21501-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="21501-172">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="21501-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="21501-173">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Office Yo.</span><span class="sxs-lookup"><span data-stu-id="21501-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="21501-174">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="21501-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="21501-175">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="21501-175">Now, let's try out the new function.</span></span> <span data-ttu-id="21501-176">Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="21501-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="21501-177">Você verá que o resultado na célula **B1** é o preço atual das ações para uma ação da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="21501-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="21501-178">Criar uma função personalizada assíncrona de streaming</span><span class="sxs-lookup"><span data-stu-id="21501-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="21501-179">A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="21501-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="21501-180">Vamos criar uma função personalizada de fluxos de dados de uma API recebendo atualizações em tempo real sobre o preço de uma atuação.</span><span class="sxs-lookup"><span data-stu-id="21501-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="21501-181">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPriceStream` que solicita o preço da ação a cada 1000 milissegundos (desde que a solicitação anterior esteja concluída).</span><span class="sxs-lookup"><span data-stu-id="21501-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="21501-182">Enquanto a solicitação inicial está em andamento, você poderá ver o valor de espaço reservado **# OBTENDO_DADOS** na célula em que a função está sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="21501-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="21501-183">Quando um valor é retornado pela função, **# OBTENDO_DADOS**será substituído por esse valor na célula.</span><span class="sxs-lookup"><span data-stu-id="21501-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="21501-184">No projeto**cotações** que o gerador do Office Yo criou, adicione o código a seguir **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="21501-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="21501-185">Antes que o Excel possa fazer esta nova função nova disponível para usuários finais, você deve especificar os metadados que descreve essa função.</span><span class="sxs-lookup"><span data-stu-id="21501-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="21501-186">No projeto**cotações** que o gerador do Office Yo criou, adicione o objeto a seguir na `functions`matriz em **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="21501-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="21501-187">Este JSON descreve a `stockPriceStream` função.</span><span class="sxs-lookup"><span data-stu-id="21501-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="21501-188">Para qualquer função streaming a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do `options` objeto, como mostra este exemplo código.</span><span class="sxs-lookup"><span data-stu-id="21501-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="21501-189">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="21501-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="21501-190">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="21501-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="21501-191">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="21501-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="21501-192">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="21501-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="21501-193">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="21501-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="21501-194">Na lista de suplementos disponíveis, localize a seção**suplementos do desenvolvedor** e selecione o suplemento**funções do Excel personalizado** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="21501-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="21501-195">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="21501-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="21501-196">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="21501-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="21501-197">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="21501-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="21501-198">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="21501-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="21501-199">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Office Yo.</span><span class="sxs-lookup"><span data-stu-id="21501-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="21501-200">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="21501-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="21501-201">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="21501-201">Now, let's try out the new function.</span></span> <span data-ttu-id="21501-202">Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="21501-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="21501-203">Desde que o mercado de ações esteja aberto, você verá que o resultado na célula **C1** é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="21501-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="21501-204">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="21501-204">Next steps</span></span>

<span data-ttu-id="21501-205">Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados em tempo real da Web.</span><span class="sxs-lookup"><span data-stu-id="21501-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="21501-206">Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="21501-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="21501-207">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="21501-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="21501-208">Informações legais</span><span class="sxs-lookup"><span data-stu-id="21501-208">Legal information</span></span>

<span data-ttu-id="21501-209">Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="21501-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="21501-210">Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="21501-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="21501-211">O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="21501-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
