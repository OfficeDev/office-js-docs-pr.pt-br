# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="6ed07-101">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="6ed07-101">Tutorial: Create custom functions in Excel</span></span>

## <a name="introduction"></a><span data-ttu-id="6ed07-102">Introdução</span><span class="sxs-lookup"><span data-stu-id="6ed07-102">Introduction</span></span>

<span data-ttu-id="6ed07-103">Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="6ed07-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="6ed07-104">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="6ed07-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="6ed07-105">Você pode criar funções personalizadas que realizam tarefas simples como cálculos personalizados ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="6ed07-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="6ed07-106">Neste tutorial, você vai:</span><span class="sxs-lookup"><span data-stu-id="6ed07-106">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="6ed07-107">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="6ed07-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="6ed07-108">Usar uma função personalizada predefinida para realizar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="6ed07-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="6ed07-109">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="6ed07-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="6ed07-110">Criar uma função personalizada que transmite os dados da web em tempo real</span><span class="sxs-lookup"><span data-stu-id="6ed07-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="6ed07-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="6ed07-111">Prerequisites</span></span>

* <span data-ttu-id="6ed07-112">[Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="6ed07-112">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="6ed07-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="6ed07-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="6ed07-114">A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="6ed07-114">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command from the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="6ed07-115">Mesmo se você já instalou o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do npm.</span><span class="sxs-lookup"><span data-stu-id="6ed07-115">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="6ed07-116">Excel para Windows (versão 1810 64 bits ou posterior) ou o Excel Online</span><span class="sxs-lookup"><span data-stu-id="6ed07-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="6ed07-117">Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")</span><span class="sxs-lookup"><span data-stu-id="6ed07-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="6ed07-118">Criar um projeto com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="6ed07-118">Create a custom functions project</span></span>

 <span data-ttu-id="6ed07-119">Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6ed07-119">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="6ed07-120">Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="6ed07-120">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="6ed07-121">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="6ed07-121">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="6ed07-122">Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="6ed07-122">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="6ed07-123">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="6ed07-123">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="6ed07-124">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="6ed07-124">What do you want to name your add-in?</span></span> `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="6ed07-126">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="6ed07-126">The generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="6ed07-127">Os arquivos do project são provenientes de [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repositório GitHub.</span><span class="sxs-lookup"><span data-stu-id="6ed07-127">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="6ed07-128">Vá até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="6ed07-128">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="6ed07-129">Confie no certificado autoassinado necessário para executar este projeto.</span><span class="sxs-lookup"><span data-stu-id="6ed07-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="6ed07-130">Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="6ed07-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="6ed07-131">Crie um projeto.</span><span class="sxs-lookup"><span data-stu-id="6ed07-131">Build and run the project</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="6ed07-132">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="6ed07-132">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="6ed07-133">Se estiver usando o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web, inicie o Excel e carregue o suplemento:</span><span class="sxs-lookup"><span data-stu-id="6ed07-133">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="6ed07-134">Depois de executar esse comando, seu prompt de comando mostrará detalhes sobre o que foi feito, outra janela do npm será aberta mostrando os detalhes da compilação, e o Excel iniciará com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="6ed07-134">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="6ed07-135">Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="6ed07-135">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="6ed07-136">Se estiver usando o Excel Online para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web:</span><span class="sxs-lookup"><span data-stu-id="6ed07-136">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="6ed07-137">Depois de executar esse comando, outra janela será aberta mostrando os detalhes da compilação.</span><span class="sxs-lookup"><span data-stu-id="6ed07-137">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="6ed07-138">Para usar suas funções, abra uma nova pasta de trabalho no Office Online.</span><span class="sxs-lookup"><span data-stu-id="6ed07-138">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="6ed07-139">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="6ed07-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="6ed07-140">O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-140">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/functions/functions.js** file.</span></span> <span data-ttu-id="6ed07-141">O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="6ed07-141">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="6ed07-142">Em sua pasta de trabalho do Excel experimente a função personalizada`ADD` preenchendo as seguintes etapas no Excel:</span><span class="sxs-lookup"><span data-stu-id="6ed07-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="6ed07-143">Em uma célula, digite **= CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-143">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="6ed07-144">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="6ed07-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="6ed07-145">Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.</span><span class="sxs-lookup"><span data-stu-id="6ed07-145">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="6ed07-146">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="6ed07-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="6ed07-147">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="6ed07-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="6ed07-148">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="6ed07-148">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="6ed07-149">E se você precisasse de uma função que pode solicitar uma API de preço de uma ação e exibir o resultado na célula de uma planilha?</span><span class="sxs-lookup"><span data-stu-id="6ed07-149">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="6ed07-150">Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da web de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="6ed07-150">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="6ed07-151">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPrice` que aceita um símbolo de cotação da bolsa (por exemplo, **MSFT**) e retorna o preço dessa ação.</span><span class="sxs-lookup"><span data-stu-id="6ed07-151">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="6ed07-152">Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="6ed07-152">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="6ed07-153">No projeto**cotações** que o gerador Yeoman criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="6ed07-153">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="6ed07-154">Em **customfunctions.js**, localize a função `increment` e adicione o seguinte código imediatamente após essa função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-154">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. <span data-ttu-id="6ed07-155">Antes que o Excel possa disponibilizar essa nova função, você deve especificar metadados para descrever a função para o Excel.</span><span class="sxs-lookup"><span data-stu-id="6ed07-155">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="6ed07-156">Abrir o arquivo **config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-156">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="6ed07-157">Adicione o seguinte objeto JSON à matriz 'funções' e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6ed07-157">Add the following object to the  array within the src/functions/functions.json file and save the file.</span></span>

    <span data-ttu-id="6ed07-158">Este JSON descreve a `stockPrice` função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-158">This JSON describes the `stockPrice` function.</span></span>

    ```JSON
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

5. <span data-ttu-id="6ed07-159">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="6ed07-159">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="6ed07-160">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="6ed07-160">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="6ed07-161">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="6ed07-161">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="6ed07-162">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="6ed07-162">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="6ed07-163">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="6ed07-164">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="6ed07-164">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="6ed07-165">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-165">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="6ed07-166">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="6ed07-166">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="6ed07-167">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="6ed07-168">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="6ed07-169">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6ed07-169">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="6ed07-170">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="6ed07-171">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-171">Now, let's try out the new function.</span></span> <span data-ttu-id="6ed07-172">Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="6ed07-172">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="6ed07-173">Você verá que o resultado na célula **B1** é o preço atual das ações para uma ação da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="6ed07-173">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="6ed07-174">Criar uma função personalizada assíncrona de streaming</span><span class="sxs-lookup"><span data-stu-id="6ed07-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="6ed07-175">A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="6ed07-175">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="6ed07-176">Vamos criar uma função personalizada de fluxos de dados de uma API recebendo atualizações em tempo real sobre o preço de uma atuação.</span><span class="sxs-lookup"><span data-stu-id="6ed07-176">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="6ed07-177">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPriceStream` que solicita o preço da ação a cada 1000 milissegundos (desde que a solicitação anterior esteja concluída).</span><span class="sxs-lookup"><span data-stu-id="6ed07-177">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="6ed07-178">Enquanto a solicitação inicial está em andamento, você poderá ver o valor de espaço reservado **# OBTENDO_DADOS** na célula em que a função está sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="6ed07-178">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="6ed07-179">Quando um valor é retornado pela função, **# OBTENDO_DADOS**será substituído por esse valor na célula.</span><span class="sxs-lookup"><span data-stu-id="6ed07-179">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="6ed07-180">No projeto**cotações** que o gerador Yeoman criou, adicione o código a seguir **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6ed07-180">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="6ed07-181">Antes que o Excel possa fazer esta nova função nova disponível para usuários, especifique os metadados que descreve essa função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-181">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="6ed07-182">No projeto**cotações** que o gerador Yeoman criou, adicione o objeto a seguir na `functions`matriz em **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="6ed07-182">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="6ed07-183">Este JSON descreve a `stockPriceStream` função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-183">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="6ed07-184">Para qualquer função streaming a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do `options` objeto, como mostra este exemplo código.</span><span class="sxs-lookup"><span data-stu-id="6ed07-184">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="6ed07-185">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="6ed07-185">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="6ed07-186">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="6ed07-186">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="6ed07-187">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="6ed07-187">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="6ed07-188">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="6ed07-188">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="6ed07-189">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="6ed07-190">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="6ed07-190">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="6ed07-191">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-191">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="6ed07-192">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="6ed07-192">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="6ed07-193">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="6ed07-193">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="6ed07-194">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="6ed07-195">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6ed07-195">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span>

        4. <span data-ttu-id="6ed07-196">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="6ed07-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="6ed07-197">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="6ed07-197">Now, let's try out the new function.</span></span> <span data-ttu-id="6ed07-198">Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="6ed07-198">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="6ed07-199">Desde que o mercado de ações esteja aberto, você verá que o resultado na célula **C1** é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="6ed07-199">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="6ed07-200">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="6ed07-200">Next steps</span></span>

<span data-ttu-id="6ed07-201">Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados em tempo real da Web.</span><span class="sxs-lookup"><span data-stu-id="6ed07-201">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="6ed07-202">Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="6ed07-202">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="6ed07-203">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="6ed07-203">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="6ed07-204">Informações legais</span><span class="sxs-lookup"><span data-stu-id="6ed07-204">Legal information</span></span>

<span data-ttu-id="6ed07-205">Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="6ed07-205">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="6ed07-206">Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="6ed07-206">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="6ed07-207">O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="6ed07-207">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
