# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="51b99-101">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="51b99-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="51b99-102">Introdução</span><span class="sxs-lookup"><span data-stu-id="51b99-102">Introduction</span></span>

<span data-ttu-id="51b99-p101">As funções personalizadas permitem adicionar novas funções ao Excel, definindo essas funções em JavaScript como parte de um suplemento. Os usuários do Excel podem acessar funções personalizadas da mesma forma que qualquer função nativa no Excel, como `SUM()`. Você pode criar funções personalizadas que executam tarefas simples, como cálculos personalizados ou tarefas mais complexas, como a transmissão de dados em tempo real da Web para uma planilha.</span><span class="sxs-lookup"><span data-stu-id="51b99-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="51b99-106">Neste tutorial, você irá:</span><span class="sxs-lookup"><span data-stu-id="51b99-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="51b99-107">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="51b99-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="51b99-108">Usar uma função personalizada pré-criada para executar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="51b99-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="51b99-109">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="51b99-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="51b99-110">Criar uma função personalizada que transmite dados em tempo real da Web</span><span class="sxs-lookup"><span data-stu-id="51b99-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="51b99-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="51b99-111">Prerequisites</span></span>

* [<span data-ttu-id="51b99-112">Node.js e npm</span><span class="sxs-lookup"><span data-stu-id="51b99-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="51b99-113">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="51b99-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="51b99-p102">A versão mais recente do [Yeoman](http://yeoman.io/) e o [gerador Yo Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando via prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="51b99-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="51b99-116">Excel para Windows (build 10827 ou posterior) ou Excel Online</span><span class="sxs-lookup"><span data-stu-id="51b99-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* <span data-ttu-id="51b99-117">Faça parte do [programa Office Insider](https://products.office.com/office-insider) (**Insider** level, antigo "Insider Fast")</span><span class="sxs-lookup"><span data-stu-id="51b99-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="51b99-118">Criar um projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="51b99-118">Create a custom functions project</span></span>

<span data-ttu-id="51b99-119">Este tutorial começa usando o gerador Yo Office para criar os arquivos que você precisa para seu projeto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="51b99-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="51b99-120">Execute o comando a seguir e responda aos prompts da forma a seguir.</span><span class="sxs-lookup"><span data-stu-id="51b99-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="51b99-121">Escolha um tipo de projeto: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="51b99-121">Choose a project type:`Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="51b99-122">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="51b99-122">Choose a script type:`JavaScript`</span></span>
    * <span data-ttu-id="51b99-123">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="51b99-123">What do you want to name your add-in?:</span></span> `stock-ticker`

    ![O Yo Office busca prompts de funções personalizadas](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="51b99-p103">Após concluir o assistente, o gerador criará os arquivos do projeto e instalará os componentes do nó de suporte. Os arquivos do projeto podem ser encontrados no repositório [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) do GitHub.</span><span class="sxs-lookup"><span data-stu-id="51b99-p103">After you complete the wizard, the generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="51b99-127">Navegue até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="51b99-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="51b99-128">Inicie o servidor Web local.</span><span class="sxs-lookup"><span data-stu-id="51b99-128">Start the local web server.</span></span>

    * <span data-ttu-id="51b99-129">Se for usar o Excel para Windows para testar suas funções personalizadas, execute o comando a seguir para iniciar o servidor Web local, inicie o Excel e faça o sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="51b99-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="51b99-130">Se for usar o Excel Online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local:</span><span class="sxs-lookup"><span data-stu-id="51b99-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="51b99-131">Experimentar uma função personalizada pré-criada</span><span class="sxs-lookup"><span data-stu-id="51b99-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="51b99-p104">O projeto de funções personalizadas criado com o gerador Yo Office contém algumas funções personalizadas pré-criados, definidas dentro do arquivo **src/customfunction.js**. O arquivo **manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="51b99-p104">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="51b99-p105">Antes de poder usar qualquer uma das funções personalizadas pré-criadas, você deve registrar o suplemento de funções personalizadas no Excel. Para isso, siga as etapas deste tutorial para a plataforma que você vai usar.</span><span class="sxs-lookup"><span data-stu-id="51b99-p105">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel. Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="51b99-136">Se for usar o Excel para Windows para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="51b99-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="51b99-137">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="51b99-p106">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-p106">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="51b99-140">Se for usar o Excel Online para testar suas funções personalizadas:</span><span class="sxs-lookup"><span data-stu-id="51b99-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="51b99-141">No Excel Online, escolha a guia **Inserir**, depois escolha **Suplementos**.  ![Inserir a faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="51b99-142">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="51b99-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="51b99-143">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="51b99-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="51b99-144">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="51b99-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="51b99-p107">Depois disso, as funções personalizadas pré-criadas do seu projeto já estarão carregadas e disponíveis dentro do Excel. Experimente a função personalizada `ADD` seguindo estas no Excel:</span><span class="sxs-lookup"><span data-stu-id="51b99-p107">At this point, the prebuilt custom functions in your project are loaded and available within Excel. Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="51b99-p108">Dentro de uma célula, digite **= CONTOSO**. Observe que o menu de preenchimento automático mostra a lista de todas as funções no namespace `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="51b99-p108">Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="51b99-149">Execute a função `CONTOSO.ADD`, com os números `10` e `200` como parâmetros de entrada, especificando o valor a seguir na célula e pressionando Enter:</span><span class="sxs-lookup"><span data-stu-id="51b99-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="51b99-p109">A função personalizada `ADD` calcula a soma dos dois números especificados por você como parâmetros de entrada. Ao digitar `=CONTOSO.ADD(10,200)` e pressionar Enter, o resultado **210** deve aparecer na célula.</span><span class="sxs-lookup"><span data-stu-id="51b99-p109">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="51b99-152">Criar uma função personalizada que solicita dados da Web</span><span class="sxs-lookup"><span data-stu-id="51b99-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="51b99-p110">E se você precisar de uma função que solicita o preço de uma ação a uma API e exibe o resultado em uma célula da planilha? Funções personalizadas são projetadas para que você possa facilmente solicitar dados da web de maneira assíncrona.</span><span class="sxs-lookup"><span data-stu-id="51b99-p110">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="51b99-p111">Complete as etapas a seguir para criar uma função personalizada denominada `stockPrice` que aceita um ticker de ações (como **MSFT**) e retorna o preço da ação. Essa função personalizada usa a API IEX de trading, que é gratuita e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="51b99-p111">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="51b99-157">No projeto **stock-ticker** criado pelo gerador Yo Office, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="51b99-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="51b99-158">Adicione o código a seguir a **customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="51b99-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="51b99-p112">Para que o Excel possa disponibilizar essa nova função para os usuários finais, você deve primeiro especificar metadados que a descrevem. No projeto **stock-ticker** criado pelo gerador Yo Office, localize o arquivo **config/customfunctions.json** e abra-o no seu editor de código. Adicione o seguinte objeto à matriz `functions` dentro do arquivo **config/customfunctions.json** e salve-o.</span><span class="sxs-lookup"><span data-stu-id="51b99-p112">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor. Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="51b99-162">Esse JSON descreve a função `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="51b99-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="51b99-p113">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="51b99-p113">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="51b99-165">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="51b99-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="51b99-166">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="51b99-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="51b99-167">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="51b99-p114">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-p114">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="51b99-170">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="51b99-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="51b99-171">No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="51b99-172">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="51b99-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="51b99-173">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="51b99-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="51b99-174">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="51b99-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="51b99-p115">Agora, vamos experimentar a nova função. Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione Enter. O resultado da célula **B1** deve ser o preço atual de uma ação da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="51b99-p115">Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="51b99-178">Criar uma função personalizada assíncrona de fluxo contínuo</span><span class="sxs-lookup"><span data-stu-id="51b99-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="51b99-p116">A função `stockPrice` que você acaba de criar retorna o preço de uma ação em um momento específico, mas os preços de ações estão em constante mudança. Agora, vamos criar uma função personalizada que transmite dados de uma API para obter atualizações do preço de uma ação em tempo real.</span><span class="sxs-lookup"><span data-stu-id="51b99-p116">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="51b99-p117">Conclua as etapas a seguir para criar uma função personalizada denominada `stockPriceStream` que solicita o preço da ação especificada a cada 1000 milissegundos (desde que a solicitação anterior tenha sido concluída). Enquanto a solicitação inicial estiver em andamento, talvez você veja o valor espaço reservado **#GETTING_DATA** na célula onde a função está sendo chamada. Quando um valor é retornado pela função, **#GETTING_DATA** é substituído por esse valor.</span><span class="sxs-lookup"><span data-stu-id="51b99-p117">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="51b99-184">No projeto **stock-ticker** criado pelo gerador Yo Office, adicione código a seguir para **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="51b99-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="51b99-p118">Para que o Excel possa disponibilizar essa nova função para os usuários finais, você deve primeiro especificar metadados que a descrevem. No projeto **stock-ticker** criado pelo gerador Yo Office, adicione o objeto a seguir à matriz `functions` no arquivo **config/customfunctions.json** e salve-o.</span><span class="sxs-lookup"><span data-stu-id="51b99-p118">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="51b99-p119">Este JSON descreve a função `stockPriceStream`. Para qualquer função de fluxo contínuo, as propriedades `stream` e `cancelable` devem ser definidas como `true` no objeto `options`, como mostrado neste exemplo de código.</span><span class="sxs-lookup"><span data-stu-id="51b99-p119">This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="51b99-p120">Você deve registrar novamente o suplemento no Excel para que a nova função fique disponível aos usuários finais. Conclua as etapas a seguir para a plataforma que estiver usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="51b99-p120">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="51b99-191">Se estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="51b99-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="51b99-192">Feche e reabra o Excel.</span><span class="sxs-lookup"><span data-stu-id="51b99-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="51b99-193">No Excel, escolha a guia **Inserir** e depois escolha a seta para baixo localizada à direita de **Meus suplementos**.   ![Insira a faixa de opções no Excel para Windows com a seta Meus suplementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="51b99-p121">Na lista de suplementos disponíveis, localize a seção **Suplementos do desenvolvedor** e selecione o suplemento de **Funções personalizados do Excel** para registrá-lo.  ![Insira a faixa de opções no Excel para Windows com o Suplemento de funções personalizados do Excel realçado na lista Meus suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-p121">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="51b99-196">Se estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="51b99-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="51b99-197">No Excel Online, escolha a guia **Inserir** e depois escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="51b99-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="51b99-198">Escolha **Gerenciar Meus Suplementos** e selecione **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="51b99-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="51b99-199">Escolha **Procurar...** e navegue até o diretório raiz do projeto criado o gerador Yo Office.</span><span class="sxs-lookup"><span data-stu-id="51b99-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="51b99-200">Selecione o arquivo **manifest.xml** e escolha **Abrir**, depois selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="51b99-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="51b99-p122">Agora, vamos experimentar a nova função. Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione Enter. Se o mercado de ações estiver aberto, o resultado na célula **C1** será constantemente atualizado para refletir o preço de uma ação da Microsoft em tempo real.</span><span class="sxs-lookup"><span data-stu-id="51b99-p122">Now, let's try out the new function. In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter. Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="51b99-204">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="51b99-204">Next steps</span></span>

<span data-ttu-id="51b99-p123">Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função pré-criada, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados da web em tempo real. Para saber mais sobre as funções personalizadas no Excel, veja o artigo a seguir:</span><span class="sxs-lookup"><span data-stu-id="51b99-p123">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="51b99-207">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="51b99-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="51b99-208">Informações jurídicas</span><span class="sxs-lookup"><span data-stu-id="51b99-208">Legal Information</span></span>

<span data-ttu-id="51b99-p124">Dados fornecidos gratuitamente pelo [IEX](https://iextrading.com/developer/). Verifique os [Termos de uso do IEX](https://iextrading.com/api-exhibit-a/). O uso da API IEX neste tutorial da Microsoft é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="51b99-p124">Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
