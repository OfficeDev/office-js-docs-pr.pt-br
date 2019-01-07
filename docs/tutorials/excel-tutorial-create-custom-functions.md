---
title: Tutorial de funções personalizadas do Excel
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 01/02/2019
ms.topic: tutorial
ms.openlocfilehash: 2a06bbff8fff23f9cb41f914a486c9cf58bea33b
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724876"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="8ac47-103">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="8ac47-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="8ac47-104">Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="8ac47-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="8ac47-105">Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="8ac47-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="8ac47-106">Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="8ac47-106">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="8ac47-107">Neste tutorial, você vai:</span><span class="sxs-lookup"><span data-stu-id="8ac47-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="8ac47-108">Criar um projeto de funções personalizadas usando o gerador Yo Office</span><span class="sxs-lookup"><span data-stu-id="8ac47-108">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="8ac47-109">Usar uma função personalizada predefinida para realizar um cálculo simples</span><span class="sxs-lookup"><span data-stu-id="8ac47-109">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="8ac47-110">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="8ac47-110">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="8ac47-111">Criar uma função personalizada que transmite os dados da web em tempo real</span><span class="sxs-lookup"><span data-stu-id="8ac47-111">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="8ac47-112">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="8ac47-112">Prerequisites</span></span>

* <span data-ttu-id="8ac47-113">[Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="8ac47-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="8ac47-114">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="8ac47-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="8ac47-115">A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="8ac47-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="8ac47-116">Mesmo se você já instalou o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do npm.</span><span class="sxs-lookup"><span data-stu-id="8ac47-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="8ac47-117">Excel para Windows (versão 1810 64 bits ou posterior) ou o Excel Online</span><span class="sxs-lookup"><span data-stu-id="8ac47-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="8ac47-118">Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")</span><span class="sxs-lookup"><span data-stu-id="8ac47-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="8ac47-119">Criar um projeto com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8ac47-119">Create a custom functions project</span></span>

 <span data-ttu-id="8ac47-120">Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8ac47-120">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="8ac47-121">Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="8ac47-121">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="8ac47-122">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="8ac47-122">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="8ac47-123">Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="8ac47-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="8ac47-124">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="8ac47-124">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="8ac47-125">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="8ac47-125">What do you want to name your add-in?</span></span> `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="8ac47-127">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="8ac47-127">The Yeoman generator will create the project files and install supporting Node components.</span></span> <span data-ttu-id="8ac47-128">Os arquivos do project são provenientes de [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repositório GitHub.</span><span class="sxs-lookup"><span data-stu-id="8ac47-128">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="8ac47-129">Vá até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="8ac47-129">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="8ac47-130">Confie no certificado autoassinado necessário para executar este projeto.</span><span class="sxs-lookup"><span data-stu-id="8ac47-130">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="8ac47-131">Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="8ac47-131">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="8ac47-132">Crie um projeto.</span><span class="sxs-lookup"><span data-stu-id="8ac47-132">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="8ac47-133">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="8ac47-133">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="8ac47-134">Se estiver usando o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web, inicie o Excel e carregue o suplemento:</span><span class="sxs-lookup"><span data-stu-id="8ac47-134">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="8ac47-135">Depois de executar esse comando, seu prompt de comando mostrará detalhes sobre o que foi feito, outra janela do npm será aberta mostrando os detalhes da compilação, e o Excel iniciará com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="8ac47-135">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="8ac47-136">Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="8ac47-136">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="8ac47-137">Se estiver usando o Excel Online para testar suas funções personalizadas, execute o seguinte comando para inciar o servidor local da web:</span><span class="sxs-lookup"><span data-stu-id="8ac47-137">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="8ac47-138">Depois de executar esse comando, outra janela será aberta mostrando os detalhes da compilação.</span><span class="sxs-lookup"><span data-stu-id="8ac47-138">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="8ac47-139">Para usar suas funções, abra uma nova pasta de trabalho no Office Online.</span><span class="sxs-lookup"><span data-stu-id="8ac47-139">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="8ac47-140">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="8ac47-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="8ac47-141">O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="8ac47-142">O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="8ac47-142">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="8ac47-143">Em sua pasta de trabalho do Excel experimente a função personalizada`ADD` preenchendo as seguintes etapas no Excel:</span><span class="sxs-lookup"><span data-stu-id="8ac47-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="8ac47-144">Em uma célula, digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="8ac47-144">Within a cell, type `=CONTOSO`.</span></span> <span data-ttu-id="8ac47-145">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="8ac47-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="8ac47-146">Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.</span><span class="sxs-lookup"><span data-stu-id="8ac47-146">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="8ac47-147">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="8ac47-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="8ac47-148">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="8ac47-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="8ac47-149">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="8ac47-149">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="8ac47-150">E se você precisasse de uma função que pode solicitar uma API de preço de uma ação e exibir o resultado na célula de uma planilha?</span><span class="sxs-lookup"><span data-stu-id="8ac47-150">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="8ac47-151">Funções personalizadas são projetadas para que você possa facilmente solicitar os dados da web de forma assíncrona.</span><span class="sxs-lookup"><span data-stu-id="8ac47-151">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="8ac47-152">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPrice` que aceita um símbolo de cotação da bolsa (por exemplo, **MSFT**) e retorna o preço dessa ação.</span><span class="sxs-lookup"><span data-stu-id="8ac47-152">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker symbol (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="8ac47-153">Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="8ac47-153">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="8ac47-154">No projeto**cotações** que o gerador Yeoman criou, localize o arquivo **src/customfunctions.js** e abra-o no seu editor de código.</span><span class="sxs-lookup"><span data-stu-id="8ac47-154">In the **stock-ticker** project that the Yeoman generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="8ac47-155">Em **customfunctions.js**, localize a função `increment` e adicione o seguinte código imediatamente após essa função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-155">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

4. <span data-ttu-id="8ac47-156">Antes que o Excel possa disponibilizar essa nova função, você deve especificar metadados para descrever a função para o Excel.</span><span class="sxs-lookup"><span data-stu-id="8ac47-156">Before Excel can make this new function available, you must specify metadata to describe the function to Excel.</span></span> <span data-ttu-id="8ac47-157">Abrir o arquivo **config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-157">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="8ac47-158">Adicione o seguinte objeto JSON à matriz 'funções' e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ac47-158">Add the following JSON object to the 'functions' array and save the file.</span></span>

    <span data-ttu-id="8ac47-159">Este JSON descreve a `stockPrice` função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-159">This JSON describes the `stockPrice` function.</span></span>

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
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. <span data-ttu-id="8ac47-160">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="8ac47-160">You must re-register the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="8ac47-161">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="8ac47-161">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="8ac47-162">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="8ac47-162">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="8ac47-163">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="8ac47-163">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="8ac47-164">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-164">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="8ac47-165">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="8ac47-165">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
            <span data-ttu-id="8ac47-166">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-166">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="8ac47-167">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="8ac47-167">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="8ac47-168">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-168">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="8ac47-169">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-169">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="8ac47-170">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="8ac47-170">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

        4. <span data-ttu-id="8ac47-171">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-171">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="8ac47-172">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-172">Now, let's try out the new function.</span></span> <span data-ttu-id="8ac47-173">Na célula **B1**, digite o texto `=CONTOSO.STOCKPRICE("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="8ac47-173">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="8ac47-174">Você verá que o resultado na célula **B1** é o preço atual das ações para uma ação da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="8ac47-174">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="8ac47-175">Criar uma função personalizada assíncrona de streaming</span><span class="sxs-lookup"><span data-stu-id="8ac47-175">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="8ac47-176">A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="8ac47-176">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="8ac47-177">Vamos criar uma função personalizada de fluxos de dados de uma API recebendo atualizações em tempo real sobre o preço de uma atuação.</span><span class="sxs-lookup"><span data-stu-id="8ac47-177">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="8ac47-178">Conclua as seguintes etapas para criar uma função personalizada chamada `stockPriceStream` que solicita o preço da ação a cada 1000 milissegundos (desde que a solicitação anterior esteja concluída).</span><span class="sxs-lookup"><span data-stu-id="8ac47-178">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="8ac47-179">Enquanto a solicitação inicial está em andamento, você poderá ver o valor de espaço reservado **# OBTENDO_DADOS** na célula em que a função está sendo exibida.</span><span class="sxs-lookup"><span data-stu-id="8ac47-179">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="8ac47-180">Quando um valor é retornado pela função, **# OBTENDO_DADOS**é substituído por esse valor na célula.</span><span class="sxs-lookup"><span data-stu-id="8ac47-180">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="8ac47-181">No projeto**cotações** que o gerador Yeoman criou, adicione o código a seguir **src/customfunctions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ac47-181">In the **stock-ticker** project that the Yeoman generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="8ac47-182">Antes que o Excel possa fazer esta nova função nova disponível para usuários, especifique os metadados que descreve essa função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-182">Before Excel can make this new function available to users, specify metadata that describes this function.</span></span> <span data-ttu-id="8ac47-183">No projeto**cotações** que o gerador Yeoman criou, adicione o objeto a seguir na `functions`matriz em **config/customfunctions.json** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="8ac47-183">In the **stock-ticker** project that the Yeoman generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="8ac47-184">Este JSON descreve a `stockPriceStream` função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-184">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="8ac47-185">Para qualquer função streaming a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do `options` objeto, como mostra este exemplo código.</span><span class="sxs-lookup"><span data-stu-id="8ac47-185">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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
                "description": "stock symbol",
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

3. <span data-ttu-id="8ac47-186">Você deverá registrar novamente o suplemento no Excel para que a nova função esteja disponível para os usuários finais.</span><span class="sxs-lookup"><span data-stu-id="8ac47-186">You must re-register the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="8ac47-187">Conclua as etapas para a plataforma que você está usando neste tutorial.</span><span class="sxs-lookup"><span data-stu-id="8ac47-187">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="8ac47-188">Se você estiver usando o Excel para Windows:</span><span class="sxs-lookup"><span data-stu-id="8ac47-188">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="8ac47-189">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="8ac47-189">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="8ac47-190">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-190">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="8ac47-191">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="8ac47-191">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
            <span data-ttu-id="8ac47-192">![Inserir faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-192">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="8ac47-193">Se você estiver usando o Excel Online:</span><span class="sxs-lookup"><span data-stu-id="8ac47-193">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="8ac47-194">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Inserir faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="8ac47-194">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="8ac47-195">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-195">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="8ac47-196">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="8ac47-196">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

        4. <span data-ttu-id="8ac47-197">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="8ac47-197">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="8ac47-198">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="8ac47-198">Now, let's try out the new function.</span></span> <span data-ttu-id="8ac47-199">Na célula **C1**, digite o texto `=CONTOSO.STOCKPRICESTREAM("MSFT")` e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="8ac47-199">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="8ac47-200">Desde que o mercado de ações esteja aberto, você verá que o resultado na célula **C1** é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="8ac47-200">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="8ac47-201">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="8ac47-201">Next steps</span></span>

<span data-ttu-id="8ac47-202">Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da web e criou uma função personalizada que transmite dados em tempo real da Web.</span><span class="sxs-lookup"><span data-stu-id="8ac47-202">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="8ac47-203">Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="8ac47-203">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="8ac47-204">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="8ac47-204">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="8ac47-205">Informações legais</span><span class="sxs-lookup"><span data-stu-id="8ac47-205">Legal information</span></span>

<span data-ttu-id="8ac47-206">Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="8ac47-206">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="8ac47-207">Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="8ac47-207">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="8ac47-208">O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="8ac47-208">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


