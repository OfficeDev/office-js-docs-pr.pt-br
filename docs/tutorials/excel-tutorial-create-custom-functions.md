---
title: Tutorial de funções personalizadas do Excel (visualização)
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 03/19/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 76f4d88b9da39a4d71927982836ee061b329a9b3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451401"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a><span data-ttu-id="b0be9-103">Tutorial: Criar funções personalizadas no Excel (visualização)</span><span class="sxs-lookup"><span data-stu-id="b0be9-103">Tutorial: Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="b0be9-104">Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0be9-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="b0be9-105">Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="b0be9-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="b0be9-106">Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="b0be9-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="b0be9-107">Neste tutorial, você vai:</span><span class="sxs-lookup"><span data-stu-id="b0be9-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="b0be9-108">Crie um suplemento de função personalizada usando o [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="b0be9-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="b0be9-109">Usar uma função personalizada predefinida para realizar um cálculo simples.</span><span class="sxs-lookup"><span data-stu-id="b0be9-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="b0be9-110">Criar uma função personalizada que solicita dados da web.</span><span class="sxs-lookup"><span data-stu-id="b0be9-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="b0be9-111">Criar uma função personalizada que transmite os dados da web em tempo real.</span><span class="sxs-lookup"><span data-stu-id="b0be9-111">Create a custom function that streams real-time data from the web.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="b0be9-112">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="b0be9-112">Prerequisites</span></span>

* <span data-ttu-id="b0be9-113">[Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="b0be9-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="b0be9-114">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="b0be9-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="b0be9-115">A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="b0be9-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="b0be9-116">Mesmo se você já instalou o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do npm.</span><span class="sxs-lookup"><span data-stu-id="b0be9-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="b0be9-117">Excel para Windows (versão 1810 64 bits ou posterior) ou o Excel Online</span><span class="sxs-lookup"><span data-stu-id="b0be9-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="b0be9-118">Ingressar o [programa Office Insider](https://products.office.com/office-insider) (nível**Insider**, anteriormente chamado de "Insider – modo rápido")</span><span class="sxs-lookup"><span data-stu-id="b0be9-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="b0be9-119">Criar um projeto com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b0be9-119">Create a custom functions project</span></span>

 <span data-ttu-id="b0be9-120">Para começar, você criará o projeto de código para criar o suplemento função personalizada.</span><span class="sxs-lookup"><span data-stu-id="b0be9-120">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="b0be9-121">Os [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office) configurará o seu projeto com algumas funções personalizados iniciais que você pode experimentar.</span><span class="sxs-lookup"><span data-stu-id="b0be9-121">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.</span></span>

1. <span data-ttu-id="b0be9-122">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="b0be9-122">Run the following command and then answer the prompts as follows.</span></span>
    
    ```
    yo office
    ```
    
    * <span data-ttu-id="b0be9-123">Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="b0be9-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="b0be9-124">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="b0be9-124">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="b0be9-125">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="b0be9-125">What do you want to name your add-in?</span></span> `stock-ticker`
    
    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)
    
    <span data-ttu-id="b0be9-127">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node.js de suporte.</span><span class="sxs-lookup"><span data-stu-id="b0be9-127">The Yeoman generator creates the project files and installs supporting Node.js components.</span></span>

2. <span data-ttu-id="b0be9-128">Vá até a pasta do projeto.</span><span class="sxs-lookup"><span data-stu-id="b0be9-128">Go to the project folder.</span></span>
    
    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="b0be9-129">Confie no certificado autoassinado necessário para executar este projeto.</span><span class="sxs-lookup"><span data-stu-id="b0be9-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="b0be9-130">Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="b0be9-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="b0be9-131">Crie um projeto.</span><span class="sxs-lookup"><span data-stu-id="b0be9-131">Build the project.</span></span>
    
    ```
    npm run build
    ```

5. <span data-ttu-id="b0be9-132">Inicie o servidor local da web, que é executado no Node.js.</span><span class="sxs-lookup"><span data-stu-id="b0be9-132">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="b0be9-133">Você pode experimentar o suplemento função personalizada no Excel para Windows ou o Excel Online.</span><span class="sxs-lookup"><span data-stu-id="b0be9-133">You can try out the custom function add-in in Excel for Windows, or Excel Online.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="b0be9-134">Excel para Windows</span><span class="sxs-lookup"><span data-stu-id="b0be9-134">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="b0be9-135">Execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="b0be9-135">Run the following command.</span></span>

```
npm start desktop
```

<span data-ttu-id="b0be9-136">Esse comando inicia o servidor web e sideloads seu suplemento da função personalizada no Excel para Windows.</span><span class="sxs-lookup"><span data-stu-id="b0be9-136">This command starts the web server, and sideloads your custom function add-in into Excel for Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="b0be9-137">Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="b0be9-137">If your add-in does not load, check that you have completed step 3 properly.</span></span> <span data-ttu-id="b0be9-138">Você também pode habilitar o **[log de tempo de execução](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** para solucionar problemas com o arquivo de manifesto XML do seu suplemento, bem como qualquer problema de instalação ou de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="b0be9-138">You can also enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as any installation or runtime problems.</span></span> <span data-ttu-id="b0be9-139">O log de `console.log` tempo de execução grava instruções em um arquivo de log para ajudá-lo a encontrar e corrigir problemas.</span><span class="sxs-lookup"><span data-stu-id="b0be9-139">Runtime logging writes `console.log` statements to a log file to help you find and fix issues.</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="b0be9-140">Excel Online</span><span class="sxs-lookup"><span data-stu-id="b0be9-140">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="b0be9-141">Execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="b0be9-141">Run the following command.</span></span>

```
npm start web
```

<span data-ttu-id="b0be9-142">Esse comando inicia o servidor web.</span><span class="sxs-lookup"><span data-stu-id="b0be9-142">This command starts the web server.</span></span> <span data-ttu-id="b0be9-143">Faça o seguinte para sideload o seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b0be9-143">Use the following steps to sideload your add-in.</span></span>

<ol type="a">
   <li><span data-ttu-id="b0be9-144">No Excel Online, escolha a guia <strong>inserir</strong> pressione e, em seguida, escolha <strong>suplementos</strong>.</span><span class="sxs-lookup"><span data-stu-id="b0be9-144">In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.</span></span><br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><span data-ttu-id="b0be9-145">Escolha <strong>Gerenciar Meus suplementos</strong> e selecione <strong>Carregar o Suplemento</strong>.</span><span class="sxs-lookup"><span data-stu-id="b0be9-145">Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</span></span></li> 
   <li><span data-ttu-id="b0be9-146">Escolha <strong>Procurar... </strong> e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b0be9-146">Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</span></span></li> 
   <li><span data-ttu-id="b0be9-147">Selecione o arquivo <strong>manifest. XML</strong> e escolha <strong>aberto</strong>, escolha <strong>Carregar</strong>.</span><span class="sxs-lookup"><span data-stu-id="b0be9-147">Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</span></span></li>
</ol>

> [!NOTE]
> <span data-ttu-id="b0be9-148">Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="b0be9-148">If your add-in does not load, check that you have completed step 3 properly.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="b0be9-149">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="b0be9-149">Try out a prebuilt custom function</span></span>

<span data-ttu-id="b0be9-150">O projeto de funções personalizados criados alrady tem duas funções personalizadas predefinidas chamadas INCREMENTO e ADICIONAR.</span><span class="sxs-lookup"><span data-stu-id="b0be9-150">The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT.</span></span> <span data-ttu-id="b0be9-151">O código dessas funções predefinidas está no arquivo **src/Functions/functions. js** .</span><span class="sxs-lookup"><span data-stu-id="b0be9-151">The code for these prebuilt functions is in the  **src/functions/functions.js** file.</span></span> <span data-ttu-id="b0be9-152">O arquivo **./manifest.xml** especifica que todas as funções personalizadas pertencem a `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="b0be9-152">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="b0be9-153">Você usará o namespace CONTOSO para acessar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="b0be9-153">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="b0be9-154">Em seguida você vai experimentar a função personalizada `ADD` preenchendo as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="b0be9-154">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="b0be9-155">No Excel, vá para qualquer célula e digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="b0be9-155">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="b0be9-156">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="b0be9-156">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="b0be9-157">Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.</span><span class="sxs-lookup"><span data-stu-id="b0be9-157">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="b0be9-158">As `ADD` função personalizada calcula a soma dos dois números que você forneceu e retorna o resultado da **210**.</span><span class="sxs-lookup"><span data-stu-id="b0be9-158">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="b0be9-159">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="b0be9-159">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="b0be9-160">Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b0be9-160">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="b0be9-161">Em seguida, você criará uma função personalizada chamada `stockPrice` que recebe uma citação ações de uma Web API e retorna o resultado para a célula de uma planilha.</span><span class="sxs-lookup"><span data-stu-id="b0be9-161">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="b0be9-162">Esta função personalizada usa IEX Trading API, que é gratuito e não requer autenticação.</span><span class="sxs-lookup"><span data-stu-id="b0be9-162">You’ll use the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="b0be9-163">No projeto de **Cotações de ações** , localize o arquivo **src/Functions/functions. js** e abra-o no editor de código.</span><span class="sxs-lookup"><span data-stu-id="b0be9-163">In the **stock-ticker** project, find the file **src/functions/functions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="b0be9-164">Em **funções. js**, localize a `increment` função e adicione o seguinte código imediatamente após essa função.</span><span class="sxs-lookup"><span data-stu-id="b0be9-164">In **functions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **functions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    <span data-ttu-id="b0be9-165">O `CustomFunctions.associate` código associa a `id` da função com o endereço de função da `increment` em JavaScript para que o Excel possa ligar para a função.</span><span class="sxs-lookup"><span data-stu-id="b0be9-165">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `increment` in JavaScript so that Excel can call your function.</span></span>

    <span data-ttu-id="b0be9-166">Antes que o Excel possa usar a função personalizada, você precisa descrever usando metadados.</span><span class="sxs-lookup"><span data-stu-id="b0be9-166">Before Excel can use your custom function, you need to describe it using metadata.</span></span> <span data-ttu-id="b0be9-167">Você precisa definir a `id` usada no método `associate` anteriormente, além de outros metadados.</span><span class="sxs-lookup"><span data-stu-id="b0be9-167">You need to define the `id` used in the `associate` method previously, along with some other metadata.</span></span>


4. <span data-ttu-id="b0be9-168">Abra o arquivo **src/Functions/functions. JSON** .</span><span class="sxs-lookup"><span data-stu-id="b0be9-168">Open the **src/functions/functions.json** file.</span></span> <span data-ttu-id="b0be9-169">Adicione o seguinte objeto JSON à matriz 'funções' e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b0be9-169">Add the following JSON object to the 'functions' array and save the file.</span></span>

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

    <span data-ttu-id="b0be9-170">Este JSON descreve a função `stockPrice`, seus parâmetros e o tipo de resultado ela retornará.</span><span class="sxs-lookup"><span data-stu-id="b0be9-170">This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.</span></span>

5. <span data-ttu-id="b0be9-171">Registre novamente o suplemento no Excel para que a nova função esteja disponível.</span><span class="sxs-lookup"><span data-stu-id="b0be9-171">Re-register the add-in in Excel so that the new function is available.</span></span> 

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="b0be9-172">Excel para Windows</span><span class="sxs-lookup"><span data-stu-id="b0be9-172">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="b0be9-173">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="b0be9-173">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="b0be9-174">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-174">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="b0be9-175">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="b0be9-175">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="b0be9-176">![Insira a faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-176">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="b0be9-177">Excel Online</span><span class="sxs-lookup"><span data-stu-id="b0be9-177">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="b0be9-178">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-178">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="b0be9-179">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="b0be9-179">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="b0be9-180">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b0be9-180">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="b0be9-181">Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="b0be9-181">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="6">
<li> <span data-ttu-id="b0be9-182">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="b0be9-182">Try out the new function.</span></span> <span data-ttu-id="b0be9-183">Na célula <strong>B1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICE("msft")</strong> e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="b0be9-183">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="b0be9-184">Você verá que o resultado na célula <strong>B1</strong> é o preço atual das ações para uma ação da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="b0be9-184">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="b0be9-185">Criar uma função personalizada assíncrona de streaming</span><span class="sxs-lookup"><span data-stu-id="b0be9-185">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="b0be9-186">A `stockPrice` função que você acabou de criar retorna o preço de uma ação em um momento específico, mas os preços das ações estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="b0be9-186">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="b0be9-187">Em seguida, você criará uma função personalizada chamada `stockPriceStream` esse é o preço de uma ação a cada 1000 milissegundos.</span><span class="sxs-lookup"><span data-stu-id="b0be9-187">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="b0be9-188">No projeto de **Cotações de ações** , adicione o seguinte código para **src/Functions/functions. js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b0be9-188">In the **stock-ticker** project, add the following code to **src/functions/functions.js** and save the file.</span></span>

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
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    <span data-ttu-id="b0be9-189">Antes que o Excel possa usar a função personalizada, você precisa descrever usando metadados.</span><span class="sxs-lookup"><span data-stu-id="b0be9-189">Before Excel can use your custom function, you need to describe it using metadata.</span></span>
    
2. <span data-ttu-id="b0be9-190">No projeto de **Cotações de ações,** adicione o seguinte objeto à `functions` matriz no arquivo **src/Functions/functions. JSON** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="b0be9-190">In the **stock-ticker** project add the following object to the `functions` array within the **src/functions/functions.json** file and save the file.</span></span>
    
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

    <span data-ttu-id="b0be9-191">Este JSON descreve a `stockPriceStream` função.</span><span class="sxs-lookup"><span data-stu-id="b0be9-191">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="b0be9-192">Para qualquer função streaming, a propriedade `stream` e a propriedade `cancelable` devem ser definidas como `true` dentro do objeto `options`, como mostra este exemplo código.</span><span class="sxs-lookup"><span data-stu-id="b0be9-192">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

3. <span data-ttu-id="b0be9-193">Registre novamente o suplemento no Excel para que a nova função esteja disponível.</span><span class="sxs-lookup"><span data-stu-id="b0be9-193">Re-register the add-in in Excel so that the new function is available.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="b0be9-194">Excel para Windows</span><span class="sxs-lookup"><span data-stu-id="b0be9-194">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="b0be9-195">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="b0be9-195">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="b0be9-196">No Excel, escolha a guia**Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **Meus Suplementos**.  ![ Inserir faixa de opções no Excel para Windows com a seta Meus complementos realçada](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-196">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="b0be9-197">Na lista de suplementos disponíveis, localize a seção**Suplementos do desenvolvedor** e selecione o suplemento **cotações** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="b0be9-197">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="b0be9-198">![Insira a faixa de opções no Excel para Windows com o suplemento Funções Personalizadas do Excel realçado na minha lista de suplementos](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-198">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="b0be9-199">Excel Online</span><span class="sxs-lookup"><span data-stu-id="b0be9-199">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="b0be9-200">No Excel Online, escolha a guia **Inserir** e, em seguida, escolha **Suplementos**.  ![Insira a faixa de opções no Excel Online com o ícone Meus Suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="b0be9-200">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="b0be9-201">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="b0be9-201">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="b0be9-202">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="b0be9-202">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="b0be9-203">Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="b0be9-203">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="b0be9-204">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="b0be9-204">Try out the new function.</span></span> <span data-ttu-id="b0be9-205">Na célula <strong>C1</strong>, digite o texto <strong>= da CONTOSO. STOCKPRICESTREAM("msft")</strong> e pressione enter.</span><span class="sxs-lookup"><span data-stu-id="b0be9-205">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="b0be9-206">Desde que o mercado de ações esteja aberto, você verá que o resultado na célula <strong>C1</strong> é constantemente atualizado para refletir o preço em tempo uma ação das ações da Microsoft.</span><span class="sxs-lookup"><span data-stu-id="b0be9-206">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="b0be9-207">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b0be9-207">Next steps</span></span>

<span data-ttu-id="b0be9-208">Parabéns!</span><span class="sxs-lookup"><span data-stu-id="b0be9-208">Congratulations!</span></span> <span data-ttu-id="b0be9-209">Neste tutorial, você criou um novo projeto de funções personalizadas, experimentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados em tempo real da Web.</span><span class="sxs-lookup"><span data-stu-id="b0be9-209">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="b0be9-210">Você também pode experimentar a depuração dessa função usando [as instruções de depuração da função personalizada](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="b0be9-210">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="b0be9-211">Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="b0be9-211">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="b0be9-212">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="b0be9-212">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="b0be9-213">Informações legais</span><span class="sxs-lookup"><span data-stu-id="b0be9-213">Legal information</span></span>

<span data-ttu-id="b0be9-214">Dados gratuito fornecidos pela [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="b0be9-214">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="b0be9-215">Modo de exibição [termos de uso IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="b0be9-215">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="b0be9-216">O uso da Microsoft dA API IEX neste tutorial é apenas para fins educacionais.</span><span class="sxs-lookup"><span data-stu-id="b0be9-216">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


