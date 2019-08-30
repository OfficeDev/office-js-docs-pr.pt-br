---
title: Tutorial de funções personalizadas do Excel
description: Neste tutorial, você criará um suplemento do Excel que contém uma função personalizada que pode executar cálculos e solicitar ou transmitir dados da web.
ms.date: 07/09/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2832df467f7e155ed026fe7f04837f18a4d2309d
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672870"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="ac37f-103">Tutorial: Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="ac37f-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="ac37f-104">Funções personalizadas permitem que você adicione novas funções do Excel definindo essas funções em JavaScript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="ac37f-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="ac37f-105">Os usuários do Excel podem acessar funções personalizadas como fariam com qualquer função nativa no Excel, como `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="ac37f-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="ac37f-106">Você pode criar funções personalizadas que realizam tarefas simples como cálculos ou tarefas mais complexas, como streaming de dados da web em tempo real em uma planilha.</span><span class="sxs-lookup"><span data-stu-id="ac37f-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="ac37f-107">Neste tutorial, você vai:</span><span class="sxs-lookup"><span data-stu-id="ac37f-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="ac37f-108">Crie um suplemento de função personalizada usando o [gerador Yeoman de suplementos do Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="ac37f-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="ac37f-109">Usar uma função personalizada predefinida para realizar um cálculo simples.</span><span class="sxs-lookup"><span data-stu-id="ac37f-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="ac37f-110">Criar uma função personalizada que solicita dados da web.</span><span class="sxs-lookup"><span data-stu-id="ac37f-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="ac37f-111">Criar uma função personalizada que transmite os dados da web em tempo real.</span><span class="sxs-lookup"><span data-stu-id="ac37f-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ac37f-112">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="ac37f-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="ac37f-113">Excel no Windows (versão 1904 ou posterior, conectada à assinatura do Office 365) ou na Web</span><span class="sxs-lookup"><span data-stu-id="ac37f-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="ac37f-114">Criar um projeto com funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac37f-114">Create a custom functions project</span></span>

 <span data-ttu-id="ac37f-115">Para começar, você criará o projeto de código para criar o suplemento função personalizada.</span><span class="sxs-lookup"><span data-stu-id="ac37f-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="ac37f-116">O [gerador Yeoman para suplementos do Office](https://www.npmjs.com/package/generator-office) configurará seu projeto com algumas funções personalizadas predefinidas que você pode experimentar. Se você já tiver executado o início rápido de funções personalizadas e gerado um projeto, continue a usar esse projeto e pule para [esta etapa](#create-a-custom-function-that-requests-data-from-the-web) .</span><span class="sxs-lookup"><span data-stu-id="ac37f-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

1. <span data-ttu-id="ac37f-117">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="ac37f-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="ac37f-118">**Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="ac37f-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="ac37f-119">**Escolha o tipo de script:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="ac37f-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="ac37f-120">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="ac37f-120">**What do you want to name your add-in?**</span></span> `starcount`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/starcountPrompt.png)
    
    <span data-ttu-id="ac37f-122">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="ac37f-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="ac37f-123">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="ac37f-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="ac37f-124">Compile o projeto.</span><span class="sxs-lookup"><span data-stu-id="ac37f-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="ac37f-125">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="ac37f-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="ac37f-126">Se você for solicitado a instalar um certificado após executar `npm run build`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="ac37f-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="ac37f-127">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="ac37f-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="ac37f-128">Você pode experimentar o suplemento função personalizada no Excel na Web ou no Windows.</span><span class="sxs-lookup"><span data-stu-id="ac37f-128">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="ac37f-129">Excel no Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="ac37f-129">Excel on Windows or Mac</span></span>](#tab/excel-windows)

<span data-ttu-id="ac37f-130">Para testar seu suplemento no Excel no Windows ou Mac, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="ac37f-130">To test your add-in in Excel on Windows or Mac, run the following command.</span></span> <span data-ttu-id="ac37f-131">Quando você executar este comando, o servidor Web local será iniciado e o Excel será aberto com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="ac37f-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="ac37f-132">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="ac37f-132">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="ac37f-133">Para testar seu suplemento no Excel em um navegador, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="ac37f-133">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="ac37f-134">Quando você executa este comando, o servidor Web local iniciará.</span><span class="sxs-lookup"><span data-stu-id="ac37f-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="ac37f-135">Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="ac37f-135">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="ac37f-136">Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ac37f-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="ac37f-137">No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-137">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="ac37f-139">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="ac37f-140">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ac37f-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="ac37f-141">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="ac37f-142">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="ac37f-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="ac37f-143">O projeto de funções personalizadas que você criou contém algumas funções personalizadas predefinidas, definidas no arquivo **./src/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="ac37f-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="ac37f-144">O arquivo **./manifest.xml** especifica que todas as funções personalizadas pertencem a `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="ac37f-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="ac37f-145">Você usará o namespace CONTOSO para acessar as funções personalizadas no Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="ac37f-146">Em seguida você vai experimentar a função personalizada `ADD` preenchendo as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="ac37f-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="ac37f-147">No Excel, vá para qualquer célula e digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="ac37f-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="ac37f-148">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="ac37f-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="ac37f-149">Executar a `CONTOSO.ADD` função, com números `10` e `200` como parâmetros de entrada, especificando o valor `=CONTOSO.ADD(10,200)` na célula e pressionando enter.</span><span class="sxs-lookup"><span data-stu-id="ac37f-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="ac37f-150">As `ADD` função personalizada calcula a soma dos dois números que você forneceu e retorna o resultado da **210**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="ac37f-151">Criar uma função personalizada que solicita dados da web</span><span class="sxs-lookup"><span data-stu-id="ac37f-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="ac37f-152">Integração de dados da Web é uma ótima maneira de ampliar o Excel por meio de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="ac37f-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="ac37f-153">Em seguida, você criará uma função `getStarCount` personalizada chamada que mostra quantas estrelas um determinado repositório do GitHub possui.</span><span class="sxs-lookup"><span data-stu-id="ac37f-153">Next you’ll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="ac37f-154">No projeto **starcount** , localize o arquivo **./src/Functions/functions.js** e abra-o no editor de código.</span><span class="sxs-lookup"><span data-stu-id="ac37f-154">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="ac37f-155">Em **Function. js**, adicione o seguinte código:</span><span class="sxs-lookup"><span data-stu-id="ac37f-155">In **function.js**, add the following code:</span></span> 

```JS
/**
  * Gets the star count for a given Github repository.
  * @customfunction 
  * @param {string} userName string name of Github user or organization.
  * @param {string} repoName string name of the Github repository.
  * @return {number} number of stars given to a Github repository.
  */
  async function getStarCount(userName, repoName) {
    try {
      //You can change this URL to any web request you want to work with.
      const url = "https://api.github.com/repos/" + userName + "/" + repoName;
      const response = await fetch(url);
      //Expect that status code is in 200-299 range
      if (!response.ok) {
        throw new Error(response.statusText)
      }
        const jsonResponse = await response.json();
        return jsonResponse.watchers_count;
    }
    catch (error) {
      return error;
    }
  }
```

3. <span data-ttu-id="ac37f-156">Execute o seguinte comando para recriar o projeto.</span><span class="sxs-lookup"><span data-stu-id="ac37f-156">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="ac37f-157">Complete as etapas a seguir (para Excel na Web, Windows ou Mac) para registrar novamente o suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-157">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="ac37f-158">Você deve concluir estas etapas para que a nova função esteja disponível.</span><span class="sxs-lookup"><span data-stu-id="ac37f-158">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="ac37f-159">Excel no Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="ac37f-159">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="ac37f-160">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-160">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="ac37f-161">No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-161">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="ac37f-162">Na lista de suplementos disponíveis, encontre a seção suplementos do **desenvolvedor** e selecione o suplemento do **starcount** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="ac37f-162">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="ac37f-163">![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-163">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="ac37f-164">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="ac37f-164">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="ac37f-165">No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.  ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-165">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="ac37f-166">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-166">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="ac37f-167">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ac37f-167">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="ac37f-168">Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-168">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="ac37f-169">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="ac37f-169">Try out the new function.</span></span> <span data-ttu-id="ac37f-170">Na célula <strong>B1</strong>, digite o texto <strong>= contoso. GETSTARCOUNT ("OfficeDev", "Excel-Custom-Functions")</strong> e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="ac37f-170">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="ac37f-171">Você verá que o resultado na célula <strong>B1</strong> é o número atual de estrelas fornecido para o [repositório GitHub do Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="ac37f-171">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="ac37f-172">Criar uma função personalizada assíncrona de streaming</span><span class="sxs-lookup"><span data-stu-id="ac37f-172">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="ac37f-173">A `getStarCount` função retorna o número de estrelas que um repositório tem em um momento específico no tempo.</span><span class="sxs-lookup"><span data-stu-id="ac37f-173">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="ac37f-174">As funções personalizadas também podem retornar dados que estão sempre mudando.</span><span class="sxs-lookup"><span data-stu-id="ac37f-174">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="ac37f-175">Essas funções são chamadas de fluxo de funções.</span><span class="sxs-lookup"><span data-stu-id="ac37f-175">These functions are called streaming functions.</span></span> <span data-ttu-id="ac37f-176">Eles devem incluir um `invocation` parâmetro que se refira à célula de onde a função foi chamada.</span><span class="sxs-lookup"><span data-stu-id="ac37f-176">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="ac37f-177">O `invocation` parâmetro é usado para atualizar o conteúdo da célula a qualquer momento.</span><span class="sxs-lookup"><span data-stu-id="ac37f-177">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="ac37f-178">No exemplo de código a seguir, você verá que há duas funções `currentTime` e. `clock`</span><span class="sxs-lookup"><span data-stu-id="ac37f-178">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="ac37f-179">A `currentTime` função é uma função estática que não usa streaming.</span><span class="sxs-lookup"><span data-stu-id="ac37f-179">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="ac37f-180">Ele retorna a data como uma cadeia de caracteres.</span><span class="sxs-lookup"><span data-stu-id="ac37f-180">It returns the date as a string.</span></span> <span data-ttu-id="ac37f-181">A `clock` função usa a `currentTime` função para fornecer a nova hora a cada segundo para uma célula no Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-181">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="ac37f-182">Ele usa `invocation.setResult` para entregar o tempo para a célula Excel e `invocation.onCanceled` para manipular o que ocorre quando a função é cancelada.</span><span class="sxs-lookup"><span data-stu-id="ac37f-182">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="ac37f-183">No projeto **starcount** , adicione o código a seguir a **./src/Functions/functions.js** e salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="ac37f-183">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

2. <span data-ttu-id="ac37f-184">Execute o seguinte comando para recriar o projeto.</span><span class="sxs-lookup"><span data-stu-id="ac37f-184">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="ac37f-185">Complete as etapas a seguir (para Excel na Web, Windows ou Mac) para registrar novamente o suplemento no Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-185">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="ac37f-186">Você deve concluir estas etapas para que a nova função esteja disponível.</span><span class="sxs-lookup"><span data-stu-id="ac37f-186">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windows-or-mactabexcel-windows"></a>[<span data-ttu-id="ac37f-187">Excel no Windows ou Mac</span><span class="sxs-lookup"><span data-stu-id="ac37f-187">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="ac37f-188">Feche o Excel e abra novamente o Excel.</span><span class="sxs-lookup"><span data-stu-id="ac37f-188">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="ac37f-189">No Excel, escolha a guia **Inserir** e, em seguida, escolha a seta para baixo localizada à direita de **meus**suplementos.  ![Inserir faixa de opções no Excel no Windows com a seta meus suplementos realçada](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="ac37f-190">Na lista de suplementos disponíveis, encontre a seção suplementos do **desenvolvedor** e selecione o suplemento do **starcount** para registrá-lo.</span><span class="sxs-lookup"><span data-stu-id="ac37f-190">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="ac37f-191">![Inserir faixa de opções no Excel no Windows com o suplemento funções personalizadas do Excel realçado na lista meus suplementos](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-191">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="ac37f-192">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="ac37f-192">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="ac37f-193">No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.  ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="ac37f-193">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="ac37f-194">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="ac37f-195">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ac37f-195">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="ac37f-196">Selecione o arquivo **manifest. XML** e escolha **abrir**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="ac37f-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="ac37f-197">Agora, vamos experimentar a nova função.</span><span class="sxs-lookup"><span data-stu-id="ac37f-197">Try out the new function.</span></span> <span data-ttu-id="ac37f-198">Na célula <strong>C1</strong>, digite o texto <strong>= contoso. RELÓGIO ())</strong> e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="ac37f-198">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK())</strong> and press enter.</span></span> <span data-ttu-id="ac37f-199">Você deve ver a data atual, que transmite uma atualização a cada segundo.</span><span class="sxs-lookup"><span data-stu-id="ac37f-199">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="ac37f-200">Embora esse relógio seja apenas um cronômetro em um loop, você pode usar a mesma ideia de definir um timer em funções mais complexas que fazem solicitações da Web para dados em tempo real.</span><span class="sxs-lookup"><span data-stu-id="ac37f-200">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="ac37f-201">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="ac37f-201">Next steps</span></span>

<span data-ttu-id="ac37f-202">Parabéns!</span><span class="sxs-lookup"><span data-stu-id="ac37f-202">Congratulations!</span></span> <span data-ttu-id="ac37f-203">Você criou um novo projeto de funções personalizadas, tentou uma função predefinida, criou uma função personalizada que solicita dados da Web e criou uma função personalizada que transmite dados.</span><span class="sxs-lookup"><span data-stu-id="ac37f-203">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="ac37f-204">Você também pode experimentar a depuração dessa função usando [as instruções de depuração da função personalizada](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="ac37f-204">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="ac37f-205">Para saber mais sobre funções personalizadas no Excel, prossiga para o seguinte artigo:</span><span class="sxs-lookup"><span data-stu-id="ac37f-205">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ac37f-206">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="ac37f-206">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)
