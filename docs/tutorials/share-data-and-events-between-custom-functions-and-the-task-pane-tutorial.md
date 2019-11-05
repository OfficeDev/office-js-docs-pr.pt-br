---
ms.date: 11/04/2019
title: 'Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas (versão prévia)'
ms.prod: excel
description: No Excel, compartilhe dados e eventos entre as funções personalizadas e o painel de tarefas.
localization_priority: Priority
ms.openlocfilehash: dcd4bced7e1419a57256f4ec54e3ff72c0edf9ef
ms.sourcegitcommit: 42bcf9059327a8d71a7ab223805aea68be9ed6b5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/04/2019
ms.locfileid: "37962096"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a><span data-ttu-id="3ce99-103">Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e o painel de tarefas (versão prévia)</span><span class="sxs-lookup"><span data-stu-id="3ce99-103">Tutorial: Share data and events between Excel custom functions and the task pane (preview)</span></span>

<span data-ttu-id="3ce99-104">As funções personalizadas do Excel e o painel de tarefas compartilham dados globais e podem fazer chamadas de função entre si.</span><span class="sxs-lookup"><span data-stu-id="3ce99-104">Excel custom functions and the task pane share global data, and can make function calls into each other.</span></span> <span data-ttu-id="3ce99-105">Para configurar o projeto para que as funções personalizadas possam funcionar com o painel de tarefas, siga as instruções neste artigo.</span><span class="sxs-lookup"><span data-stu-id="3ce99-105">To configure your project so that custom functions can work with the task pane, follow the instructions in this article.</span></span>

> [!NOTE]
> <span data-ttu-id="3ce99-106">Os recursos descritos neste artigo estão em versão prévia e sujeitos a alterações.</span><span class="sxs-lookup"><span data-stu-id="3ce99-106">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="3ce99-107">No momento, eles não têm suporte para utilização em ambientes de produção.</span><span class="sxs-lookup"><span data-stu-id="3ce99-107">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="3ce99-108">Os recursos de versão prévia deste artigo só estão disponíveis no Excel no Windows.</span><span class="sxs-lookup"><span data-stu-id="3ce99-108">The preview features in this article are only available on Excel on Windows.</span></span> <span data-ttu-id="3ce99-109">Para experimentar os recursos de versão prévia, você precisará [ingressar no Office Insider](https://insider.office.com/pt-BR/join).</span><span class="sxs-lookup"><span data-stu-id="3ce99-109">To try the preview features, you will need to [join Office Insider](https://insider.office.com/pt-BR/join).</span></span>  <span data-ttu-id="3ce99-110">Uma boa maneira de experimentar recursos de versão prévia é usar uma assinatura do Office 365.</span><span class="sxs-lookup"><span data-stu-id="3ce99-110">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="3ce99-111">Caso ainda não tenha uma assinatura do Office 365, obtenha uma ingressando no [Programa para Desenvolvedores do Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="3ce99-111">If you don't already have an Office 365 subscription, you can get one by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="3ce99-112">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="3ce99-112">Create the add-in project</span></span>

<span data-ttu-id="3ce99-113">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="3ce99-113">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="3ce99-114">Execute o comando a seguir e responda às solicitações com as seguintes respostas:</span><span class="sxs-lookup"><span data-stu-id="3ce99-114">Run the following command and then answer the prompts as shown in the following screenshot:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="3ce99-115">Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**</span><span class="sxs-lookup"><span data-stu-id="3ce99-115">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="3ce99-116">Escolha um tipo de script: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="3ce99-116">Choose a script type: </span></span>
- <span data-ttu-id="3ce99-117">Qual será o nome do seu suplemento? **Meu suplemento do Office**</span><span class="sxs-lookup"><span data-stu-id="3ce99-117">What do you want to name your add-in? **My Office Add-in**</span></span>

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

<span data-ttu-id="3ce99-119">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="3ce99-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3ce99-120">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="3ce99-120">Configure the add-in manifest</span></span>

1. <span data-ttu-id="3ce99-121">Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-121">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="3ce99-122">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-122">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="3ce99-123">Altere a seção `<Requirements>` para usar o **CustomFunctionsRuntime** versão **1.2**, como mostrado no código a seguir.</span><span class="sxs-lookup"><span data-stu-id="3ce99-123">Change the `<Requirements>` section to use **CustomFunctionsRuntime** version **1.2** as shown in the following code.</span></span>
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. <span data-ttu-id="3ce99-124">No elemento `<Host>` da pasta de trabalho, adicione a seção `<Runtimes>` a seguir.</span><span class="sxs-lookup"><span data-stu-id="3ce99-124">Under the `<Host>` element for the workbook, add the following `<Runtimes>` section.</span></span> <span data-ttu-id="3ce99-125">O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="3ce99-125">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. <span data-ttu-id="3ce99-126">No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. <span data-ttu-id="3ce99-127">Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-127">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. <span data-ttu-id="3ce99-128">Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **TaskPaneAndCustomFunction.Url**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-128">In the `<Action>` section, change the source location from **Taskpane.Url** to **TaskPaneAndCustomFunction.Url**.</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. <span data-ttu-id="3ce99-129">Adicione uma nova **ID de Url** para **TaskPaneAndCustomFunction.Url** que aponta para **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-129">Add a new **Url id** for **TaskPaneAndCustomFunction.Url** that points to **taskpane.html**.</span></span>
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. <span data-ttu-id="3ce99-130">Salve suas alterações e recompile o projeto.</span><span class="sxs-lookup"><span data-stu-id="3ce99-130">Save your changes and rebuild the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="3ce99-131">Compartilhar o estado entre as funções personalizadas e o código do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3ce99-131">Share state between custom function and task pane code</span></span> 

<span data-ttu-id="3ce99-132">Agora que as funções personalizadas são executadas no mesmo contexto que o código do painel de tarefas, elas podem compartilhar o estado diretamente sem usar o objeto **Armazenamento**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-132">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="3ce99-133">As instruções a seguir mostram como compartilhar uma variável global entre as funções personalizadas e o código do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3ce99-133">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="3ce99-134">Criar funções personalizadas para obter ou armazenar o estado compartilhado</span><span class="sxs-lookup"><span data-stu-id="3ce99-134">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="3ce99-135">No código do Visual Studio, abra o arquivo **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-135">In Visual Studio Code, open the file **config\config.json**.</span></span>
2. <span data-ttu-id="3ce99-136">Na linha 1, insira o código a seguir na parte superior.</span><span class="sxs-lookup"><span data-stu-id="3ce99-136">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="3ce99-137">Isso inicializará uma variável global chamada **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-137">This will initialize a global variable named **sharedState**.</span></span>
    
    ```js
    window.sharedState = "empty";
    ```
    
3. <span data-ttu-id="3ce99-138">Adicione o código a seguir para criar uma função personalizada que armazena valores para a variável **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-138">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>
    
    ```js
    /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
    function storeValue(sharedValue) {
    window.sharedState = sharedValue;
    return "value stored";
    }
    ```
    
4. <span data-ttu-id="3ce99-139">Adicione o código a seguir para criar uma função personalizada que obtém o valor atual da variável **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-139">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

    ```js
    /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
    function getValue() {
    return window.sharedState;
    }
    ```
    
5. <span data-ttu-id="3ce99-140">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="3ce99-140">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="3ce99-141">Criar controles do painel de tarefas para trabalhar com dados globais</span><span class="sxs-lookup"><span data-stu-id="3ce99-141">Create task pane controls to work with global data</span></span> 

1. <span data-ttu-id="3ce99-142">Abra o arquivo**src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="3ce99-142">Open the file**src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="3ce99-143">Após o elemento de fechamento `</main>`, adicione o seguinte HTML.</span><span class="sxs-lookup"><span data-stu-id="3ce99-143">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="3ce99-144">O HTML cria duas caixas de texto e botões usados para obter ou armazenar dados globais.</span><span class="sxs-lookup"><span data-stu-id="3ce99-144">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
    <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>
    <p>Store new value to shared state</p>
    <div>
    <input type="text" id="storeBox" />
    <button onclick="storeSharedValue()">Store</button>
    </div>
     
    <p>Get shared state value</p>
    <div>
    <input type="text" id="getBox" />
    <button onclick="getSharedValue()">Get</button>
    </div>
    ```
    
3. <span data-ttu-id="3ce99-145">Antes do elemento `<body>`, adicione o seguinte script.</span><span class="sxs-lookup"><span data-stu-id="3ce99-145">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="3ce99-146">Esse código manipulará os eventos de clique do botão quando o usuário desejar armazenar ou obter os dados globais.</span><span class="sxs-lookup"><span data-stu-id="3ce99-146">This code will handle the button click events when the user wants to store or get global data.</span></span>
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
4. <span data-ttu-id="3ce99-147">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="3ce99-147">Save the file.</span></span>
5. <span data-ttu-id="3ce99-148">Compilar o projeto</span><span class="sxs-lookup"><span data-stu-id="3ce99-148">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="3ce99-149">Experimente compartilhar dados entre as funções personalizadas e o painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3ce99-149">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="3ce99-150">Inicie o projeto usando o comando a seguir.</span><span class="sxs-lookup"><span data-stu-id="3ce99-150">Start the migration by using the following command.</span></span>

    ```command&nbsp;line
    npm run start
    ```

<span data-ttu-id="3ce99-151">Após a inicialização do Excel, você pode usar os botões do painel de tarefas para armazenar ou obter os dados compartilhados.</span><span class="sxs-lookup"><span data-stu-id="3ce99-151">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="3ce99-152">Insira `=CONTOSO.GETVALUE()` em uma célula para que a função personalizada recupere os mesmos dados compartilhados.</span><span class="sxs-lookup"><span data-stu-id="3ce99-152">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="3ce99-153">Ou use `=CONTOSO.STOREVALUE(“new value”)` para alterar os dados compartilhados para um novo valor.</span><span class="sxs-lookup"><span data-stu-id="3ce99-153">Or use `=CONTOSO.STOREVALUE(“new value”)` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="3ce99-154">A configuração do seu projeto, como mostrado neste artigo, compartilhará o contexto entre as funções personalizadas e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3ce99-154">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="3ce99-155">Não há suporte para chamar APIs do Office a partir de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3ce99-155">Calling Office APIs from custom functions is not supported.</span></span> <span data-ttu-id="3ce99-156">Se você precisar interagir com o documento, implemente chamadas para as APIs do Office no [evento onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).</span><span class="sxs-lookup"><span data-stu-id="3ce99-156">If you need to interact with the document, implement calls to the Office APIs in the [onCalculated event](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).</span></span>

