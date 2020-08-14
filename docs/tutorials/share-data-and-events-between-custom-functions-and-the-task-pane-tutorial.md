---
title: 'Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas'
description: Aprenda como compartilhar dados e eventos no Excel entre as funções personalizadas e o painel de tarefas.
ms.date: 05/17/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a48d43270787648d8e5a53c885eab4b69cd8842e
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641148"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a><span data-ttu-id="80acb-103">Tutorial: Compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="80acb-103">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>

<span data-ttu-id="80acb-104">Você pode configurar o suplemento do Excel para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="80acb-104">You can configure your Excel add-in to use a shared runtime.</span></span> <span data-ttu-id="80acb-105">Isso permite compartilhar dados globais ou enviar eventos entre o painel de tarefas e as funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="80acb-105">This makes it possible to shared global data, or send events between the task pane and custom functions.</span></span>

<span data-ttu-id="80acb-106">Para a maioria dos cenários de funções personalizadas, recomendamos usar um tempo de execução compartilhada, a menos que você tenha uma razão específica para usar uma função personalizada fora do painel de tarefa (sem IU).</span><span class="sxs-lookup"><span data-stu-id="80acb-106">For most custom functions scenarios, we recommend using a shared runtime, unless you have a specific reason to use a non-task pane (UI-less) custom function.</span></span>

<span data-ttu-id="80acb-107">Este tutorial presume que você esteja familiarizado com o uso do gerador Yo do Office para criar adicionais no projetos de.</span><span class="sxs-lookup"><span data-stu-id="80acb-107">This tutorial assumes you're familiar with using the Yo Office generator to create add-in projects.</span></span> <span data-ttu-id="80acb-108">Considere concluir o [Tutorial de funções personalizadas do Excel](./excel-tutorial-create-custom-functions.md), se ainda não o fez.</span><span class="sxs-lookup"><span data-stu-id="80acb-108">Consider completing the [Excel custom functions tutorial](./excel-tutorial-create-custom-functions.md), if you haven't already.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="80acb-109">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="80acb-109">Create the add-in project</span></span>

<span data-ttu-id="80acb-110">Use o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="80acb-110">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="80acb-111">Execute o comando a seguir e responda às solicitações com as seguintes respostas:</span><span class="sxs-lookup"><span data-stu-id="80acb-111">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="80acb-112">Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**</span><span class="sxs-lookup"><span data-stu-id="80acb-112">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="80acb-113">Escolha um tipo de script: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="80acb-113">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="80acb-114">Qual será o nome do seu suplemento? **Meu suplemento do Office**</span><span class="sxs-lookup"><span data-stu-id="80acb-114">What do you want to name your add-in? **My Office Add-in**</span></span>

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

<span data-ttu-id="80acb-116">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="80acb-116">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="80acb-117">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="80acb-117">Configure the manifest</span></span>

1. <span data-ttu-id="80acb-118">Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="80acb-118">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="80acb-119">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="80acb-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="80acb-120">Localize a seção `<VersionOverrides>` e adicione a seguinte seção `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="80acb-120">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="80acb-121">O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="80acb-121">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="80acb-122">No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="80acb-122">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="80acb-123">Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="80acb-123">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="80acb-124">Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="80acb-124">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="80acb-125">Adicione um novo **ID de URL** para **ContosoAddin.Url** que aponte para **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="80acb-125">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="80acb-126">Salve suas alterações e recompile o projeto.</span><span class="sxs-lookup"><span data-stu-id="80acb-126">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a><span data-ttu-id="80acb-127">Compartilhar o estado entre as funções personalizadas e o código do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="80acb-127">Share state between custom function and task pane code</span></span>

<span data-ttu-id="80acb-128">Agora que as funções personalizadas são executadas no mesmo contexto que o código do painel de tarefas, elas podem compartilhar o estado diretamente sem usar o objeto **Armazenamento**.</span><span class="sxs-lookup"><span data-stu-id="80acb-128">Now that custom functions run in the same context as your task pane code, they can share state directly without using the **Storage** object.</span></span> <span data-ttu-id="80acb-129">As instruções a seguir mostram como compartilhar uma variável global entre as funções personalizadas e o código do painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="80acb-129">The following instructions show how to share a global variable between custom function and task pane code.</span></span>

### <a name="create-custom-functions-to-get-or-store-shared-state"></a><span data-ttu-id="80acb-130">Criar funções personalizadas para obter ou armazenar o estado compartilhado</span><span class="sxs-lookup"><span data-stu-id="80acb-130">Create custom functions to get or store shared state</span></span>

1. <span data-ttu-id="80acb-131">No código do Visual Studio, abra o arquivo **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="80acb-131">In Visual Studio Code open the file **src/functions/functions.js**.</span></span>
2. <span data-ttu-id="80acb-132">Na linha 1, insira o código a seguir na parte superior.</span><span class="sxs-lookup"><span data-stu-id="80acb-132">On line 1, insert the following code at the very top.</span></span> <span data-ttu-id="80acb-133">Isso inicializará uma variável global chamada **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="80acb-133">This will initialize a global variable named **sharedState**.</span></span>

   ```js
   window.sharedState = "empty";
   ```

3. <span data-ttu-id="80acb-134">Adicione o código a seguir para criar uma função personalizada que armazena valores para a variável **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="80acb-134">Add the following code to create a custom function that stores values to the **sharedState** variable.</span></span>

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

4. <span data-ttu-id="80acb-135">Adicione o código a seguir para criar uma função personalizada que obtém o valor atual da variável **sharedState**.</span><span class="sxs-lookup"><span data-stu-id="80acb-135">Add the following code to create a custom function that gets the current value of the **sharedState** variable.</span></span>

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

5. <span data-ttu-id="80acb-136">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="80acb-136">Save the file.</span></span>

### <a name="create-task-pane-controls-to-work-with-global-data"></a><span data-ttu-id="80acb-137">Criar controles do painel de tarefas para trabalhar com dados globais</span><span class="sxs-lookup"><span data-stu-id="80acb-137">Create task pane controls to work with global data</span></span>

1. <span data-ttu-id="80acb-138">Abra o arquivo **src/taskpane/taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="80acb-138">Open the file **src/taskpane/taskpane.html**.</span></span>
2. <span data-ttu-id="80acb-139">Adicione o seguinte elemento de script antes do elemento `</head>`.</span><span class="sxs-lookup"><span data-stu-id="80acb-139">Add the following script element just before the `</head>` element.</span></span>

   ```html
   <script src="functions.js"></script>
   ```

3. <span data-ttu-id="80acb-140">Após o elemento de fechamento `</main>`, adicione o seguinte HTML.</span><span class="sxs-lookup"><span data-stu-id="80acb-140">After the closing `</main>` element, add the following HTML.</span></span> <span data-ttu-id="80acb-141">O HTML cria duas caixas de texto e botões usados para obter ou armazenar dados globais.</span><span class="sxs-lookup"><span data-stu-id="80acb-141">The HTML creates two text boxes and buttons used to get or store global data.</span></span>

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
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

4. <span data-ttu-id="80acb-142">Antes do elemento `<body>`, adicione o seguinte script.</span><span class="sxs-lookup"><span data-stu-id="80acb-142">Before the `<body>` element add the following script.</span></span> <span data-ttu-id="80acb-143">Esse código manipulará os eventos de clique do botão quando o usuário desejar armazenar ou obter os dados globais.</span><span class="sxs-lookup"><span data-stu-id="80acb-143">This code will handle the button click events when the user wants to store or get global data.</span></span>

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

5. <span data-ttu-id="80acb-144">Salve o arquivo.</span><span class="sxs-lookup"><span data-stu-id="80acb-144">Save the file.</span></span>
6. <span data-ttu-id="80acb-145">Compilar o projeto</span><span class="sxs-lookup"><span data-stu-id="80acb-145">Build the project</span></span>

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a><span data-ttu-id="80acb-146">Experimente compartilhar dados entre as funções personalizadas e o painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="80acb-146">Try sharing data between the custom functions and task pane</span></span>

- <span data-ttu-id="80acb-147">Inicie o projeto usando o comando a seguir.</span><span class="sxs-lookup"><span data-stu-id="80acb-147">Start the project by using the following command.</span></span>

  ```command line
  npm run start
  ```

<span data-ttu-id="80acb-148">Após a inicialização do Excel, você pode usar os botões do painel de tarefas para armazenar ou obter os dados compartilhados.</span><span class="sxs-lookup"><span data-stu-id="80acb-148">Once Excel starts, you can use the task pane buttons to store or get shared data.</span></span> <span data-ttu-id="80acb-149">Insira `=CONTOSO.GETVALUE()` em uma célula para que a função personalizada recupere os mesmos dados compartilhados.</span><span class="sxs-lookup"><span data-stu-id="80acb-149">Enter `=CONTOSO.GETVALUE()` into a cell for the custom function to retrieve the same shared data.</span></span> <span data-ttu-id="80acb-150">Ou use `=CONTOSO.STOREVALUE("new value")` para alterar os dados compartilhados para um novo valor.</span><span class="sxs-lookup"><span data-stu-id="80acb-150">Or use `=CONTOSO.STOREVALUE("new value")` to change the shared data to a new value.</span></span>

> [!NOTE]
> <span data-ttu-id="80acb-151">A configuração do seu projeto, como mostrado neste artigo, compartilhará o contexto entre as funções personalizadas e o painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="80acb-151">Configuring your project as shown in this article will share context between custom functions and the task pane.</span></span> <span data-ttu-id="80acb-152">É possível chamar algumas APIs do Office a partir de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="80acb-152">Calling some Office APIs from custom functions is possible.</span></span> <span data-ttu-id="80acb-153">[Consulte chamada de APIs do Microsoft Excel a partir de uma função personalizada](../excel/call-excel-apis-from-custom-function.md) para mais detalhes.</span><span class="sxs-lookup"><span data-stu-id="80acb-153">[See Call Microsoft Excel APIs from a custom function](../excel/call-excel-apis-from-custom-function.md) for more details.</span></span>
