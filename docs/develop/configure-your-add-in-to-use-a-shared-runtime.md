---
ms.date: 12/28/2020
title: Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado
ms.prod: non-product-specific
description: Configure seu suplemento do Office para usar um tempo de execução de JavaScript compartilhado para oferecer suporte à faixa de opções adicional, painel de tarefas e recursos de funções personalizadas.
localization_priority: Priority
ms.openlocfilehash: e1248ce28a45ad63ac9b02093a39810ee042bb80
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789221"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="3d85a-103">Configure seu Suplemento do Office para usar um tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="3d85a-103">Configure your Office Add-in to use a shared JavaScript runtime</span></span>

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

<span data-ttu-id="3d85a-104">É possível configurar o Suplemento do Office para executar todo o seu código em um único tempo de execução JavaScript compartilhado (também conhecido como tempo de execução compartilhado).</span><span class="sxs-lookup"><span data-stu-id="3d85a-104">You can configure your Office Add-in to run all of its code in a single shared JavaScript runtime (also known as a shared runtime).</span></span> <span data-ttu-id="3d85a-105">Isso permite uma melhor coordenação em seu suplemento e acesso ao DOM e CORS de todas as partes de seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3d85a-105">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="3d85a-106">Ele também ativa recursos adicionais, como a execução de código quando o documento é aberto ou a ativação ou desativação de botões da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="3d85a-106">It also enables additional features such as running code when the document opens, or enabling or disabling ribbon buttons.</span></span> <span data-ttu-id="3d85a-107">Para configurar seu suplemento para usar um tempo de execução JavaScript compartilhado, siga as instruções neste artigo.</span><span class="sxs-lookup"><span data-stu-id="3d85a-107">To configure your add-in to use a shared JavaScript runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="3d85a-108">Criar o projeto de suplemento</span><span class="sxs-lookup"><span data-stu-id="3d85a-108">Create the add-in project</span></span>

<span data-ttu-id="3d85a-109">Se você estiver iniciando um novo projeto, siga estas etapas para usar o [ gerador Yeoman para suplementos do Office ](https://github.com/OfficeDev/generator-office) para criar um projeto de suplemento do Excel ou PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="3d85a-109">If you are starting a new project, follow these steps to use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create an Excel or PowerPoint add-in project.</span></span>

<span data-ttu-id="3d85a-110">Faça um dos seguintes:</span><span class="sxs-lookup"><span data-stu-id="3d85a-110">Do one of the following:</span></span>

- <span data-ttu-id="3d85a-111">Para gerar um suplemento do Excel com funções personalizadas, execute o comando `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js`.</span><span class="sxs-lookup"><span data-stu-id="3d85a-111">To generate an Excel add-in with custom functions, run the command `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js`.</span></span>
    
    <span data-ttu-id="3d85a-112">ou</span><span class="sxs-lookup"><span data-stu-id="3d85a-112">or</span></span>
    
- <span data-ttu-id="3d85a-113">Para gerar um suplemento do PowerPoint, execute o comando `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js`.</span><span class="sxs-lookup"><span data-stu-id="3d85a-113">To generate a PowerPoint add-in, run the command `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js`.</span></span>

<span data-ttu-id="3d85a-114">O gerador criará o projeto e instalará os componentes do Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="3d85a-114">The generator will create the project and install supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="3d85a-115">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="3d85a-115">Configure the manifest</span></span>

<span data-ttu-id="3d85a-116">Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-116">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span> <span data-ttu-id="3d85a-117">Estas etapas pressupõem que você gerou seu projeto usando o [Gerador Yeoman para Suplementos do Office ](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="3d85a-117">These steps assume you have generated your project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

1. <span data-ttu-id="3d85a-118">Inicie o Visual Studio Code e abra o projeto de suplemento do Excel ou PowerPoint que você gerou.</span><span class="sxs-lookup"><span data-stu-id="3d85a-118">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
2. <span data-ttu-id="3d85a-119">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-119">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="3d85a-120">Se você gerou um suplemento do Excel, atualize a seção de requisitos para usar o tempo de execução compartilhado em vez do tempo de execução da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="3d85a-120">If you generated an Excel add-in, update the requirements section to use the shared runtime instead of the custom function runtime.</span></span> <span data-ttu-id="3d85a-121">O XML deve aparecer da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="3d85a-121">The XML should appear as follows.</span></span>
    
    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
    </Requirements>
    ```
        
4. <span data-ttu-id="3d85a-122">Localize a `<VersionOverrides>`seção e adicione a seguinte`<Runtimes>` seção apenas dentro da `<Host ...>`marca.</span><span class="sxs-lookup"><span data-stu-id="3d85a-122">Find the `<VersionOverrides>` section and add the following `<Runtimes>` section just inside the `<Host ...>` tag.</span></span> <span data-ttu-id="3d85a-123">A vida útil deve ser **longa** para que o código do suplemento possa ser executado mesmo quando o painel de tarefas está fechado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-123">The lifetime needs to be **long** so that your add-in code can run even when the task pane is closed.</span></span> <span data-ttu-id="3d85a-124">O `resid`valor é **Taskpane.Url**, que faz referência ao local do arquivo **taskpane.html** especificado na ` <bt:Urls>`seção próxima à parte inferior do arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-124">The `resid` value is **Taskpane.Url**, which references the **taskpane.html** file location specified in the ` <bt:Urls>` section near the bottom of the **manifest.xml** file.</span></span>

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       ...
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

5. <span data-ttu-id="3d85a-125">Se você gerou um Suplemento do Excel com funções personalizadas, localize o elemento `<Page>`.</span><span class="sxs-lookup"><span data-stu-id="3d85a-125">If you generated an Excel add-in with custom functions, find the `<Page>` element.</span></span> <span data-ttu-id="3d85a-126">Em seguida, altere o local de origem de **Functions.Page.Url** para **Taskpane.Url**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-126">Then change the source location from **Functions.Page.Url** to **Taskpane.Url**.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

6. <span data-ttu-id="3d85a-127">Localize a marca`<FunctionFile ...>` e altere o `resid` de **Commands.Url** para **Taskpane.Url**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-127">Find the `<FunctionFile ...>` tag and change the `resid` from **Commands.Url** to  **Taskpane.Url**.</span></span> <span data-ttu-id="3d85a-128">Observe que, se você não tiver comandos de ação, não terá uma entrada **FunctionFile** e pode pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="3d85a-128">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

7. <span data-ttu-id="3d85a-129">Salve o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-129">Save the **manifest.xml** file.</span></span>

## <a name="configure-the-webpackconfigjs-file"></a><span data-ttu-id="3d85a-130">Configurar o arquivo webpack.config.js</span><span class="sxs-lookup"><span data-stu-id="3d85a-130">Configure the webpack.config.js file</span></span>

<span data-ttu-id="3d85a-131">O **webpack.config.js** construirá vários carregadores de tempo de execução.</span><span class="sxs-lookup"><span data-stu-id="3d85a-131">The **webpack.config.js** will build multiple runtime loaders.</span></span> <span data-ttu-id="3d85a-132">É necessário modificá-lo para carregar apenas o tempo de execução JavaScript compartilhado por meio do arquivo **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-132">You need to modify it to load only the shared JavaScript runtime via the **taskpane.html** file.</span></span>

1. <span data-ttu-id="3d85a-133">Inicie o Visual Studio Code e abra o projeto de suplemento do Excel ou PowerPoint que você gerou.</span><span class="sxs-lookup"><span data-stu-id="3d85a-133">Start Visual Studio Code and open the Excel or PowerPoint add-in project you generated.</span></span>
2. <span data-ttu-id="3d85a-134">Abra o arquivo **webpack.config.js**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-134">Open the **webpack.config.js** file.</span></span>
3. <span data-ttu-id="3d85a-135">Se o arquivo **webpack.config.js** tiver o seguinte código de plug-in **functions.html**, remova-o.</span><span class="sxs-lookup"><span data-stu-id="3d85a-135">If your **webpack.config.js** file has the following **functions.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

4. <span data-ttu-id="3d85a-136">Se o seu arquivo **webpack.config.js** tiver o seguinte código de plug-in **functions.html**, remova-o.</span><span class="sxs-lookup"><span data-stu-id="3d85a-136">If your **webpack.config.js** file has the following **commands.html** plugin code, remove it.</span></span>

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

5. <span data-ttu-id="3d85a-137">Se o seu projeto usou as **functions** ou os blocos de **commands**, adicione-os à lista de blocos conforme mostrado a seguir (o código a seguir é para se o seu projeto usou os dois blocos).</span><span class="sxs-lookup"><span data-stu-id="3d85a-137">If your project used either the **functions** or **commands** chunks, add them to the chunks list as shown next (the following code is for if your project used both chunks).</span></span>

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

6. <span data-ttu-id="3d85a-138">Salvar suas alterações e reconstrua o projeto.</span><span class="sxs-lookup"><span data-stu-id="3d85a-138">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

> [!NOTE]
> <span data-ttu-id="3d85a-139">Se o seu projeto tiver um arquivo **functions.html** ou um arquivo **commands.html**, eles podem ser removidos.</span><span class="sxs-lookup"><span data-stu-id="3d85a-139">If your project has a **functions.html** file or **commands.html** file, they can be removed.</span></span> <span data-ttu-id="3d85a-140">O **taskpane.html** carregará o código **functions.js** e **commands.js** no tempo de execução JavaScript compartilhado por meio das atualizações do webpack que você acabou de fazer.</span><span class="sxs-lookup"><span data-stu-id="3d85a-140">The **taskpane.html** will load the **functions.js** and **commands.js** code into the shared JavaScript runtime via the webpack updates you just made.</span></span>

## <a name="test-your-office-add-in-changes"></a><span data-ttu-id="3d85a-141">Teste as alterações do Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="3d85a-141">Test your Office Add-in changes</span></span>

<span data-ttu-id="3d85a-142">É possível confirmar que está usando o tempo de execução de JavaScript compartilhado corretamente usando as instruções a seguir.</span><span class="sxs-lookup"><span data-stu-id="3d85a-142">You can confirm that you are using the shared JavaScript runtime correctly by using the following instructions.</span></span>

1. <span data-ttu-id="3d85a-143">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-143">Open the **manifest.xml** file.</span></span>
2. <span data-ttu-id="3d85a-144">Localize a seção `<Control xsi:type="Button" id="TaskpaneButton">` e altere o seguinte `<Action ...>` XML.</span><span class="sxs-lookup"><span data-stu-id="3d85a-144">Find the `<Control xsi:type="Button" id="TaskpaneButton">` section and change the following `<Action ...>` XML.</span></span>
    
    <span data-ttu-id="3d85a-145">de:</span><span class="sxs-lookup"><span data-stu-id="3d85a-145">from:</span></span>
    
    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```
    
    <span data-ttu-id="3d85a-146">para:</span><span class="sxs-lookup"><span data-stu-id="3d85a-146">to:</span></span>
    
    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```
3. <span data-ttu-id="3d85a-147">Abra o arquivo **./src/commands/commands.js**.</span><span class="sxs-lookup"><span data-stu-id="3d85a-147">Open the **./src/commands/commands.js** file.</span></span>
4. <span data-ttu-id="3d85a-148">Substitua a função **ação** pelo código abaixo.</span><span class="sxs-lookup"><span data-stu-id="3d85a-148">Replace the **action** function with the code below.</span></span> <span data-ttu-id="3d85a-149">Isso atualizará a função para abrir e modificar o botão do painel de tarefas para incrementar um contador.</span><span class="sxs-lookup"><span data-stu-id="3d85a-149">This will update the function to open and modify the task pane button to increment a counter.</span></span> <span data-ttu-id="3d85a-150">Abrir e acessar o DOM do painel de tarefas a partir de um comando só funciona com o tempo de execução JavaScript compartilhado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-150">Opening and accessing the task pane DOM from a command only works with the shared JavaScript runtime.</span></span>
    
    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

5. <span data-ttu-id="3d85a-151">Salve suas alterações e execute o projeto.</span><span class="sxs-lookup"><span data-stu-id="3d85a-151">Save your changes and run the project.</span></span>

   ```command line
   npm start
   ```

<span data-ttu-id="3d85a-152">Cada vez que você selecionar o botão suplementos, ele mudará o texto do botão **executar** para **ir** e incrementará um contador após ele.</span><span class="sxs-lookup"><span data-stu-id="3d85a-152">Each time you select the add-ins button, it will change the **run** button text to **go** and increment a counter after it.</span></span>

## <a name="runtime-lifetime"></a><span data-ttu-id="3d85a-153">Duração do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="3d85a-153">Runtime lifetime</span></span>

<span data-ttu-id="3d85a-154">Ao adicionar o elemento `Runtime`, você também especifica uma vida útil com um valor de `long` ou `short`.</span><span class="sxs-lookup"><span data-stu-id="3d85a-154">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="3d85a-155">Defina esse valor como `long` para aproveitar os recursos, como iniciar o suplemento quando o documento for aberto, continuar executando o código após o fechamento do painel de tarefas ou usar o CORS e o DOM nas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="3d85a-155">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

>[!NOTE]
> <span data-ttu-id="3d85a-156">O valor padrão de tempo de vida é `short`, mas recomendamos usar o `long` em suplementos do Excel. Se você definir o tempo de execução como `short` neste exemplo, o suplemento do Excel será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá ser encerrado depois que o manipulador da faixa de opções for concluído.</span><span class="sxs-lookup"><span data-stu-id="3d85a-156">The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="3d85a-157">Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas pode ser encerrado quando o painel de tarefas for fechado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-157">Similarly, your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

>[!NOTE]
> <span data-ttu-id="3d85a-158">Se seu suplemento inclui o elemento `Runtimes` no manifesto (necessário para um tempo de execução compartilhado), ele utiliza o Internet Explorer 11 independentemente da versão do Windows ou do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="3d85a-158">If your add-in includes the `Runtimes` element in the manifest (required for a shared runtime), it uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span> <span data-ttu-id="3d85a-159">Para mais informações, consulte [Runtimes](../reference/manifest/runtimes.md).</span><span class="sxs-lookup"><span data-stu-id="3d85a-159">For more information, see [Runtimes](../reference/manifest/runtimes.md).</span></span>

## <a name="about-the-shared-javascript-runtime"></a><span data-ttu-id="3d85a-160">Sobre o tempo de execução de JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="3d85a-160">About the shared JavaScript runtime</span></span>

<span data-ttu-id="3d85a-161">No Windows ou Mac, seu suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados.</span><span class="sxs-lookup"><span data-stu-id="3d85a-161">On Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="3d85a-162">Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.</span><span class="sxs-lookup"><span data-stu-id="3d85a-162">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="3d85a-163">No entanto, você pode configurar o Suplemento do Office para compartilhar código no mesmo tempo de execução JavaScript (também conhecido como tempo de execução compartilhado).</span><span class="sxs-lookup"><span data-stu-id="3d85a-163">However, you can configure your Office Add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="3d85a-164">Isso permite uma melhor coordenação entre o suplemento e o acesso ao DOM e CORS do painel de tarefas de todas as partes do suplemento.</span><span class="sxs-lookup"><span data-stu-id="3d85a-164">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="3d85a-165">Configurar um tempo de execução compartilhado permite os seguintes cenários.</span><span class="sxs-lookup"><span data-stu-id="3d85a-165">Configuring a shared runtime enables the following scenarios.</span></span>

- <span data-ttu-id="3d85a-166">Seu Suplemento do Office pode usar recursos adicionais da IU:</span><span class="sxs-lookup"><span data-stu-id="3d85a-166">Your Office Add-in can use additional UI features:</span></span>
    - [<span data-ttu-id="3d85a-167">Adicionar atalhos de teclado Personalizados aos Suplementos do Office (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="3d85a-167">Add Custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
    - [<span data-ttu-id="3d85a-168">Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="3d85a-168">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
    - [<span data-ttu-id="3d85a-169">Ativar e Desativar Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="3d85a-169">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
    - [<span data-ttu-id="3d85a-170">Execute o código em seu Suplemento do Office quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="3d85a-170">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
    - [<span data-ttu-id="3d85a-171">Mostre ou oculte o painel de tarefas de seu Suplemento do Office </span><span class="sxs-lookup"><span data-stu-id="3d85a-171">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- <span data-ttu-id="3d85a-172">Para suplementos do Excel:</span><span class="sxs-lookup"><span data-stu-id="3d85a-172">For Excel add-ins:</span></span>
    - <span data-ttu-id="3d85a-173">As funções personalizadas terão suporte CORS completo.</span><span class="sxs-lookup"><span data-stu-id="3d85a-173">Custom functions will have full CORS support.</span></span>
    - <span data-ttu-id="3d85a-174">Funções personalizadas podem chamar APIs Office.js para ler dados de documentos de planilhas.</span><span class="sxs-lookup"><span data-stu-id="3d85a-174">Custom functions can call Office.js APIs to read spreadsheet document data.</span></span>

<span data-ttu-id="3d85a-175">Para o Office no Windows, o tempo de execução compartilhado requer uma instância do navegador Microsoft Internet Explorer 11, conforme explicado em [Navegadores usados ​​por suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md). Além disso, todos os botões que seu suplemento exibir na faixa de opções serão executados no mesmo tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-175">For Office on Windows, the shared runtime requires a Microsoft Internet Explorer 11 browser instance, as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="3d85a-176">A imagem a seguir mostra como as funções personalizadas, a interface do usuário da faixa de opções e o código do painel de tarefas serão executados no mesmo tempo de execução do JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3d85a-176">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Diagrama de uma função personalizada, painel de tarefas e botões da faixa de opções, todos em execução em um tempo de execução do navegador IE compartilhado no Excel](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a><span data-ttu-id="3d85a-178">Depuração</span><span class="sxs-lookup"><span data-stu-id="3d85a-178">Debugging</span></span>

<span data-ttu-id="3d85a-179">Ao usar um tempo de execução compartilhado, não é possível usar o Código do Visual Studio para depurar funções personalizadas no Excel no Windows no momento.</span><span class="sxs-lookup"><span data-stu-id="3d85a-179">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="3d85a-180">Em vez disso, você precisará usar as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="3d85a-180">You'll need to use developer tools instead.</span></span> <span data-ttu-id="3d85a-181">Para obter mais informações, consulte [Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="3d85a-181">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

### <a name="multiple-task-panes"></a><span data-ttu-id="3d85a-182">Vários painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="3d85a-182">Multiple task panes</span></span>

<span data-ttu-id="3d85a-183">Não projete seu suplemento para usar vários painéis de tarefas se você planeja usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="3d85a-183">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="3d85a-184">Um tempo de execução compartilhado tem suporte para o uso de apenas um único painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="3d85a-184">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="3d85a-185">Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.</span><span class="sxs-lookup"><span data-stu-id="3d85a-185">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="3d85a-186">Envie-nos seus comentários</span><span class="sxs-lookup"><span data-stu-id="3d85a-186">Give us feedback</span></span>

<span data-ttu-id="3d85a-187">Adoraríamos ouvir seus comentários sobre esse recurso.</span><span class="sxs-lookup"><span data-stu-id="3d85a-187">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="3d85a-188">Se você encontrar algum bug ou problema, ou tiver solicitações sobre esse recurso, informe-nos criando um problema do GitHub no [repositório office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="3d85a-188">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="3d85a-189">Confira também</span><span class="sxs-lookup"><span data-stu-id="3d85a-189">See also</span></span>

- [<span data-ttu-id="3d85a-190">Chamar APIs do Excel a partir de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="3d85a-190">Call Excel APIs from a custom function</span></span>](../excel/call-excel-apis-from-custom-function.md)
- [<span data-ttu-id="3d85a-191">Adicione atalhos de teclado personalizados aos suplementos do Office (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="3d85a-191">Add custom keyboard shortcuts to your Office Add-ins (preview)</span></span>](../design/keyboard-shortcuts.md)
- [<span data-ttu-id="3d85a-192">Crie guias contextuais Personalizadas em Suplementos do Office (pré-visualização)</span><span class="sxs-lookup"><span data-stu-id="3d85a-192">Create custom contextual tabs in Office Add-ins (preview)</span></span>](../design/contextual-tabs.md)
- [<span data-ttu-id="3d85a-193">Ativar e Desativar Comandos de Suplemento</span><span class="sxs-lookup"><span data-stu-id="3d85a-193">Enable and Disable Add-in Commands</span></span>](../design/disable-add-in-commands.md)
- [<span data-ttu-id="3d85a-194">Execute o código em seu Suplemento do Office quando o documento for aberto</span><span class="sxs-lookup"><span data-stu-id="3d85a-194">Run code in your Office Add-in when the document opens</span></span>](run-code-on-document-open.md)
- [<span data-ttu-id="3d85a-195">Mostre ou oculte o painel de tarefas de seu Suplemento do Office </span><span class="sxs-lookup"><span data-stu-id="3d85a-195">Show or hide the task pane of your Office Add-in</span></span>](show-hide-add-in.md)
- [<span data-ttu-id="3d85a-196">Tutorial: compartilhar dados e eventos entre as funções personalizadas do Excel e do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="3d85a-196">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
