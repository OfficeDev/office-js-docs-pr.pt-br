---
ms.date: 05/17/2020
title: Configurar seu suplemento do Excel para compartilhar o tempo de execução do navegador
ms.prod: excel
description: Configure o suplemento do Excel para compartilhar o tempo de execução do navegador e executar a faixa de opções, o painel de tarefas e o código de função personalizado no mesmo tempo de execução.
localization_priority: Priority
ms.openlocfilehash: 166e0f9f2c2fe99b303484d5212e5d8b23726d6a
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44276012"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime"></a><span data-ttu-id="b31ec-103">Configurar seu suplemento do Excel para usar um tempo de execução do JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="b31ec-103">Configure your Excel add-in to use a shared JavaScript runtime</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="b31ec-104">Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados.</span><span class="sxs-lookup"><span data-stu-id="b31ec-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="b31ec-105">Isso cria limitações, como não ser capaz de compartilhar facilmente dados globais e não ter acesso a todas as funcionalidades de CORS de uma função personalizada.</span><span class="sxs-lookup"><span data-stu-id="b31ec-105">This creates limitations such as not being able to easily share global data, and not having access to all CORS functionality from a custom function.</span></span>

<span data-ttu-id="b31ec-106">No entanto, você pode configurar o suplemento do Excel para compartilhar código em um tempo de execução JavaScript compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="b31ec-107">Isso permite uma melhor coordenação entre seu suplemento e acesso ao DOM e CORS de todas as partes do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b31ec-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="b31ec-108">Também permite executar o código quando o documento é aberto ou executar o código enquanto o painel de tarefas está fechado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="b31ec-109">Para configurar seu suplemento para usar um tempo de execução compartilhado, siga as instruções neste artigo.</span><span class="sxs-lookup"><span data-stu-id="b31ec-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="b31ec-110">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="b31ec-110">Create the add-in project</span></span>

<span data-ttu-id="b31ec-111">Se você estiver iniciando um novo projeto, siga estas etapas para usar o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="b31ec-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="b31ec-112">Execute o comando a seguir e responda às solicitações com as seguintes respostas:</span><span class="sxs-lookup"><span data-stu-id="b31ec-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="b31ec-113">Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**</span><span class="sxs-lookup"><span data-stu-id="b31ec-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="b31ec-114">Escolha um tipo de script: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="b31ec-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="b31ec-115">Qual será o nome do seu suplemento? **Meu suplemento do Office**</span><span class="sxs-lookup"><span data-stu-id="b31ec-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

<span data-ttu-id="b31ec-117">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="b31ec-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b31ec-118">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="b31ec-118">Configure the manifest</span></span>

<span data-ttu-id="b31ec-119">Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="b31ec-120">Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="b31ec-121">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="b31ec-122">Localize a seção `<VersionOverrides>` e adicione a seguinte seção `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="b31ec-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="b31ec-123">O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="b31ec-124">O resid é `ContosoAddin.Url`, que faz referência a uma sequência na seção de recursos posteriormente.</span><span class="sxs-lookup"><span data-stu-id="b31ec-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="b31ec-125">Você pode usar qualquer valor de resid que desejar, mas deve corresponder ao resid dos outros elementos nos elementos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b31ec-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="b31ec-126">No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="b31ec-127">Este resid corresponde ao elemento resid `<Runtime>`.</span><span class="sxs-lookup"><span data-stu-id="b31ec-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="b31ec-128">Observe que, se você não tiver funções personalizadas, não terá uma entrada **Page** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b31ec-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="b31ec-129">Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="b31ec-130">Observe que, se você não possui comandos de ação, não terá uma entrada **FunctionFile** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b31ec-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="b31ec-131">Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="b31ec-132">Observe que, se você não tiver um painel de tarefas, não terá uma ação **ShowTaskpane** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b31ec-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="b31ec-133">Adicione um novo **ID de URL** para **ContosoAddin.Url** que aponte para **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="b31ec-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="b31ec-134">Salve suas alterações e recompile o projeto.</span><span class="sxs-lookup"><span data-stu-id="b31ec-134">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="b31ec-135">Duração do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="b31ec-135">Runtime lifetime</span></span>

<span data-ttu-id="b31ec-136">Ao adicionar o elemento `Runtime`, você também especifica uma vida útil com um valor de `long` ou `short`.</span><span class="sxs-lookup"><span data-stu-id="b31ec-136">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="b31ec-137">Defina esse valor como `long` para aproveitar os recursos, como iniciar o suplemento quando o documento for aberto, continuar executando o código após o fechamento do painel de tarefas ou usar o CORS e o DOM nas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b31ec-137">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

><span data-ttu-id="b31ec-138">! Observação O valor de tempo de vida padrão é `short` , mas recomendamos o uso `long` de suplementos do Excel. Se você definir seu tempo de execução como `short` neste exemplo, o suplemento do Excel será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá ser desligado após a execução do manipulador de faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b31ec-138">![NOTE] The default lifetime value is `short`, but we recommend using `long` in Excel add-ins. If you set your runtime to `short` in this example, your Excel add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="b31ec-139">Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas poderá ser desativado quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-139">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a><span data-ttu-id="b31ec-140">Vários painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="b31ec-140">Multiple task panes</span></span>

<span data-ttu-id="b31ec-141">Não projete seu suplemento para usar vários painéis de tarefas se você estiver planejando usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-141">Don't design your add-in to use multiple task panes if you are planning to use a shared runtime.</span></span> <span data-ttu-id="b31ec-142">Um tempo de execução compartilhado só dá suporte ao uso de um painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b31ec-142">A shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="b31ec-143">Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.</span><span class="sxs-lookup"><span data-stu-id="b31ec-143">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b31ec-144">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b31ec-144">Next steps</span></span>

- <span data-ttu-id="b31ec-145">Leia o artigo [Chamar APIs do Excel de uma função personalizada](call-excel-apis-from-custom-function.md) para obter detalhes sobre o uso das APIs JavaScript do Excel e funções personalizadas do Excel em um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b31ec-145">Read the [Call Excel APIs from a custom function](call-excel-apis-from-custom-function.md) article for details on using the Excel JavaScript APIs and custom Excel functions in a shared runtime.</span></span>
- <span data-ttu-id="b31ec-146">Explore o exemplo de padrões e práticas [Gerenciar a interface do usuário da faixa de opções e do painel de tarefas e executar o código no documento aberto](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) para ver um exemplo maior do tempo de execução compartilhado JavaScript em ação.</span><span class="sxs-lookup"><span data-stu-id="b31ec-146">Explore the patterns-and-practices sample [Manage ribbon and task pane UI, and run code on doc open](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/excel-shared-runtime-scenario) to see a larger example of the shared JavaScript runtime in action.</span></span>

## <a name="see-also"></a><span data-ttu-id="b31ec-147">Confira também</span><span class="sxs-lookup"><span data-stu-id="b31ec-147">See also</span></span>

- [<span data-ttu-id="b31ec-148">Visão geral: executar o código do seu suplemento em um tempo de execução JavaScript compartilhado</span><span class="sxs-lookup"><span data-stu-id="b31ec-148">Overview: Run your add-in code in a shared JavaScript runtime</span></span>](custom-functions-shared-overview.md)
