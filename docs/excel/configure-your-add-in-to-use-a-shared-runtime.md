---
ms.date: 02/20/2020
title: Configure o suplemento do Excel para compartilhar o tempo de execução do navegador (visualização)
ms.prod: excel
description: Configure o suplemento do Excel para compartilhar o tempo de execução do navegador e executar a faixa de opções, o painel de tarefas e o código de função personalizado no mesmo tempo de execução.
localization_priority: Priority
ms.openlocfilehash: 7fad97039d5a463978ca9a2d450af9bb698b49b5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719326"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a><span data-ttu-id="b758e-103">Configure seu suplemento do Excel para usar um tempo de execução JavaScript compartilhado (visualização)</span><span class="sxs-lookup"><span data-stu-id="b758e-103">Configure your Excel add-in to use a shared JavaScript runtime (preview)</span></span>

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="b758e-104">Ao executar o Excel no Windows ou Mac, o suplemento executará o código para botões da faixa de opções, funções personalizadas e o painel de tarefas em ambientes de tempo de execução JavaScript separados.</span><span class="sxs-lookup"><span data-stu-id="b758e-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="b758e-105">Isso cria limitações, como não poder compartilhar facilmente dados globais e não poder acessar todas as funcionalidades do CORS a partir de uma função customizada.</span><span class="sxs-lookup"><span data-stu-id="b758e-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="b758e-106">No entanto, você pode configurar o suplemento do Excel para compartilhar código em um tempo de execução JavaScript compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b758e-106">However, you can configure your Excel add-in to share code in a shared JavaScript runtime.</span></span> <span data-ttu-id="b758e-107">Isso permite uma melhor coordenação entre seu suplemento e acesso ao DOM e CORS de todas as partes do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="b758e-107">This enables better coordination across your add-in and access to the DOM and CORS from all parts of your add-in.</span></span> <span data-ttu-id="b758e-108">Também permite executar o código quando o documento é aberto ou executar o código enquanto o painel de tarefas está fechado.</span><span class="sxs-lookup"><span data-stu-id="b758e-108">It also enables you to run code when the document opens, or to run code while the task pane is closed.</span></span> <span data-ttu-id="b758e-109">Para configurar seu suplemento para usar um tempo de execução compartilhado, siga as instruções neste artigo.</span><span class="sxs-lookup"><span data-stu-id="b758e-109">To configure your add-in to use a shared runtime, follow the instructions in this article.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="b758e-110">Criar o projeto do suplemento</span><span class="sxs-lookup"><span data-stu-id="b758e-110">Create the add-in project</span></span>

<span data-ttu-id="b758e-111">Se você estiver iniciando um novo projeto, siga estas etapas para usar o gerador Yeoman para criar um projeto de suplemento do Excel.</span><span class="sxs-lookup"><span data-stu-id="b758e-111">If you are starting a new project, follow these steps to use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="b758e-112">Execute o comando a seguir e responda às solicitações com as seguintes respostas:</span><span class="sxs-lookup"><span data-stu-id="b758e-112">Run the following command and then answer the prompts with the following answers:</span></span>

```command line
yo office
```

- <span data-ttu-id="b758e-113">Escolha um tipo de projeto: **Projeto de suplemento de funções personalizadas do Excel**</span><span class="sxs-lookup"><span data-stu-id="b758e-113">Choose a project type: **Excel Custom Functions Add-in project**</span></span>
- <span data-ttu-id="b758e-114">Escolha um tipo de script: **JavaScript**</span><span class="sxs-lookup"><span data-stu-id="b758e-114">Choose a script type: **JavaScript**</span></span>
- <span data-ttu-id="b758e-115">Qual será o nome do seu suplemento? **Meu suplemento do Office**</span><span class="sxs-lookup"><span data-stu-id="b758e-115">What do you want to name your add-in? **My Office Add-in**</span></span>

![Captura de tela das solicitações de resposta do seu Office para criar o projeto de suplemento.](../images/yo-office-excel-project.png)

<span data-ttu-id="b758e-117">Depois que você concluir o assistente, o gerador criará o projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="b758e-117">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b758e-118">Configurar o manifesto</span><span class="sxs-lookup"><span data-stu-id="b758e-118">Configure the manifest</span></span>

<span data-ttu-id="b758e-119">Siga estas etapas para um projeto novo ou existente para configurá-lo para usar um tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b758e-119">Follow these steps for a new or existing project to configure it to use a shared runtime.</span></span>

1. <span data-ttu-id="b758e-120">Inicie o código do Visual Studio e abra o projeto **Meu suplemento do Office**.</span><span class="sxs-lookup"><span data-stu-id="b758e-120">Start Visual Studio Code and open the **My Office Add-in** project.</span></span>
2. <span data-ttu-id="b758e-121">Abra o arquivo **manifest.xml**.</span><span class="sxs-lookup"><span data-stu-id="b758e-121">Open the **manifest.xml** file.</span></span>
3. <span data-ttu-id="b758e-122">Localize a seção `<VersionOverrides>` e adicione a seguinte seção `<Runtimes>`.</span><span class="sxs-lookup"><span data-stu-id="b758e-122">Find the `<VersionOverrides>` section, and add the following `<Runtimes>` section.</span></span> <span data-ttu-id="b758e-123">O tempo de vida precisa ser **longo** para que as funções personalizadas ainda possam funcionar, mesmo quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="b758e-123">The lifetime needs to be **long** so that the custom functions can still work even when the task pane is closed.</span></span> <span data-ttu-id="b758e-124">O resid é `ContosoAddin.Url`, que faz referência a uma sequência na seção de recursos posteriormente.</span><span class="sxs-lookup"><span data-stu-id="b758e-124">The resid is `ContosoAddin.Url` which references a string in the resources section later.</span></span> <span data-ttu-id="b758e-125">Você pode usar qualquer valor de resid que desejar, mas deve corresponder ao resid dos outros elementos nos elementos do suplemento.</span><span class="sxs-lookup"><span data-stu-id="b758e-125">You can use any resid value you want, but it should match the resid of the other elements in your add-in elements.</span></span>

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. <span data-ttu-id="b758e-126">No elemento `<Page>`, altere o local de origem de **Functions.Page.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b758e-126">In the `<Page>` element, change the source location from **Functions.Page.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="b758e-127">Este resid corresponde ao elemento resid `<Runtime>`.</span><span class="sxs-lookup"><span data-stu-id="b758e-127">This resid matches the `<Runtime>` resid element.</span></span> <span data-ttu-id="b758e-128">Observe que, se você não tiver funções personalizadas, não terá uma entrada **Page** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b758e-128">Note that if you don't have custom functions, you will not have a **Page** entry and can skip this step.</span></span>

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. <span data-ttu-id="b758e-129">Na seção `<DesktopFormFactor>`, altere o **FunctionFile** de **Commands.Url** para usar **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b758e-129">In the `<DesktopFormFactor>` section, change the **FunctionFile** from **Commands.Url** to use **ContosoAddin.Url**.</span></span> <span data-ttu-id="b758e-130">Observe que, se você não possui comandos de ação, não terá uma entrada **FunctionFile** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b758e-130">Note that if you don't have action commands, you won't have a **FunctionFile** entry, and can skip this step.</span></span>

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. <span data-ttu-id="b758e-131">Na seção `<Action>`, altere o local de origem de **Taskpane.Url** para **ContosoAddin.Url**.</span><span class="sxs-lookup"><span data-stu-id="b758e-131">In the `<Action>` section, change the source location from **Taskpane.Url** to **ContosoAddin.Url**.</span></span> <span data-ttu-id="b758e-132">Observe que, se você não tiver um painel de tarefas, não terá uma ação **ShowTaskpane** e poderá pular esta etapa.</span><span class="sxs-lookup"><span data-stu-id="b758e-132">Note that if you don't have a task pane, you won't have a **ShowTaskpane** action, and can skip this step.</span></span>

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. <span data-ttu-id="b758e-133">Adicione um novo **ID de URL** para **ContosoAddin.Url** que aponte para **taskpane.html**.</span><span class="sxs-lookup"><span data-stu-id="b758e-133">Add a new **Url id** for **ContosoAddin.Url** that points to **taskpane.html**.</span></span>

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. <span data-ttu-id="b758e-134">Salve suas alterações e recompile o projeto.</span><span class="sxs-lookup"><span data-stu-id="b758e-134">Save your changes and rebuild the project.</span></span>

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a><span data-ttu-id="b758e-135">Duração do tempo de execução</span><span class="sxs-lookup"><span data-stu-id="b758e-135">Runtime lifetime</span></span>

<span data-ttu-id="b758e-136">Ao adicionar o elemento `Runtime`, você também especifica uma vida útil com um valor de `long` ou `short`.</span><span class="sxs-lookup"><span data-stu-id="b758e-136">When you add the `Runtime` element, you also specify a lifetime with a value of `long` or `short`.</span></span> <span data-ttu-id="b758e-137">Defina esse valor como `long` para aproveitar os recursos, como iniciar o suplemento quando o documento for aberto, continuar executando o código após o fechamento do painel de tarefas ou usar o CORS e o DOM nas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="b758e-137">Set this value to `long` to take advantage of features such as starting your add-in when the document opens, continuing to run code after the task pane is closed, or using CORS and DOM from custom functions.</span></span>

<span data-ttu-id="b758e-138">Se você definir esse valor como `short`, o suplemento se comportará de maneira semelhante ao comportamento padrão.</span><span class="sxs-lookup"><span data-stu-id="b758e-138">If you set this value to `short`, the add-in will behave similar to the default behavior.</span></span> <span data-ttu-id="b758e-139">O suplemento será iniciado quando um dos botões da faixa de opções for pressionado, mas poderá desligar após a execução do manipulador da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="b758e-139">The add-in will start when one of your ribbon buttons is pressed, but it may shut down after your ribbon handler is done running.</span></span> <span data-ttu-id="b758e-140">Da mesma forma, o suplemento será iniciado quando o painel de tarefas for aberto, mas poderá ser desativado quando o painel de tarefas estiver fechado.</span><span class="sxs-lookup"><span data-stu-id="b758e-140">Similarly your add-in will start when the task pane is opened, but it may shut down when the task pane is closed.</span></span>

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a><span data-ttu-id="b758e-141">Vários painéis de tarefas</span><span class="sxs-lookup"><span data-stu-id="b758e-141">Multiple task panes</span></span>

<span data-ttu-id="b758e-142">Não projete seu suplemento para usar vários painéis de tarefas se você planeja usar o tempo de execução compartilhado.</span><span class="sxs-lookup"><span data-stu-id="b758e-142">Don't design your add-in to use multiple task panes if you are planning to use shared runtime.</span></span> <span data-ttu-id="b758e-143">O tempo de execução compartilhado tem suporte para o uso de apenas um único painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="b758e-143">Shared runtime only supports the use of one task pane.</span></span> <span data-ttu-id="b758e-144">Observe que qualquer painel de tarefas sem um `<TaskpaneID>` é considerado um painel de tarefas diferente.</span><span class="sxs-lookup"><span data-stu-id="b758e-144">Note that any task pane without a `<TaskpaneID>` is considered a different task pane.</span></span>

## <a name="next-steps"></a><span data-ttu-id="b758e-145">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="b758e-145">Next steps</span></span>

<span data-ttu-id="b758e-146">Agora experimente alguns recursos do tempo de execução compartilhado, observando os seguintes artigos.</span><span class="sxs-lookup"><span data-stu-id="b758e-146">Now try out some features of the shared runtime by taking a look at the following articles.</span></span>

- [<span data-ttu-id="b758e-147">Chamar APIs do Excel a partir de uma função personalizada</span><span class="sxs-lookup"><span data-stu-id="b758e-147">Call Excel APIs from a custom function</span></span>](call-excel-apis-from-custom-function.md)

## <a name="see-also"></a><span data-ttu-id="b758e-148">Confira também</span><span class="sxs-lookup"><span data-stu-id="b758e-148">See also</span></span>

- [<span data-ttu-id="b758e-149">Visão geral: Execute seu código de suplemento em um tempo de execução JavaScript compartilhado (visualização)</span><span class="sxs-lookup"><span data-stu-id="b758e-149">Overview: Run your add-in code in a shared JavaScript runtime (preview)</span></span>](custom-functions-shared-overview.md)
