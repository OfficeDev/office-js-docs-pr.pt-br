---
ms.date: 09/06/2019
description: Desenvolvimento de funções personalizadas no guia de início rápido do Excel.
title: Início rápido de funções personalizadas
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b3344b19ec29b24407c83bb043dcbaa613d5e8ec
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782230"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="e6f9f-103">Introdução ao desenvolvimento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="e6f9f-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="e6f9f-104">Com funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel, definindo-as em JavaScript ou typescript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="e6f9f-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa `SUM()`no Excel, como.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e6f9f-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="e6f9f-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="e6f9f-107">Excel no Windows (versão 1904 ou posterior, conectada à assinatura do Office 365) ou Excel na Web</span><span class="sxs-lookup"><span data-stu-id="e6f9f-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="e6f9f-108">As funções personalizadas do Excel têm suporte no Office no Mac (conectado à assinatura do Office 365) e uma atualização para este tutorial está em breve.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="e6f9f-109">As funções personalizadas do Excel não são suportadas no Office 2019 (compra única).</span><span class="sxs-lookup"><span data-stu-id="e6f9f-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="e6f9f-110">Criar seu primeiro projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="e6f9f-110">Build your first custom functions project</span></span>

<span data-ttu-id="e6f9f-111">Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="e6f9f-112">Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="e6f9f-113">Em uma pasta de sua preferência, execute o comando a seguir e responda aos prompts da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-113">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="e6f9f-114">**Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="e6f9f-114">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="e6f9f-115">**Escolha o tipo de script:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="e6f9f-115">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="e6f9f-116">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="e6f9f-116">**What do you want to name your add-in?**</span></span> `starcount`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/starcountPrompt.png)

    <span data-ttu-id="e6f9f-118">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-118">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="e6f9f-119">O gerador Yeoman fornecerá algumas instruções na linha de comando sobre o que fazer com o projeto, mas ignorará e continuarão seguindo as instruções.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-119">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="e6f9f-120">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="e6f9f-121">Compile o projeto.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-121">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="e6f9f-122">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e6f9f-123">Se você for solicitado a instalar um certificado após executar `npm run build`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-123">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="e6f9f-124">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-124">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="e6f9f-125">Você pode experimentar o suplemento função personalizada no Excel na Web ou no Windows.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-125">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="e6f9f-126">Você pode ser solicitado a abrir o painel de tarefas do suplemento, embora isso seja opcional.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-126">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="e6f9f-127">Você ainda pode executar suas funções personalizadas sem abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-127">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="e6f9f-128">Excel no Windows</span><span class="sxs-lookup"><span data-stu-id="e6f9f-128">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="e6f9f-129">Para testar seu suplemento no Excel no Windows, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-129">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="e6f9f-130">Quando você executar este comando, o servidor Web local será iniciado e o Excel será aberto com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="e6f9f-131">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="e6f9f-131">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="e6f9f-132">Para testar seu suplemento no Excel na Web, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-132">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="e6f9f-133">Quando você executa este comando, o servidor Web local iniciará.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="e6f9f-134">Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel em um navegador.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-134">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="e6f9f-135">Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="e6f9f-136">No Excel, escolha a guia **Inserir** e, em seguida, escolha **suplementos**.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-136">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Inserir faixa de opções no Excel na Web com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="e6f9f-138">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="e6f9f-139">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="e6f9f-140">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="e6f9f-141">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="e6f9f-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="e6f9f-142">O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas, definidas no arquivo **./src/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="e6f9f-142">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="e6f9f-143">O arquivo **./manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-143">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="e6f9f-144">Na sua pasta de trabalho do Excel, `ADD` Experimente a função personalizada realizando as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="e6f9f-144">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="e6f9f-145">Selecione uma célula e digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-145">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="e6f9f-146">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-146">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="e6f9f-147">Execute a `CONTOSO.ADD` função, usando números `10` e `200` como parâmetros de entrada, digitando o `=CONTOSO.ADD(10,200)` valor na célula e pressionando ENTER.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-147">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="e6f9f-148">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-148">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="e6f9f-149">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-149">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e6f9f-150">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e6f9f-150">Next steps</span></span>

<span data-ttu-id="e6f9f-151">Parabéns, você criou com êxito uma função personalizada em um suplemento do Excel!</span><span class="sxs-lookup"><span data-stu-id="e6f9f-151">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="e6f9f-152">Em seguida, crie um suplemento mais complexo com recurso de dados de streaming.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-152">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="e6f9f-153">O link a seguir o orienta pelas próximas etapas do tutorial do suplemento do Excel com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e6f9f-153">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e6f9f-154">Tutorial de suplemento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="e6f9f-154">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="e6f9f-155">Confira também</span><span class="sxs-lookup"><span data-stu-id="e6f9f-155">See also</span></span>

* [<span data-ttu-id="e6f9f-156">Visão geral das funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="e6f9f-156">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="e6f9f-157">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="e6f9f-157">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="e6f9f-158">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="e6f9f-158">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)