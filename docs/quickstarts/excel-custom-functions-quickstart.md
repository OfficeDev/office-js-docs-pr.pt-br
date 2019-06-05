---
ms.date: 05/30/2019
description: Desenvolvimento de funções personalizadas no guia de início rápido do Excel.
title: Início rápido de funções personalizadas
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4bf0d6a5bf020ee4196ce89d763fa994b3fd489c
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706039"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="0884f-103">Introdução ao desenvolvimento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0884f-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="0884f-104">Com funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel, definindo-as em JavaScript ou typescript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="0884f-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="0884f-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa `SUM()`no Excel, como.</span><span class="sxs-lookup"><span data-stu-id="0884f-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0884f-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="0884f-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="0884f-107">Excel no Windows (versão 1810 ou posterior) ou Excel online</span><span class="sxs-lookup"><span data-stu-id="0884f-107">Excel on Windows (version 1810 or later) or Excel Online</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="0884f-108">Criar seu primeiro projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0884f-108">Build your first custom functions project</span></span>

<span data-ttu-id="0884f-109">Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0884f-109">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="0884f-110">Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0884f-110">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="0884f-111">Em uma pasta de sua preferência, execute o comando a seguir e responda aos prompts da seguinte maneira.</span><span class="sxs-lookup"><span data-stu-id="0884f-111">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="0884f-112">**Escolha o tipo de projeto:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="0884f-112">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="0884f-113">**Escolha o tipo de script:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0884f-113">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="0884f-114">**Qual será o nome do suplemento?**</span><span class="sxs-lookup"><span data-stu-id="0884f-114">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/UpdatedYoOfficePrompt.png)

    <span data-ttu-id="0884f-116">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="0884f-116">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="0884f-117">O gerador Yeoman fornecerá algumas instruções na linha de comando sobre o que fazer com o projeto, mas ignorará e continuarão seguindo as instruções.</span><span class="sxs-lookup"><span data-stu-id="0884f-117">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="0884f-118">Navegue até a pasta raiz do projeto.</span><span class="sxs-lookup"><span data-stu-id="0884f-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="0884f-119">Compile o projeto.</span><span class="sxs-lookup"><span data-stu-id="0884f-119">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="0884f-120">Os Suplementos do Office devem usar HTTPS, e não HTTP, mesmo durante o desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="0884f-120">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="0884f-121">Se você for solicitado a instalar um certificado após executar `npm run build`, aceite a solicitação para instalar o certificado que o gerador do Yeoman fornecer.</span><span class="sxs-lookup"><span data-stu-id="0884f-121">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="0884f-122">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="0884f-122">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="0884f-123">Você pode experimentar o suplemento função personalizada no Excel no Windows ou no Excel online.</span><span class="sxs-lookup"><span data-stu-id="0884f-123">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="0884f-124">Você pode ser solicitado a abrir o painel de tarefas do suplemento, embora isso seja opcional.</span><span class="sxs-lookup"><span data-stu-id="0884f-124">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="0884f-125">Você ainda pode executar suas funções personalizadas sem abrir o painel de tarefas do suplemento.</span><span class="sxs-lookup"><span data-stu-id="0884f-125">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0884f-126">Excel no Windows</span><span class="sxs-lookup"><span data-stu-id="0884f-126">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="0884f-127">Para testar seu suplemento no Excel no Windows, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="0884f-127">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="0884f-128">Quando você executar este comando, o servidor Web local será iniciado e o Excel será aberto com o seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="0884f-128">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="0884f-129">Excel Online</span><span class="sxs-lookup"><span data-stu-id="0884f-129">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="0884f-130">Para testar seu suplemento no Excel online, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="0884f-130">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="0884f-131">Quando você executa este comando, o servidor Web local iniciará.</span><span class="sxs-lookup"><span data-stu-id="0884f-131">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="0884f-132">Para usar seu suplemento de funções personalizadas, abra uma nova pasta de trabalho no Excel online.</span><span class="sxs-lookup"><span data-stu-id="0884f-132">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="0884f-133">Nesta pasta de trabalho, conclua as seguintes etapas para Sideload seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="0884f-133">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="0884f-134">No Excel Online, escolha a guia **inserir** pressione e, em seguida, escolha **suplementos**.</span><span class="sxs-lookup"><span data-stu-id="0884f-134">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Inserir faixa de opções no Excel online com o ícone meus suplementos realçado](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="0884f-136">Escolha **Gerenciar Meus suplementos** e selecione **Carregar o Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="0884f-136">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0884f-137">Escolha \*\*Procurar... \*\* e navegue até o diretório raiz do projeto criado pelo gerador Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0884f-137">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0884f-138">Selecione o arquivo **manifest. XML** e escolha **aberto**, escolha **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="0884f-138">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0884f-139">Experimente uma função personalizada predefinida</span><span class="sxs-lookup"><span data-stu-id="0884f-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0884f-140">O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas, definidas no arquivo **./src/Functions/functions.js** .</span><span class="sxs-lookup"><span data-stu-id="0884f-140">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="0884f-141">O arquivo **./manifest.xml** no diretório raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="0884f-141">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="0884f-142">Na sua pasta de trabalho do Excel, `ADD` Experimente a função personalizada realizando as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="0884f-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="0884f-143">Selecione uma célula e digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0884f-143">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="0884f-144">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="0884f-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0884f-145">Execute a `CONTOSO.ADD` função, usando números `10` e `200` como parâmetros de entrada, digitando o `=CONTOSO.ADD(10,200)` valor na célula e pressionando ENTER.</span><span class="sxs-lookup"><span data-stu-id="0884f-145">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="0884f-146">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="0884f-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="0884f-147">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="0884f-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0884f-148">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="0884f-148">Next steps</span></span>

<span data-ttu-id="0884f-149">Parabéns, você criou com êxito uma função personalizada em um suplemento do Excel!</span><span class="sxs-lookup"><span data-stu-id="0884f-149">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="0884f-150">Em seguida, crie um suplemento mais complexo com recurso de dados de streaming.</span><span class="sxs-lookup"><span data-stu-id="0884f-150">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="0884f-151">O link a seguir o orienta pelas próximas etapas do tutorial do suplemento do Excel com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="0884f-151">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0884f-152">Tutorial de suplemento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0884f-152">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="0884f-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="0884f-153">See also</span></span>

* [<span data-ttu-id="0884f-154">Visão geral das funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0884f-154">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="0884f-155">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="0884f-155">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="0884f-156">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="0884f-156">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* <span data-ttu-id="0884f-157">[Práticas recomendadas de funções personalizadas](../excel/custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="0884f-157">[Custom functions best practices](../excel/custom-functions-best-practices.md)</span></span>
