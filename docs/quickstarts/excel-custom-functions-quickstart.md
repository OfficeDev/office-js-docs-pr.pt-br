---
ms.date: 03/06/2019
description: Desenvolvimento de funções personalizadas no guia de início rápido do Excel.
title: Início rápido de funções personalizadas (visualização)
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 3ea7ec4c2089aaa4e9f193a45e7c4a31c691f213
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517069"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="643f0-103">Introdução ao desenvolvimento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="643f0-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="643f0-104">Com funções personalizadas, os desenvolvedores agora podem adicionar novas funções ao Excel, definindo-as em JavaScript ou typescript como parte de um suplemento.</span><span class="sxs-lookup"><span data-stu-id="643f0-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="643f0-105">Os usuários do Excel podem acessar funções personalizadas da mesma forma que fariam com qualquer função nativa `SUM()`no Excel, como.</span><span class="sxs-lookup"><span data-stu-id="643f0-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="643f0-106">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="643f0-106">Prerequisites</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="643f0-107">Você precisará das seguintes ferramentas e recursos relacionados para começar a criar funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="643f0-107">You'll need the following tools and related resources to begin creating custom functions.</span></span>

- <span data-ttu-id="643f0-108">[Node](https://nodejs.org/en/) (versão 8.0.0 ou posterior)</span><span class="sxs-lookup"><span data-stu-id="643f0-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="643f0-109">[Git Bash](https://git-scm.com/downloads) (ou outro cliente Git)</span><span class="sxs-lookup"><span data-stu-id="643f0-109">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="643f0-110">A versão mais recente do [Yeoman](https://yeoman.io/) e do [Yeoman gerador de suplementos do Office](https://www.npmjs.com/package/generator-office). Para instalar essas ferramentas globalmente, execute o seguinte comando por meio do prompt de comando:</span><span class="sxs-lookup"><span data-stu-id="643f0-110">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="643f0-111">Mesmo que você já tenha instalado o gerador Yeoman, recomendamos atualizar seu pacote para a versão mais recente do NPM.</span><span class="sxs-lookup"><span data-stu-id="643f0-111">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="643f0-112">Criar seu primeiro projeto de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="643f0-112">Build your first custom functions project</span></span>

<span data-ttu-id="643f0-113">Para começar, você usará o gerador Yeoman para criar projeto com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="643f0-113">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="643f0-114">Isso configurará seu projeto com a estrutura de pastas, arquivos de origem e dependências corretos para começar a codificar suas funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="643f0-114">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="643f0-115">Execute o comando a seguir e responda aos prompts da seguinte forma.</span><span class="sxs-lookup"><span data-stu-id="643f0-115">Run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="643f0-116">Escolha o tipo de projeto:`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="643f0-116">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    - <span data-ttu-id="643f0-117">Escolha um tipo de script: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="643f0-117">Choose a script type: `JavaScript`</span></span>

    - <span data-ttu-id="643f0-118">Qual será o nome do suplemento?</span><span class="sxs-lookup"><span data-stu-id="643f0-118">What do you want to name your add-in?</span></span> `stock-ticker`

    ![O gerador Yeoman para suplementos do Office solicita funções personalizadas](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="643f0-120">O gerador Yeoman criará os arquivos do projeto e instalará os componentes Node de suporte.</span><span class="sxs-lookup"><span data-stu-id="643f0-120">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="643f0-121">Navegue até a pasta do projeto que você acabou de criar.</span><span class="sxs-lookup"><span data-stu-id="643f0-121">Navigate to the project folder you just created.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="643f0-122">Confie no certificado autoassinado necessário para executar este projeto.</span><span class="sxs-lookup"><span data-stu-id="643f0-122">Trust the self-signed certificate you need to run this project.</span></span> <span data-ttu-id="643f0-123">Para obter instruções detalhadas para Windows ou Mac, confira [Adicionando Certificados Autoassinados como Certificado Raiz Confiável](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="643f0-123">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="643f0-124">Crie um projeto.</span><span class="sxs-lookup"><span data-stu-id="643f0-124">Build the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

5. <span data-ttu-id="643f0-125">Inicie o servidor local da web, que é executado no Node.</span><span class="sxs-lookup"><span data-stu-id="643f0-125">Start the local web server, which runs in Node.js.</span></span>

    - <span data-ttu-id="643f0-126">Se você usar o Excel para Windows para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local, inicie o Excel e Sideload o suplemento:</span><span class="sxs-lookup"><span data-stu-id="643f0-126">If you use Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```command&nbsp;line
         npm run start
        ```
        <span data-ttu-id="643f0-127">Depois de executar esse comando, o prompt de comando mostrará detalhes sobre como iniciar o servidor Web.</span><span class="sxs-lookup"><span data-stu-id="643f0-127">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="643f0-128">O Excel começará com seu suplemento carregado.</span><span class="sxs-lookup"><span data-stu-id="643f0-128">Excel will start with your add-in loaded.</span></span> <span data-ttu-id="643f0-129">Se o suplemento não carregar, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="643f0-129">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    - <span data-ttu-id="643f0-130">Se você usar o Excel online para testar suas funções personalizadas, execute o seguinte comando para iniciar o servidor Web local:</span><span class="sxs-lookup"><span data-stu-id="643f0-130">If you use Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```command&nbsp;line
        npm run start-web
        ```

         <span data-ttu-id="643f0-131">Depois de executar esse comando, o prompt de comando mostrará detalhes sobre como iniciar o servidor Web.</span><span class="sxs-lookup"><span data-stu-id="643f0-131">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="643f0-132">Para usar suas funções, abra uma nova pasta de trabalho no Excel online.</span><span class="sxs-lookup"><span data-stu-id="643f0-132">To use your functions, open a new workbook in Excel Online.</span></span> <span data-ttu-id="643f0-133">Nesta pasta de trabalho, você precisará carregar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="643f0-133">In this workbook, you'll need to load your add-in.</span></span> 

        <span data-ttu-id="643f0-134">Para fazer isso, selecione a guia **Inserir** na faixa de opções e selecione **obter suplementos**. Na nova janela resultante, verifique se você está na guia **meus suplementos** . Em seguida, selecione **gerenciar meus suplementos _GT_ carregar meu suplemento**.</span><span class="sxs-lookup"><span data-stu-id="643f0-134">To do this, select the **Insert** tab on the ribbon and select **Get Add-ins**. In the resulting new window, ensure you are on the **My Add-ins** tab. Next, select **Manage My Add-ins > Upload My Add-in**.</span></span> <span data-ttu-id="643f0-135">Procure o arquivo de manifesto e carregue-o.</span><span class="sxs-lookup"><span data-stu-id="643f0-135">Browse for your manifest file and upload it.</span></span> <span data-ttu-id="643f0-136">Se o suplemento não for carregado, verifique se você concluiu a etapa 3 corretamente.</span><span class="sxs-lookup"><span data-stu-id="643f0-136">If your add-in does not load, check you've completed step 3 correctly.</span></span>

## <a name="try-out-the-prebuilt-custom-functions"></a><span data-ttu-id="643f0-137">Experimentar as funções personalizadas predefinidas</span><span class="sxs-lookup"><span data-stu-id="643f0-137">Try out the prebuilt custom functions</span></span>

<span data-ttu-id="643f0-138">O projeto de funções personalizadas criado usando o gerador Yeoman contém algumas funções personalizadas predefinidas definidas no arquivo **src/customfunction.js**.</span><span class="sxs-lookup"><span data-stu-id="643f0-138">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="643f0-139">O arquivo **manifest. XML** na pasta raiz do projeto especifica que todas as funções personalizadas pertencem ao `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="643f0-139">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="643f0-140">Na sua pasta de trabalho do Excel, `ADD` Experimente a função personalizada realizando as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="643f0-140">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="643f0-141">Selecione uma célula e digite `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="643f0-141">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="643f0-142">Observe que o menu de preenchimento automático mostra a lista de todas as funções na `CONTOSO` namespace.</span><span class="sxs-lookup"><span data-stu-id="643f0-142">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="643f0-143">Execute a `CONTOSO.ADD` função, usando números `10` e `200` como parâmetros de entrada, digitando o `=CONTOSO.ADD(10,200)` valor na célula e pressionando ENTER.</span><span class="sxs-lookup"><span data-stu-id="643f0-143">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="643f0-144">O `ADD` função personalizada calcula a soma de dois números que você especificar como os parâmetros de entrada.</span><span class="sxs-lookup"><span data-stu-id="643f0-144">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="643f0-145">Digitando `=CONTOSO.ADD(10,200)` deve obter o resultado **210** na célula, depois pressionar enter.</span><span class="sxs-lookup"><span data-stu-id="643f0-145">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="643f0-146">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="643f0-146">Next steps</span></span>

<span data-ttu-id="643f0-147">Parabéns, você criou com êxito uma função personalizada em um suplemento do Excel!</span><span class="sxs-lookup"><span data-stu-id="643f0-147">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="643f0-148">Em seguida, crie um suplemento mais complexo com recurso de dados de streaming.</span><span class="sxs-lookup"><span data-stu-id="643f0-148">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="643f0-149">O link a seguir o orienta pelas próximas etapas do tutorial do suplemento do Excel com funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="643f0-149">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="643f0-150">Tutorial de suplemento de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="643f0-150">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="643f0-151">Confira também</span><span class="sxs-lookup"><span data-stu-id="643f0-151">See also</span></span>

* [<span data-ttu-id="643f0-152">Visão geral de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="643f0-152">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="643f0-153">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="643f0-153">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="643f0-154">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="643f0-154">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* <span data-ttu-id="643f0-155">[Práticas recomendadas de funções personalizadas](../excel/custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="643f0-155">[Custom functions best practices](../excel/custom-functions-best-practices.md)</span></span>
