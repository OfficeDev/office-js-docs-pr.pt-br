---
ms.date: 03/13/2019
description: Depurar suas funções personalizadas no Excel.
title: Depuração de funções personalizadas (visualização)
localization_priority: Normal
ms.openlocfilehash: 66b55855fdbdc3b3cfc7a316cb8fd7e06f073213
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478948"
---
# <a name="custom-functions-debugging-preview"></a><span data-ttu-id="cb8eb-103">Depuração de funções personalizadas (visualização)</span><span class="sxs-lookup"><span data-stu-id="cb8eb-103">Custom functions debugging (preview)</span></span>

<span data-ttu-id="cb8eb-104">A depuração de funções personalizadas pode ser realizada por vários meios, dependendo de qual plataforma você está usando.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="cb8eb-105">No Windows:</span><span class="sxs-lookup"><span data-stu-id="cb8eb-105">On Windows:</span></span>
- [<span data-ttu-id="cb8eb-106">Depurador de área de trabalho do Excel e Visual Studio (VS Code)</span><span class="sxs-lookup"><span data-stu-id="cb8eb-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="cb8eb-107">O Excel online e o depurador de código VS</span><span class="sxs-lookup"><span data-stu-id="cb8eb-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="cb8eb-108">Excel online e ferramentas de navegador</span><span class="sxs-lookup"><span data-stu-id="cb8eb-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="cb8eb-109">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="cb8eb-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="cb8eb-110">No Mac:</span><span class="sxs-lookup"><span data-stu-id="cb8eb-110">On Mac:</span></span>
- [<span data-ttu-id="cb8eb-111">Excel online e ferramentas de navegador</span><span class="sxs-lookup"><span data-stu-id="cb8eb-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="cb8eb-112">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="cb8eb-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> <span data-ttu-id="cb8eb-113">Para simplificar, este artigo mostra a depuração no contexto de uso do Visual Studio Code para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="cb8eb-114">Se você estiver usando um editor ou uma ferramenta de linha de comando diferente, consulte as [instruções de linha de comando](#Use-the-command-line-tools-to-debug) no final deste artigo.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-114">If you are using a different editor or command line tool, see the [command line instructions](#Use-the-command-line-tools-to-debug) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="cb8eb-115">Requisitos</span><span class="sxs-lookup"><span data-stu-id="cb8eb-115">Requirements</span></span>

<span data-ttu-id="cb8eb-116">Antes de começar a depurar, você deve criar um projeto de suplemento de funções personalizadas usando o gerador de Yo Office e garantiu que você tenha certificados autoassinados confiáveis para o seu projeto.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-116">Before starting to debug, you should create a custom functions add-in project using the Yo Office generator and ensured that you have trusted self-signed certificates for your project.</span></span> <span data-ttu-id="cb8eb-117">Para obter instruções sobre como criar um projeto, consulte o [tutorial funções personalizadas](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-117">For instructions to create a project, see the [custom functions tutorial](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions).</span></span> <span data-ttu-id="cb8eb-118">Para obter instruções sobre como confiar em certificados, consulte [adicionando certificados autoassinados como certificados raiz confiáveis](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-118">For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="cb8eb-119">Usar o depurador de código VS para a área de trabalho do Excel</span><span class="sxs-lookup"><span data-stu-id="cb8eb-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="cb8eb-120">Você pode usar o VS Code para depurar funções personalizadas no Office Excel na área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-120">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="cb8eb-121">A depuração de área de trabalho do Mac não está disponível, mas pode ser obtida [usando as ferramentas de navegador para depurar o Excel online](#debug-in-excel-online-by-using-the-browser-developer-tools).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#debug-in-excel-online-by-using-the-browser-developer-tools).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="cb8eb-122">Executar seu suplemento de VS Code</span><span class="sxs-lookup"><span data-stu-id="cb8eb-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="cb8eb-123">Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="cb8eb-124">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="cb8eb-125">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="cb8eb-126">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="cb8eb-127">Iniciar o depurador do VS Code</span><span class="sxs-lookup"><span data-stu-id="cb8eb-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="cb8eb-128">Escolha **Exibir _GT_ Debug** ou Enter **Ctrl + Shift + D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-128">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="cb8eb-129">Nas opções de depuração, escolha **área de trabalho do Excel**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-129">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="cb8eb-130">Selecione **F5** (ou escolha **debug-> iniciar depuração** no menu) para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-130">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="cb8eb-131">Uma nova pasta de trabalho do Excel será aberta com seu suplemento já suplementos foi feito e pronto para uso.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="cb8eb-132">Iniciar Depuração</span><span class="sxs-lookup"><span data-stu-id="cb8eb-132">Start debugging</span></span>

1. <span data-ttu-id="cb8eb-133">No VS Code, abra o arquivo de script do código-fonte (funções. js ou funções. TS).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-133">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="cb8eb-134">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="cb8eb-135">Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="cb8eb-136">Nesse ponto, a execução será interrompida na linha de código em que você definir o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="cb8eb-137">Agora você pode percorrer seu código, definir inspeções e usar quaisquer recursos de depuração de código VS necessários.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="cb8eb-138">Usar o depurador de código VS para o Excel online no Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="cb8eb-138">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="cb8eb-139">Você pode usar o VS Code para depurar funções personalizadas no Excel online no navegador Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-139">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="cb8eb-140">Para usar o VS Code com o Microsoft Edge, você deve instalar o depurador para a extensão do [Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="cb8eb-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="cb8eb-141">Executar seu suplemento de VS Code</span><span class="sxs-lookup"><span data-stu-id="cb8eb-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="cb8eb-142">Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="cb8eb-143">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="cb8eb-144">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="cb8eb-145">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="cb8eb-146">Iniciar o depurador do VS Code</span><span class="sxs-lookup"><span data-stu-id="cb8eb-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="cb8eb-147">Escolha **Exibir _GT_ Debug** ou Enter **Ctrl + Shift + D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-147">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="cb8eb-148">Nas opções de depuração, escolha **Office Online (borda)**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-148">From the Debug options, choose **Office Online (Edge)**.</span></span>
6. <span data-ttu-id="cb8eb-149">Abra o Excel online usando o navegador do Microsoft Edge, abra o Excel online e crie uma nova pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-149">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="cb8eb-150">Escolha **compartilhar** na faixa de opções e copie o link para a URL dessa nova pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="cb8eb-151">Selecione **F5** (ou escolha **depurar > iniciar depuração** no menu) para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-151">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="cb8eb-152">Um prompt será exibido, solicitando a URL do seu documento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="cb8eb-153">Cole na URL da sua pasta de trabalho e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="cb8eb-154">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="cb8eb-154">Sideload your add-in</span></span>   

1. <span data-ttu-id="cb8eb-155">Selecione a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-155">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="cb8eb-156">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-156">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="cb8eb-158">**Navegue** até o arquivo de manifesto do suplemento e selecione **carregar**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="cb8eb-160">Definir pontos de interrupção</span><span class="sxs-lookup"><span data-stu-id="cb8eb-160">Set breakpoints</span></span>
1. <span data-ttu-id="cb8eb-161">No VS Code, abra o arquivo de script do código-fonte (funções. js ou funções. TS).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-161">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="cb8eb-162">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="cb8eb-163">Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="cb8eb-164">Usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online</span><span class="sxs-lookup"><span data-stu-id="cb8eb-164">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="cb8eb-165">Você pode usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-165">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="cb8eb-166">As etapas a seguir funcionam para o Windows e o macOS.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="cb8eb-167">Executar seu suplemento do Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="cb8eb-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="cb8eb-168">Abra a pasta do projeto raiz de suas funções personalizadas no [Visual Studio Code (vs Code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="cb8eb-169">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="cb8eb-170">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="cb8eb-171">Escolha **terminal _GT_ executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="cb8eb-172">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="cb8eb-172">Sideload your add-in</span></span>   

1. <span data-ttu-id="cb8eb-173">Abra o [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-173">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="cb8eb-174">Abra uma nova pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="cb8eb-175">Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-175">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="cb8eb-176">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-176">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="cb8eb-178">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="cb8eb-180">Depois que você tiver suplementos foi feito para o documento, ele permanecerá suplementos foi feito cada vez que você abrir o documento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="cb8eb-181">Iniciar Depuração</span><span class="sxs-lookup"><span data-stu-id="cb8eb-181">Start debugging</span></span>

1. <span data-ttu-id="cb8eb-182">Abra as ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-182">Open developer tools in the browser.</span></span> <span data-ttu-id="cb8eb-183">Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="cb8eb-184">Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte usando **cmd + p** ou **Ctrl + p** (funções. js ou funções. TS).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (functions.js or functions.ts).</span></span>
3. <span data-ttu-id="cb8eb-185">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="cb8eb-186">Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="cb8eb-187">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="cb8eb-188">Usar as ferramentas de linha de comando para depurar</span><span class="sxs-lookup"><span data-stu-id="cb8eb-188">Use the command line tools to debug</span></span>

<span data-ttu-id="cb8eb-189">Se você não estiver usando o VS, poderá usar a linha de comando (como bash ou PowerShell) para executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="cb8eb-190">Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel online.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-190">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="cb8eb-191">Não é possível depurar a versão da área de trabalho do Excel usando a linha de comando.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="cb8eb-192">A partir da linha de `npm run watch` comando, execute para observar e recriar quando ocorrerem alterações de código.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="cb8eb-193">Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução da inspeção).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="cb8eb-194">Se você deseja iniciar o suplemento na versão da área de trabalho do Excel, execute o seguinte comando</span><span class="sxs-lookup"><span data-stu-id="cb8eb-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start desktop`
    
    <span data-ttu-id="cb8eb-195">Ou se preferir iniciar seu suplemento no Excel online, execute o seguinte comando</span><span class="sxs-lookup"><span data-stu-id="cb8eb-195">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start web`
    
    <span data-ttu-id="cb8eb-196">Para o Excel online, você também precisa Sideload seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-196">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="cb8eb-197">Siga as etapas em [Sideload seu suplemento](#Sideload-your-add-in) para Sideload o suplemento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-197">Follow the steps in [Sideload your add-in](#Sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="cb8eb-198">Em seguida, prossiga para a próxima seção para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="cb8eb-199">Abra as ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-199">Open developer tools in the browser.</span></span> <span data-ttu-id="cb8eb-200">Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="cb8eb-201">Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte (funções. js ou funções. TS).</span><span class="sxs-lookup"><span data-stu-id="cb8eb-201">In developer tools, open your source code script file (functions.js or functions.ts).</span></span> <span data-ttu-id="cb8eb-202">O código de suas funções personalizadas pode estar localizado próximo ao final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="cb8eb-203">No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="cb8eb-204">Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="cb8eb-205">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="cb8eb-206">Comandos para compilar e executar o suplemento</span><span class="sxs-lookup"><span data-stu-id="cb8eb-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="cb8eb-207">Há várias tarefas de compilação disponíveis:</span><span class="sxs-lookup"><span data-stu-id="cb8eb-207">There are several build tasks available:</span></span>
- `npm run watch`<span data-ttu-id="cb8eb-208">: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo</span><span class="sxs-lookup"><span data-stu-id="cb8eb-208">: builds for development and automatically rebuilds when a source file is saved</span></span>
- `npm run build-dev`<span data-ttu-id="cb8eb-209">: cria para desenvolvimento uma vez</span><span class="sxs-lookup"><span data-stu-id="cb8eb-209">: builds for development once</span></span>
- `npm run build`<span data-ttu-id="cb8eb-210">: compilações para produção</span><span class="sxs-lookup"><span data-stu-id="cb8eb-210">: builds for production</span></span>
- `npm run dev-server`<span data-ttu-id="cb8eb-211">: executa o servidor Web usado para desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="cb8eb-211">: runs the web server used for development</span></span>

<span data-ttu-id="cb8eb-212">Você pode usar as seguintes tarefas para iniciar a depuração no desktop ou online.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- `npm run start desktop`<span data-ttu-id="cb8eb-213">: Inicia o Excel na área de trabalho e sideloads seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-213">: Starts Excel on desktop and sideloads your add-in.</span></span>
- `npm run start web`<span data-ttu-id="cb8eb-214">: Inicia o Excel online e o sideloads do suplemento.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-214">: Starts Excel Online and sideloads your add-in.</span></span>
- `npm run stop`<span data-ttu-id="cb8eb-215">: Interrompe o Excel e a depuração.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-215">: Stops Excel and debugging.</span></span>

## <a name="see-also"></a><span data-ttu-id="cb8eb-216">Confira também</span><span class="sxs-lookup"><span data-stu-id="cb8eb-216">See also</span></span>

* [<span data-ttu-id="cb8eb-217">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cb8eb-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="cb8eb-218">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="cb8eb-218">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cb8eb-219">Práticas recomendadas de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="cb8eb-219">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cb8eb-220">Log de alteração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="cb8eb-220">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="cb8eb-221">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="cb8eb-221">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
