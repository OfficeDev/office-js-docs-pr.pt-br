---
ms.date: 06/17/2019
description: Depurar suas funções personalizadas no Excel.
title: Depuração de funções personalizadas
localization_priority: Normal
ms.openlocfilehash: 414944e66a6c55228ea009291be42218038fc6fa
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059864"
---
# <a name="custom-functions-debugging"></a><span data-ttu-id="e8bd4-103">Depuração de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="e8bd4-103">Custom functions debugging</span></span>

<span data-ttu-id="e8bd4-104">A depuração de funções personalizadas pode ser realizada por vários meios, dependendo de qual plataforma você está usando.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="e8bd4-105">No Windows:</span><span class="sxs-lookup"><span data-stu-id="e8bd4-105">On Windows:</span></span>
- [<span data-ttu-id="e8bd4-106">Depurador de área de trabalho do Excel e Visual Studio (VS Code)</span><span class="sxs-lookup"><span data-stu-id="e8bd4-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="e8bd4-107">O Excel online e o depurador de código VS</span><span class="sxs-lookup"><span data-stu-id="e8bd4-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="e8bd4-108">Excel online e ferramentas de navegador</span><span class="sxs-lookup"><span data-stu-id="e8bd4-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="e8bd4-109">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="e8bd4-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="e8bd4-110">No Mac:</span><span class="sxs-lookup"><span data-stu-id="e8bd4-110">On Mac:</span></span>
- [<span data-ttu-id="e8bd4-111">Excel online e ferramentas de navegador</span><span class="sxs-lookup"><span data-stu-id="e8bd4-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="e8bd4-112">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="e8bd4-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="e8bd4-113">Para simplificar, este artigo mostra a depuração no contexto de uso do Visual Studio Code para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="e8bd4-114">Se você estiver usando um editor ou uma ferramenta de linha de comando diferente, consulte as [instruções de linha de comando](#commands-for-building-and-running-your-add-in) no final deste artigo.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="e8bd4-115">Requisitos</span><span class="sxs-lookup"><span data-stu-id="e8bd4-115">Requirements</span></span>

<span data-ttu-id="e8bd4-116">Antes de começar a depurar, você deve usar o [gerador Yeoman para suplementos do Office](https://github.com/OfficeDev/generator-office) para criar um projeto de funções personalizadas.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-116">Before starting to debug, you should use the [Yeoman generator for Office add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="e8bd4-117">Para obter orientação sobre como criar um projeto de funções personalizadas, consulte o [tutorial de funções personalizadas](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="e8bd4-118">Usar o depurador de código VS para a área de trabalho do Excel</span><span class="sxs-lookup"><span data-stu-id="e8bd4-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="e8bd4-119">Você pode usar o VS Code para depurar funções personalizadas no Office Excel na área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-119">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="e8bd4-120">A depuração de área de trabalho do Mac não está disponível, mas pode ser obtida [usando as ferramentas de navegador e a linha de comando para depurar o Excel online](#use-the-command-line-tools-to-debug) ).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel Online](#use-the-command-line-tools-to-debug) ).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="e8bd4-121">Executar seu suplemento de VS Code</span><span class="sxs-lookup"><span data-stu-id="e8bd4-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="e8bd4-122">Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="e8bd4-123">Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="e8bd4-124">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="e8bd4-125">Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="e8bd4-126">Iniciar o depurador do VS Code</span><span class="sxs-lookup"><span data-stu-id="e8bd4-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="e8bd4-127">Escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-127">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="e8bd4-128">Nas opções de depuração, escolha **área de trabalho do Excel**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-128">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="e8bd4-129">Selecione **F5** (ou escolha **debug-> iniciar a depuração** no menu) para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-129">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="e8bd4-130">Uma nova pasta de trabalho do Excel será aberta com seu suplemento já suplementos foi feito e pronto para uso.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="e8bd4-131">Iniciar Depuração</span><span class="sxs-lookup"><span data-stu-id="e8bd4-131">Start debugging</span></span>

1. <span data-ttu-id="e8bd4-132">No VS Code, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="e8bd4-133">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="e8bd4-134">Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="e8bd4-135">Nesse ponto, a execução será interrompida na linha de código em que você definir o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="e8bd4-136">Agora você pode percorrer seu código, definir inspeções e usar quaisquer recursos de depuração de código VS necessários.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="e8bd4-137">Usar o depurador de código VS para o Excel online no Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="e8bd4-137">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="e8bd4-138">Você pode usar o VS Code para depurar funções personalizadas no Excel online no navegador Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-138">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="e8bd4-139">Para usar o VS Code com o Microsoft Edge, você deve instalar o depurador para a extensão do [Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="e8bd4-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="e8bd4-140">Executar seu suplemento de VS Code</span><span class="sxs-lookup"><span data-stu-id="e8bd4-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="e8bd4-141">Abra a pasta do projeto raiz de funções personalizadas no [vs Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="e8bd4-142">Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="e8bd4-143">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="e8bd4-144">Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="e8bd4-145">Iniciar o depurador do VS Code</span><span class="sxs-lookup"><span data-stu-id="e8bd4-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="e8bd4-146">Escolha **exibir > depurar** ou digite **Ctrl + Shift + D** para alternar para o modo de depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-146">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="e8bd4-147">Nas opções de depuração, escolha **Office Online (Microsoft Edge)**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-147">From the Debug options, choose **Office Online (Microsoft Edge)**.</span></span>
6. <span data-ttu-id="e8bd4-148">Abra o Excel online usando o navegador do Microsoft Edge, abra o Excel online e crie uma nova pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-148">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="e8bd4-149">Escolha **compartilhar** na faixa de opções e copie o link para a URL dessa nova pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="e8bd4-150">Selecione **F5** (ou escolha **debug > iniciar a depuração** no menu) para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-150">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="e8bd4-151">Um prompt será exibido, solicitando a URL do seu documento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="e8bd4-152">Cole na URL da sua pasta de trabalho e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="e8bd4-153">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="e8bd4-153">Sideload your add-in</span></span>   

1. <span data-ttu-id="e8bd4-154">Selecione a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-154">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="e8bd4-155">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-155">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="e8bd4-157">**Navegue** até o arquivo de manifesto do suplemento e selecione **carregar**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="e8bd4-159">Definir pontos de interrupção</span><span class="sxs-lookup"><span data-stu-id="e8bd4-159">Set breakpoints</span></span>
1. <span data-ttu-id="e8bd4-160">No VS Code, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="e8bd4-161">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="e8bd4-162">Na pasta de trabalho do Excel, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="e8bd4-163">Usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online</span><span class="sxs-lookup"><span data-stu-id="e8bd4-163">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="e8bd4-164">Você pode usar as ferramentas de desenvolvedor do navegador para depurar as funções personalizadas no Excel online.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-164">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="e8bd4-165">As etapas a seguir funcionam para o Windows e o macOS.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="e8bd4-166">Executar seu suplemento do Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="e8bd4-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="e8bd4-167">Abra a pasta do projeto raiz de suas funções personalizadas no [Visual Studio Code (vs Code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="e8bd4-168">Escolha **Terminal > executar tarefa** e digite ou selecione **Watch**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="e8bd4-169">Isso irá monitorar e recriar qualquer alteração de arquivo.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="e8bd4-170">Escolha **Terminal > executar tarefa** e digite ou selecione **servidor de desenvolvimento**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="e8bd4-171">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="e8bd4-171">Sideload your add-in</span></span>   

1. <span data-ttu-id="e8bd4-172">Abra o [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-172">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="e8bd4-173">Abra uma nova pasta de trabalho do Excel.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="e8bd4-174">Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-174">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="e8bd4-175">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-175">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="e8bd4-177">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="e8bd4-179">Depois que você tiver suplementos foi feito para o documento, ele permanecerá suplementos foi feito cada vez que você abrir o documento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="e8bd4-180">Iniciar Depuração</span><span class="sxs-lookup"><span data-stu-id="e8bd4-180">Start debugging</span></span>

1. <span data-ttu-id="e8bd4-181">Abra as ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-181">Open developer tools in the browser.</span></span> <span data-ttu-id="e8bd4-182">Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="e8bd4-183">Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte usando **cmd + p** ou **Ctrl + p** (**funções. js** ou **funções. TS**).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="e8bd4-184">[Defina um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="e8bd4-185">Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="e8bd4-186">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="e8bd4-187">Usar as ferramentas de linha de comando para depurar</span><span class="sxs-lookup"><span data-stu-id="e8bd4-187">Use the command line tools to debug</span></span>

<span data-ttu-id="e8bd4-188">Se você não estiver usando o VS, poderá usar a linha de comando (como bash ou PowerShell) para executar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="e8bd4-189">Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código no Excel online.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-189">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="e8bd4-190">Não é possível depurar a versão da área de trabalho do Excel usando a linha de comando.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="e8bd4-191">A partir da linha de `npm run watch` comando, execute para observar e recriar quando ocorrerem alterações de código.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="e8bd4-192">Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução da inspeção).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="e8bd4-193">Se você deseja iniciar o suplemento na versão da área de trabalho do Excel, execute o seguinte comando</span><span class="sxs-lookup"><span data-stu-id="e8bd4-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="e8bd4-194">Ou se preferir iniciar seu suplemento no Excel online, execute o seguinte comando</span><span class="sxs-lookup"><span data-stu-id="e8bd4-194">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="e8bd4-195">Para o Excel online, você também precisa Sideload seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-195">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="e8bd4-196">Siga as etapas em [Sideload seu suplemento](#sideload-your-add-in) para Sideload o suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="e8bd4-197">Em seguida, prossiga para a próxima seção para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="e8bd4-198">Abra as ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-198">Open developer tools in the browser.</span></span> <span data-ttu-id="e8bd4-199">Para o Chrome e a maioria dos navegadores, o F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="e8bd4-200">Em ferramentas de desenvolvedor, abra o arquivo de script do código-fonte (**funções. js** ou **funções. TS**).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="e8bd4-201">O código de suas funções personalizadas pode estar localizado próximo ao final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="e8bd4-202">No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="e8bd4-203">Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="e8bd4-204">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="e8bd4-205">Comandos para compilar e executar o suplemento</span><span class="sxs-lookup"><span data-stu-id="e8bd4-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="e8bd4-206">Há várias tarefas de compilação disponíveis:</span><span class="sxs-lookup"><span data-stu-id="e8bd4-206">There are several build tasks available:</span></span>
- <span data-ttu-id="e8bd4-207">`npm run watch`: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo</span><span class="sxs-lookup"><span data-stu-id="e8bd4-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="e8bd4-208">`npm run build-dev`: cria para desenvolvimento uma vez</span><span class="sxs-lookup"><span data-stu-id="e8bd4-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="e8bd4-209">`npm run build`: compilações para produção</span><span class="sxs-lookup"><span data-stu-id="e8bd4-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="e8bd4-210">`npm run dev-server`: executa o servidor Web usado para desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="e8bd4-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="e8bd4-211">Você pode usar as seguintes tarefas para iniciar a depuração no desktop ou online.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="e8bd4-212">`npm run start:desktop`: Inicia o Excel na área de trabalho e sideloads seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="e8bd4-213">`npm run start:web`: Inicia o Excel online e o sideloads do suplemento.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-213">`npm run start:web`: Starts Excel Online and sideloads your add-in.</span></span>
- <span data-ttu-id="e8bd4-214">`npm run stop`: Interrompe o Excel e a depuração.</span><span class="sxs-lookup"><span data-stu-id="e8bd4-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e8bd4-215">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="e8bd4-215">Next steps</span></span>
<span data-ttu-id="e8bd4-216">Saiba mais sobre as [práticas de autenticação em funções personalizadas](custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-216">Learn about [authentication practices in custom functions](custom-functions-authentication.md).</span></span> <span data-ttu-id="e8bd4-217">Ou, revise a [arquitetura exclusiva da função personalizada](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-217">Or, review [custom function's unique architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e8bd4-218">Confira também</span><span class="sxs-lookup"><span data-stu-id="e8bd4-218">See also</span></span>

* [<span data-ttu-id="e8bd4-219">Solução de problemas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="e8bd4-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* <span data-ttu-id="e8bd4-220">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="e8bd4-220">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="e8bd4-221">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="e8bd4-221">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="e8bd4-222">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="e8bd4-222">Create custom functions in Excel</span></span>](custom-functions-overview.md)
