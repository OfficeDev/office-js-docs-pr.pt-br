---
ms.date: 04/12/2021
description: Saiba como depurar suas Excel funções personalizadas que não usam um painel de tarefas.
title: Depuração de funções personalizadas sem interface do usuário
localization_priority: Normal
ms.openlocfilehash: e0e2b7bf49836a9b88de9ceaa21a66a454e6f05a
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349641"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="f1c6d-103">Depuração de funções personalizadas sem interface do usuário</span><span class="sxs-lookup"><span data-stu-id="f1c6d-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="f1c6d-104">Este artigo discute a depuração *apenas* para funções personalizadas que não usam um painel de tarefas ou outros elementos de interface do usuário (funções personalizadas sem interface do usuário).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="f1c6d-105">No Windows:</span><span class="sxs-lookup"><span data-stu-id="f1c6d-105">On Windows:</span></span>
- [<span data-ttu-id="f1c6d-106">Excel Depurador Visual Studio Code (VS Code)</span><span class="sxs-lookup"><span data-stu-id="f1c6d-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="f1c6d-107">Excel na Web e VS Code depurador</span><span class="sxs-lookup"><span data-stu-id="f1c6d-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="f1c6d-108">Excel na Web e ferramentas do navegador</span><span class="sxs-lookup"><span data-stu-id="f1c6d-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="f1c6d-109">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="f1c6d-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="f1c6d-110">No Mac:</span><span class="sxs-lookup"><span data-stu-id="f1c6d-110">On Mac:</span></span>
- [<span data-ttu-id="f1c6d-111">Excel na Web e ferramentas do navegador</span><span class="sxs-lookup"><span data-stu-id="f1c6d-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="f1c6d-112">Linha de comando</span><span class="sxs-lookup"><span data-stu-id="f1c6d-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="f1c6d-113">Para simplificar, este artigo mostra a depuração no contexto de uso Visual Studio Code para editar, executar tarefas e, em alguns casos, usar o modo de exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="f1c6d-114">Se você estiver usando uma ferramenta de linha de comando ou editor diferente, consulte [as](#commands-for-building-and-running-your-add-in) instruções de linha de comando no final deste artigo.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="f1c6d-115">Requisitos</span><span class="sxs-lookup"><span data-stu-id="f1c6d-115">Requirements</span></span>

<span data-ttu-id="f1c6d-116">Esse processo de depuração funciona **apenas** para funções personalizadas sem interface do usuário, que não usam um painel de tarefas ou outros elementos da interface do usuário.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="f1c6d-117">Uma função personalizada sem interface do usuário pode ser criada seguindo as etapas no tutorial Criar funções [personalizadas](../tutorials/excel-tutorial-create-custom-functions.md) no Excel e, em seguida, remover todos os elementos do painel de tarefas e da interface do usuário instalados pelo gerador [Yeoman para Office Add-ins](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="f1c6d-118">Observe que esse processo de depuração não é compatível com projetos de funções personalizadas usando um [tempo de execução compartilhado.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="f1c6d-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="f1c6d-119">Use o VS Code depurador para Excel Desktop</span><span class="sxs-lookup"><span data-stu-id="f1c6d-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="f1c6d-120">Você pode usar VS Code para depurar funções personalizadas sem interface do usuário Office Excel na área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="f1c6d-121">A depuração de área de trabalho para o Mac não está disponível, mas pode ser atingida usando as ferramentas do navegador e a linha de comando para [depurar](#use-the-command-line-tools-to-debug)Excel na Web ).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="f1c6d-122">Execute o seu complemento do VS Code</span><span class="sxs-lookup"><span data-stu-id="f1c6d-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="f1c6d-123">Abra sua pasta de projeto raiz de funções personalizadas [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f1c6d-124">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f1c6d-125">Isso monitorará e reconstruirá todas as alterações de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f1c6d-126">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="f1c6d-127">Iniciar o VS Code depurador</span><span class="sxs-lookup"><span data-stu-id="f1c6d-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="f1c6d-128">Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="f1c6d-129">No menu suspenso Executar, escolha Excel **Desktop (Funções Personalizadas)**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-129">From the Run drop-down menu, choose **Excel Desktop (Custom Functions)**.</span></span>
6. <span data-ttu-id="f1c6d-130">Selecione **F5** (ou selecione **Executar -> Iniciar Depuração** no menu) para começar a depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="f1c6d-131">Uma nova Excel de trabalho será aberta com seu complemento já sideload e pronto para uso.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="f1c6d-132">Iniciar a depuração</span><span class="sxs-lookup"><span data-stu-id="f1c6d-132">Start debugging</span></span>

1. <span data-ttu-id="f1c6d-133">Em VS Code, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="f1c6d-134">[Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="f1c6d-135">Na Excel de trabalho, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="f1c6d-136">Neste ponto, a execução será parada na linha de código onde você definirá o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="f1c6d-137">Agora você pode passar pelo código, definir relógios e usar qualquer VS Code recursos de depuração necessários.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="f1c6d-138">Use o VS Code depurador para Excel em Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="f1c6d-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="f1c6d-139">Você pode usar VS Code para depurar funções personalizadas sem interface do usuário Excel no navegador Microsoft Edge usuário.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="f1c6d-140">Para usar VS Code com Microsoft Edge, você deve instalar o [Depurador para Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extensão.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="f1c6d-141">Execute o seu complemento do VS Code</span><span class="sxs-lookup"><span data-stu-id="f1c6d-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="f1c6d-142">Abra sua pasta de projeto raiz de funções personalizadas [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f1c6d-143">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f1c6d-144">Isso monitorará e reconstruirá todas as alterações de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f1c6d-145">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="f1c6d-146">Iniciar o VS Code depurador</span><span class="sxs-lookup"><span data-stu-id="f1c6d-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="f1c6d-147">Escolha **Exibir > Executar ou** insira **Ctrl+Shift+D** para alternar para o exibição de depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="f1c6d-148">Nas opções Depurar, escolha **Office Online (Edge Chromium)**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="f1c6d-149">Abra Excel no navegador Microsoft Edge e crie uma nova workbook.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="f1c6d-150">Escolha **Compartilhar** na faixa de opções e copie o link para a URL dessa nova workbook.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="f1c6d-151">Selecione **F5** (ou **selecione Executar > Iniciar Depuração** no menu) para começar a depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="f1c6d-152">Um prompt será exibido, que solicita a URL do documento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="f1c6d-153">Colar na URL da pasta de trabalho e pressione Enter.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="f1c6d-154">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="f1c6d-154">Sideload your add-in</span></span>

1. <span data-ttu-id="f1c6d-155">Selecione a **guia** Inserir na faixa de opções e, na seção **Complementos,** escolha Office **Adicionar.**</span><span class="sxs-lookup"><span data-stu-id="f1c6d-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="f1c6d-156">Na caixa **de diálogo Office de** Office, selecione a guia MEUS **ADD-INS,** escolha Gerenciar Meus **Complementos** e, em seguida, **Upload Meu Complemento**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![A caixa Office de Office com um drop-down na leitura superior direita "Gerenciar meus complementos" e um drop-down abaixo dele com a opção "Upload Meu Complemento".](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="f1c6d-158">**Navegue** até o arquivo de manifesto do complemento e selecione **Upload**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="f1c6d-160">Definir pontos de interrupção</span><span class="sxs-lookup"><span data-stu-id="f1c6d-160">Set breakpoints</span></span>
1. <span data-ttu-id="f1c6d-161">Em VS Code, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="f1c6d-162">[Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="f1c6d-163">Na Excel de trabalho, insira uma fórmula que usa sua função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="f1c6d-164">Use as ferramentas de desenvolvedor do navegador para depurar funções personalizadas em Excel na Web</span><span class="sxs-lookup"><span data-stu-id="f1c6d-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="f1c6d-165">Você pode usar as ferramentas de desenvolvedor do navegador para depurar funções personalizadas sem interface do usuário Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="f1c6d-166">As etapas a seguir funcionam para o Windows e macOS.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="f1c6d-167">Execute o seu complemento do Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="f1c6d-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="f1c6d-168">Abra sua pasta de projeto raiz de funções personalizadas [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="f1c6d-169">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Assistir**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="f1c6d-170">Isso monitorará e reconstruirá todas as alterações de arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="f1c6d-171">Escolha **Terminal > Executar Tarefa** e digite ou selecione **Dev Server**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="f1c6d-172">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="f1c6d-172">Sideload your add-in</span></span>

1. <span data-ttu-id="f1c6d-173">Abra [Office na Web](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="f1c6d-174">Abra uma nova Excel de trabalho.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="f1c6d-175">Abra a **guia** Inserir na faixa de opções e, na seção **Add-ins,** escolha Office **Adicionar.**</span><span class="sxs-lookup"><span data-stu-id="f1c6d-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="f1c6d-176">Na caixa **de diálogo Office de** Office, selecione a guia MEUS **ADD-INS,** escolha Gerenciar Meus **Complementos** e, em seguida, **Upload Meu Complemento**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![A caixa Office de Office com um drop-down na leitura superior direita "Gerenciar meus complementos" e um drop-down abaixo dele com a opção "Upload Meu Complemento".](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="f1c6d-178">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="f1c6d-180">Depois de fazer sideload no documento, ele permanecerá sideload sempre que você abrir o documento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="f1c6d-181">Iniciar a depuração</span><span class="sxs-lookup"><span data-stu-id="f1c6d-181">Start debugging</span></span>

1. <span data-ttu-id="f1c6d-182">Abra ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-182">Open developer tools in the browser.</span></span> <span data-ttu-id="f1c6d-183">Para o Chrome e a maioria dos navegadores F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="f1c6d-184">Em ferramentas de desenvolvedor, abra seu arquivo de script de código-fonte usando **Cmd+P** ou **Ctrl+P** (**functions.js** **ou functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="f1c6d-185">[Definir um ponto de interrupção](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) no código-fonte da função personalizada.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="f1c6d-186">Se você precisar alterar o código, poderá fazer edições no VS Code e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="f1c6d-187">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="f1c6d-188">Usar as ferramentas de linha de comando para depurar</span><span class="sxs-lookup"><span data-stu-id="f1c6d-188">Use the command line tools to debug</span></span>

<span data-ttu-id="f1c6d-189">Se você não estiver usando VS Code, poderá usar a linha de comando (como bash ou PowerShell) para executar o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="f1c6d-190">Você precisará usar as ferramentas de desenvolvedor do navegador para depurar seu código Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="f1c6d-191">Não é possível depurar a versão da área de trabalho Excel usando a linha de comando.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="f1c6d-192">Na linha de comando, `npm run watch` execute para observar e reconstruir quando ocorrerem alterações de código.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="f1c6d-193">Abra uma segunda janela de linha de comando (a primeira será bloqueada durante a execução do relógio).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="f1c6d-194">Se você quiser iniciar o seu complemento na versão da área de trabalho Excel, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-194">If you want to start your add-in in the desktop version of Excel, run the following command.</span></span>

    `npm run start:desktop`

    <span data-ttu-id="f1c6d-195">Ou se você preferir iniciar o seu Excel na Web, execute o seguinte comando.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-195">Or if you prefer to start your add-in in Excel on the web, run the following command.</span></span>

    `npm run start:web`

    <span data-ttu-id="f1c6d-196">Para Excel na Web você também precisa fazer sideload do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="f1c6d-197">Siga as etapas em [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="f1c6d-198">Em seguida, continue até a próxima seção para iniciar a depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-198">Then continue to the next section to start debugging.</span></span>

4. <span data-ttu-id="f1c6d-199">Abra ferramentas de desenvolvedor no navegador.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-199">Open developer tools in the browser.</span></span> <span data-ttu-id="f1c6d-200">Para o Chrome e a maioria dos navegadores F12 abrirá as ferramentas de desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="f1c6d-201">Em ferramentas de desenvolvedor, abra seu arquivo de script de código-fonte (**functions.js** **ou functions.ts**).</span><span class="sxs-lookup"><span data-stu-id="f1c6d-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="f1c6d-202">Seu código de funções personalizadas pode estar localizado perto do final do arquivo.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="f1c6d-203">No código-fonte da função personalizada, aplique um ponto de interrupção selecionando uma linha de código.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="f1c6d-204">Se você precisar alterar o código, poderá fazer edições no Visual Studio e salvar as alterações.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="f1c6d-205">Atualize o navegador para ver as alterações carregadas.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="f1c6d-206">Comandos para criar e executar o seu complemento</span><span class="sxs-lookup"><span data-stu-id="f1c6d-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="f1c6d-207">Há várias tarefas de com build disponíveis:</span><span class="sxs-lookup"><span data-stu-id="f1c6d-207">There are several build tasks available:</span></span>
- <span data-ttu-id="f1c6d-208">`npm run watch`: cria para desenvolvimento e recria automaticamente quando um arquivo de origem é salvo</span><span class="sxs-lookup"><span data-stu-id="f1c6d-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="f1c6d-209">`npm run build-dev`: builds para desenvolvimento uma vez</span><span class="sxs-lookup"><span data-stu-id="f1c6d-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="f1c6d-210">`npm run build`: builds para produção</span><span class="sxs-lookup"><span data-stu-id="f1c6d-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="f1c6d-211">`npm run dev-server`: executa o servidor Web usado para desenvolvimento</span><span class="sxs-lookup"><span data-stu-id="f1c6d-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="f1c6d-212">Você pode usar as seguintes tarefas para iniciar a depuração na área de trabalho ou online.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="f1c6d-213">`npm run start:desktop`: Inicia Excel na área de trabalho e faz o sideload do seu complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="f1c6d-214">`npm run start:web`: Inicia Excel na Web e descarrega o seu complemento.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="f1c6d-215">`npm run stop`: Interrompe Excel e depuração.</span><span class="sxs-lookup"><span data-stu-id="f1c6d-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f1c6d-216">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="f1c6d-216">Next steps</span></span>
<span data-ttu-id="f1c6d-217">Saiba mais sobre as práticas de autenticação para funções [personalizadas sem interface do usuário.](custom-functions-authentication.md)</span><span class="sxs-lookup"><span data-stu-id="f1c6d-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f1c6d-218">Confira também</span><span class="sxs-lookup"><span data-stu-id="f1c6d-218">See also</span></span>

* [<span data-ttu-id="f1c6d-219">Solução de problemas de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f1c6d-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="f1c6d-220">Tratamento de erros para funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="f1c6d-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="f1c6d-221">Criar funções personalizadas no Excel</span><span class="sxs-lookup"><span data-stu-id="f1c6d-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
