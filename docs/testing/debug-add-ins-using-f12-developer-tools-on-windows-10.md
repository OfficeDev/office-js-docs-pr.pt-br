---
title: Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: a2090eca41f59f0e7fab1a172aff96cbbca28ed7
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454878"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="732c4-102">Depurar suplementos usando as ferramentas de desenvolvedor no Windows 10</span><span class="sxs-lookup"><span data-stu-id="732c4-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="732c4-103">Há ferramentas de desenvolvedor fora dos IDEs disponíveis para ajudá-lo a depurar seus suplementos no Windows 10.</span><span class="sxs-lookup"><span data-stu-id="732c4-103">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="732c4-104">Elas são úteis quando você precisa investigar um problema enquanto executa seu suplemento fora do IDE.</span><span class="sxs-lookup"><span data-stu-id="732c4-104">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="732c4-105">A ferramenta que você usa depende se o suplemento está sendo executado no Microsoft Edge ou no Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="732c4-105">The tool that you use depends on whether the add-in is running in Edge or Internet Explorer.</span></span> <span data-ttu-id="732c4-106">Isso é determinado pela versão do Windows 10 e a versão do Office que estão instaladas no computador.</span><span class="sxs-lookup"><span data-stu-id="732c4-106">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="732c4-107">Para determinar qual navegador está sendo usado em seu computador de desenvolvimento, confira [Navegadores usados pelos Suplementos do Office](../concepts/browsers-used-by-office-web-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="732c4-107">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span> 


> [!NOTE]
> <span data-ttu-id="732c4-108">As instruções neste artigo não podem ser utilizadas para depurar um suplemento do Outlook que usa Funções Executar.</span><span class="sxs-lookup"><span data-stu-id="732c4-108">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="732c4-109">Para depurar um suplemento do Outlook que usa Funções Executar, é recomendável que você anexe ao Visual Studio no modo de script ou outro depurador de scripts.</span><span class="sxs-lookup"><span data-stu-id="732c4-109">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-edge"></a><span data-ttu-id="732c4-110">Quando o suplemento estiver sendo executado no Edge</span><span class="sxs-lookup"><span data-stu-id="732c4-110">When the add-in is running in Edge</span></span>

<span data-ttu-id="732c4-111">Quando o suplemento estiver sendo executado, você pode usar o [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span><span class="sxs-lookup"><span data-stu-id="732c4-111">When the add-in is running in Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span> 

1. <span data-ttu-id="732c4-112">Execute o suplemento.</span><span class="sxs-lookup"><span data-stu-id="732c4-112">Run the add-in</span></span> 

2. <span data-ttu-id="732c4-113">Execute o Microsoft Edge DevTools.</span><span class="sxs-lookup"><span data-stu-id="732c4-113">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="732c4-114">Nas ferramentas, abra a guia **Local**. Seu suplemento será listado por nome.</span><span class="sxs-lookup"><span data-stu-id="732c4-114">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="732c4-115">Clique no nome do suplemento para abri-lo nas ferramentas.</span><span class="sxs-lookup"><span data-stu-id="732c4-115">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="732c4-116">Abra a guia **Depurador**.</span><span class="sxs-lookup"><span data-stu-id="732c4-116">Open the **Sharing** tab.</span></span> 

6. <span data-ttu-id="732c4-117">Escolha o ícone de pasta acima do painel **script** (à esquerda).</span><span class="sxs-lookup"><span data-stu-id="732c4-117">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="732c4-118">Na lista de arquivos disponíveis exibidos na lista suspensa, selecione o arquivo JavaScript que você deseja depurar.</span><span class="sxs-lookup"><span data-stu-id="732c4-118">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="732c4-119">Para definir um ponto de interrupção, selecione a linha.</span><span class="sxs-lookup"><span data-stu-id="732c4-119">To set a breakpoint, select the line.</span></span> <span data-ttu-id="732c4-120">Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas** (canto inferior direito).</span><span class="sxs-lookup"><span data-stu-id="732c4-120">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span>

8. <span data-ttu-id="732c4-121">Execute funções no suplemento conforme necessário para disparar o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="732c4-121">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="732c4-122">Quando o suplemento estiver sendo executado no Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="732c4-122">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="732c4-123">Quando o suplemento estiver sendo executado no Internet Explorer, você poderá usar o depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="732c4-123">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="732c4-124">Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="732c4-124">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="732c4-125">As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="732c4-125">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="732c4-p107">O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador.</span><span class="sxs-lookup"><span data-stu-id="732c4-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="732c4-128">Este exemplo usa o Word e um suplemento gratuito do AppSource.</span><span class="sxs-lookup"><span data-stu-id="732c4-128">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="732c4-129">Abra o Word e escolha um documento em branco.</span><span class="sxs-lookup"><span data-stu-id="732c4-129">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="732c4-130">Na guia **Inserir**, no grupo Suplementos e selecione **Store** e selecione o suplemento **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="732c4-130">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="732c4-131">(Você pode carregar qualquer suplemento da Store ou seu catálogo de suplemento.)</span><span class="sxs-lookup"><span data-stu-id="732c4-131">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="732c4-132">Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:</span><span class="sxs-lookup"><span data-stu-id="732c4-132">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="732c4-133">Para a versão de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="732c4-133">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="732c4-134">Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="732c4-134">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="732c4-135">Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar.</span><span class="sxs-lookup"><span data-stu-id="732c4-135">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="732c4-136">Selecione o aplicativo do seu interesse.</span><span class="sxs-lookup"><span data-stu-id="732c4-136">Select the application that you are interested in.</span></span> <span data-ttu-id="732c4-137">Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost.</span><span class="sxs-lookup"><span data-stu-id="732c4-137">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="732c4-138">Por exemplo, selecione **home.html**.</span><span class="sxs-lookup"><span data-stu-id="732c4-138">For example, select **home.html**.</span></span> 
    
   ![Tela do IEChooser, apontando para o suplemento bolhas](../images/choose-target-to-debug.png)

4. <span data-ttu-id="732c4-140">Na janela F12, selecione o arquivo que você deseja depurar.</span><span class="sxs-lookup"><span data-stu-id="732c4-140">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="732c4-141">Para selecionar o arquivo na janela F12, escolha o ícone de pasta acima do painel **script** (à esquerda).</span><span class="sxs-lookup"><span data-stu-id="732c4-141">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="732c4-142">Na lista de arquivos disponíveis exibido na lista suspensa, selecione **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="732c4-142">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="732c4-143">Defina o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="732c4-143">Set the breakpoint.</span></span>
    
   <span data-ttu-id="732c4-144">Para definir o ponto de interrupção no **Home.js**, escolha a linha 144, que está na função  `textChanged`.</span><span class="sxs-lookup"><span data-stu-id="732c4-144">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="732c4-145">Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel Pilha de Chamadas e Pontos de Interrupção (canto inferior direito).</span><span class="sxs-lookup"><span data-stu-id="732c4-145">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="732c4-146">Para ver outras maneiras de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="732c4-146">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="732c4-148">Execute o suplemento para acionar o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="732c4-148">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="732c4-149">No Word, escolha a caixa de texto na parte superior da URL do painel **QR4Office** e tente digitar algum texto.</span><span class="sxs-lookup"><span data-stu-id="732c4-149">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="732c4-150">No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações.</span><span class="sxs-lookup"><span data-stu-id="732c4-150">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="732c4-151">Talvez você precise atualizar o depurador para ver os resultados.</span><span class="sxs-lookup"><span data-stu-id="732c4-151">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="732c4-153">Confira também</span><span class="sxs-lookup"><span data-stu-id="732c4-153">See also</span></span>

- <span data-ttu-id="732c4-154">[Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="732c4-154">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="732c4-155">[Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="732c4-155">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
