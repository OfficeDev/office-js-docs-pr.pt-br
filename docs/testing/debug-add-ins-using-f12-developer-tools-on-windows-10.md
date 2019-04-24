---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 750411bea187a0ade9b3723e3198d82f7c482c9f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450135"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="1077e-102">Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10</span><span class="sxs-lookup"><span data-stu-id="1077e-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="1077e-103">As ferramentas de desenvolvedor F12 incluídas no Windows 10 o ajudam a depurar, testar e acelerar suas páginas da Web.</span><span class="sxs-lookup"><span data-stu-id="1077e-103">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages.</span></span> <span data-ttu-id="1077e-104">Você também pode usá-las para desenvolver e a depurar seu suplemento do Office se não estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execução do suplemento fora do IDE.</span><span class="sxs-lookup"><span data-stu-id="1077e-104">You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE.</span></span> <span data-ttu-id="1077e-105">Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="1077e-105">This article describes how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1077e-106">As instruções neste artigo não podem ser utilizadas para depurar um suplemento do Outlook que usa Funções Executar.</span><span class="sxs-lookup"><span data-stu-id="1077e-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="1077e-107">Para depurar um suplemento do Outlook que usa Funções Executar, é recomendável que você anexe ao Visual Studio no modo de script ou outro depurador de scripts.</span><span class="sxs-lookup"><span data-stu-id="1077e-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1077e-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="1077e-108">Prerequisites</span></span>

<span data-ttu-id="1077e-109">Você precisa dos seguintes softwares:</span><span class="sxs-lookup"><span data-stu-id="1077e-109">You need the following software:</span></span>

- <span data-ttu-id="1077e-110">As ferramentas do desenvolvedor F12, que estão incluídas no Windows 10.</span><span class="sxs-lookup"><span data-stu-id="1077e-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="1077e-111">O aplicativo cliente do Office que hospeda seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1077e-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="1077e-112">Seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="1077e-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="1077e-113">Utilização do depurador</span><span class="sxs-lookup"><span data-stu-id="1077e-113">Using the Debugger</span></span>

<span data-ttu-id="1077e-114">Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. É possível testar os suplementos do AppSource ou os suplementos que você adicionou de outros locais. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="1077e-114">You can use the Debugger from the F12 developer tools in Windows 10 to test add-ins from AppSource or add-ins that you have added from other locations.</span></span> <span data-ttu-id="1077e-115">Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="1077e-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="1077e-116">As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="1077e-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="1077e-p104">O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador.</span><span class="sxs-lookup"><span data-stu-id="1077e-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="1077e-119">Este exemplo usa o Word e um suplemento gratuito do AppSource.</span><span class="sxs-lookup"><span data-stu-id="1077e-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="1077e-120">Abra o Word e escolha um documento em branco.</span><span class="sxs-lookup"><span data-stu-id="1077e-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="1077e-121">Na guia **Inserir**, no grupo Suplementos e selecione **Store** e selecione o suplemento **QR4Office**.</span><span class="sxs-lookup"><span data-stu-id="1077e-121">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="1077e-122">(Você pode carregar qualquer suplemento da Store ou seu catálogo de suplemento.)</span><span class="sxs-lookup"><span data-stu-id="1077e-122">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="1077e-123">Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:</span><span class="sxs-lookup"><span data-stu-id="1077e-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="1077e-124">Para a versão de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="1077e-124">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="1077e-125">Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="1077e-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="1077e-126">Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar.</span><span class="sxs-lookup"><span data-stu-id="1077e-126">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="1077e-127">Selecione o aplicativo do seu interesse.</span><span class="sxs-lookup"><span data-stu-id="1077e-127">Select the application that you are interested in.</span></span> <span data-ttu-id="1077e-128">Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost.</span><span class="sxs-lookup"><span data-stu-id="1077e-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="1077e-129">Por exemplo, selecione **home.html**.</span><span class="sxs-lookup"><span data-stu-id="1077e-129">For example, select **home.html**.</span></span> 
    
   ![Tela do IEChooser, apontando para o suplemento bolhas](../images/choose-target-to-debug.png)

4. <span data-ttu-id="1077e-131">Na janela F12, selecione o arquivo que você deseja depurar.</span><span class="sxs-lookup"><span data-stu-id="1077e-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="1077e-132">Para selecionar o arquivo na janela F12, escolha o ícone de pasta acima do painel **script** (à esquerda).</span><span class="sxs-lookup"><span data-stu-id="1077e-132">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="1077e-133">Na lista de arquivos disponíveis exibido na lista suspensa, selecione **Home.js**.</span><span class="sxs-lookup"><span data-stu-id="1077e-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="1077e-134">Defina o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="1077e-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="1077e-135">Para definir o ponto de interrupção no **Home.js**, escolha a linha 144, que está na função  `textChanged`.</span><span class="sxs-lookup"><span data-stu-id="1077e-135">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="1077e-136">Você verá um ponto vermelho à esquerda da linha e uma linha correspondente no painel Pilha de Chamadas e Pontos de Interrupção (canto inferior direito).</span><span class="sxs-lookup"><span data-stu-id="1077e-136">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="1077e-137">Para ver outras maneiras de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="1077e-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="1077e-139">Execute o suplemento para acionar o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="1077e-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="1077e-140">No Word, escolha a caixa de texto na parte superior da URL do painel **QR4Office** e tente digitar algum texto.</span><span class="sxs-lookup"><span data-stu-id="1077e-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="1077e-141">No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações.</span><span class="sxs-lookup"><span data-stu-id="1077e-141">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="1077e-142">Talvez você precise atualizar o depurador para ver os resultados.</span><span class="sxs-lookup"><span data-stu-id="1077e-142">You might need to refresh the Debugger to see the results.</span></span>
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="1077e-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="1077e-144">See also</span></span>

- <span data-ttu-id="1077e-145">[Inspecionar executando JavaScript com o Depurador](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="1077e-145">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="1077e-146">[Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="1077e-146">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
