---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 57e0818da0a115380bcb11f1226c48d3ded1078a
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943723"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="21b9c-102">Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10</span><span class="sxs-lookup"><span data-stu-id="21b9c-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="21b9c-p101">As ferramentas de desenvolvedor F12 incluídas no Windows 10 o ajudam a depurar, testar e acelerar suas páginas da Web. Você também pode usá-las para desenvolver e a depurar seu suplemento do Office se não estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execução do suplemento fora do IDE. Você pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execução.</span><span class="sxs-lookup"><span data-stu-id="21b9c-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="21b9c-p102">Este artigo mostra como é possível usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. É possível testar os suplementos do AppSource ou os suplementos que você adicionou de outros locais. As ferramentas F12 são exibidas em uma janela separada e não usam o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="21b9c-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="21b9c-p103">O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As versões anteriores do Windows não incluem o Depurador.</span><span class="sxs-lookup"><span data-stu-id="21b9c-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="21b9c-111">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="21b9c-111">Prerequisites</span></span>

<span data-ttu-id="21b9c-112">Você precisa dos seguintes softwares:</span><span class="sxs-lookup"><span data-stu-id="21b9c-112">You need the following software:</span></span>

- <span data-ttu-id="21b9c-113">As ferramentas do desenvolvedor F12, que estão incluídas no Windows 10.</span><span class="sxs-lookup"><span data-stu-id="21b9c-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="21b9c-114">O aplicativo cliente do Office que hospeda seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="21b9c-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="21b9c-115">Seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="21b9c-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="21b9c-116">Utilização do depurador</span><span class="sxs-lookup"><span data-stu-id="21b9c-116">Using the Debugger</span></span>

<span data-ttu-id="21b9c-117">Este exemplo usa o Word e um suplemento gratuito do AppSource.</span><span class="sxs-lookup"><span data-stu-id="21b9c-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="21b9c-118">Abra o Word e escolha um documento em branco.</span><span class="sxs-lookup"><span data-stu-id="21b9c-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="21b9c-p104">Na guia  **Inserir**, no grupo Suplementos, escolha **Repositório** e selecione o suplemento QR4Office. (É possível carregar qualquer suplemento do Repositório ou de seu catálogo de suplementos.)</span><span class="sxs-lookup"><span data-stu-id="21b9c-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="21b9c-121">Inicie as ferramentas de desenvolvimento F12 que correspondem à sua versão do Office:</span><span class="sxs-lookup"><span data-stu-id="21b9c-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="21b9c-122">Para a versão de 32 bits do Office, use C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="21b9c-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="21b9c-123">Para a versão de 64 bits do Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="21b9c-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="21b9c-124">Quando você inicia IEChooser, uma janela separada denominada "Escolher destino para depurar" exibe os possíveis aplicativos para depurar.</span><span class="sxs-lookup"><span data-stu-id="21b9c-124">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="21b9c-125">Selecione o aplicativo do seu interesse.</span><span class="sxs-lookup"><span data-stu-id="21b9c-125">Select the application that you are interested in.</span></span> <span data-ttu-id="21b9c-126">Se você estiver escrevendo seu próprio suplemento, selecione o site onde você deseja ter o suplemento implantado, que pode ser uma URL de localhost.</span><span class="sxs-lookup"><span data-stu-id="21b9c-126">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="21b9c-127">Por exemplo, selecione **home.html**.</span><span class="sxs-lookup"><span data-stu-id="21b9c-127">For example, select **home.html**.</span></span> 
    
   ![Tela do IEChooser, apontando para o suplemento de bolhas](../images/choose-target-to-debug.png)

4. <span data-ttu-id="21b9c-129">Na janela F12, selecione o arquivo que você deseja depurar.</span><span class="sxs-lookup"><span data-stu-id="21b9c-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="21b9c-p106">Para selecionar o arquivo, escolha o ícone de pasta acima do painel **script** (à esquerda). A lista suspensa mostra os arquivos disponíveis. Selecione home.js.</span><span class="sxs-lookup"><span data-stu-id="21b9c-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="21b9c-133">Defina o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="21b9c-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="21b9c-p107">Para definir um ponto de interrupção, escolha a linha 144, que está na função _textChanged_. Você verá um ponto vermelho na parte esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas e Pontos de Interrupção** (parte inferior direita). Para ver outras formas de definir um ponto de interrupção, confira [Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="21b9c-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![Depurador com ponto de interrupção no arquivo home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="21b9c-138">Execute o suplemento para acionar o ponto de interrupção.</span><span class="sxs-lookup"><span data-stu-id="21b9c-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="21b9c-p108">Escolha a caixa de texto da URL na parte superior do painel QR4Office para alterar o texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrupção**, você verá que o ponto de interrupção está disparado e mostra várias informações. Talvez você precise atualizar a ferramenta F12 para ver os resultados.</span><span class="sxs-lookup"><span data-stu-id="21b9c-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Depurador com resultados do ponto de interrupção disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="21b9c-143">Veja também</span><span class="sxs-lookup"><span data-stu-id="21b9c-143">See also</span></span>

- <span data-ttu-id="21b9c-144">[Inspecionar executando JavaScript com o Depurador](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="21b9c-144">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="21b9c-145">[Usando as ferramentas de desenvolvedor F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="21b9c-145">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
    
