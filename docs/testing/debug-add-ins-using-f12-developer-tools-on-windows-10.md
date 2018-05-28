---
title: Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="a19ff-102">Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10</span><span class="sxs-lookup"><span data-stu-id="a19ff-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="a19ff-p101">As ferramentas de desenvolvedor F12 inclu?das no Windows 10 o ajudam a depurar, testar e acelerar suas p?ginas da Web. Voc? tamb?m pode us?-las para desenvolver e a depurar seu suplemento do Office se n?o estiver usando um IDE como o Visual Studio ou se precisar investigar um problema durante a execu??o do suplemento fora do IDE. Voc? pode iniciar as ferramentas de desenvolvedor F12 depois que o suplemento estiver em execu??o.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="a19ff-p102">Este artigo mostra como ? poss?vel usar a ferramenta Depurador das ferramentas de desenvolvedor F12 no Windows 10 para testar seu suplemento do Office. ? poss?vel testar os suplementos do AppSource ou os suplementos que voc? adicionou de outros locais. As ferramentas F12 s?o exibidas em uma janela separada e n?o usam o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="a19ff-p103">O Depurador faz parte das ferramentas de desenvolvedor F12 no Windows 10 e no Internet Explorer. As vers?es anteriores do Windows n?o incluem o Depurador.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="a19ff-111">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="a19ff-111">Prerequisites</span></span>

<span data-ttu-id="a19ff-112">Voc? precisa dos seguintes softwares:</span><span class="sxs-lookup"><span data-stu-id="a19ff-112">You need the following software:</span></span>

- <span data-ttu-id="a19ff-113">As ferramentas do desenvolvedor F12, que est?o inclu?das no Windows 10.</span><span class="sxs-lookup"><span data-stu-id="a19ff-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="a19ff-114">O aplicativo cliente do Office que hospeda seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="a19ff-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="a19ff-115">Seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="a19ff-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="a19ff-116">Utiliza??o do depurador</span><span class="sxs-lookup"><span data-stu-id="a19ff-116">Using the Debugger</span></span>

<span data-ttu-id="a19ff-117">Este exemplo usa o Word e um suplemento gratuito do AppSource.</span><span class="sxs-lookup"><span data-stu-id="a19ff-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="a19ff-118">Abra o Word e escolha um documento em branco.</span><span class="sxs-lookup"><span data-stu-id="a19ff-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="a19ff-p104">Na guia **Inserir**, no grupo Suplementos, escolha **Reposit?rio** e selecione o suplemento QR4Office. ? poss?vel carregar qualquer suplemento da Store ou seu cat?logo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="a19ff-121">Inicie as ferramentas de desenvolvimento F12 que correspondem ? sua vers?o do Office:</span><span class="sxs-lookup"><span data-stu-id="a19ff-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="a19ff-122">Para a vers?o de 32 bits do Office, use C:\Windows\System32\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="a19ff-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="a19ff-123">Para a vers?o de 64 bits do Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="a19ff-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="a19ff-p105">Quando voc? inicia F12Chooser, uma janela separada denominada "Escolher destino para depurar" exibe os poss?veis aplicativos para depurar. Selecione o aplicativo do seu interesse. Se voc? estiver escrevendo seu pr?prio suplemento, selecione o site onde voc? deseja ter o suplemento implantado, que pode ser uma URL de localhost.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p105">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug. Select the application that you are interested in. If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="a19ff-127">Por exemplo, selecione **home.html**.</span><span class="sxs-lookup"><span data-stu-id="a19ff-127">For example, select **home.html**.</span></span> 
    
   ![Tela do F12Chooser, apontando para o suplemento bolhas](../images/choose-target-to-debug.png)

4. <span data-ttu-id="a19ff-129">Na janela F12, selecione o arquivo que voc? deseja depurar.</span><span class="sxs-lookup"><span data-stu-id="a19ff-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="a19ff-p106">Para selecionar o arquivo, escolha o ?cone de pasta acima do painel **script** (? esquerda). A lista suspensa mostra os arquivos dispon?veis. Selecione home.js.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="a19ff-133">Defina o ponto de interrup??o.</span><span class="sxs-lookup"><span data-stu-id="a19ff-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="a19ff-p107">Para definir um ponto de interrup??o, escolha a linha 144, que est? na fun??o _textChanged_. Voc? ver? um ponto vermelho na parte esquerda da linha e uma linha correspondente no painel **Pilha de Chamadas e Pontos de Interrup??o** (parte inferior direita). Para ver outras formas de definir um ponto de interrup??o, confira [Inspecionar executando JavaScript com o Depurador](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="a19ff-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx).</span></span> 
    
   ![Depurador com ponto de interrup??o no arquivo home.js](../images/debugger-home-js-02.png)

6. <span data-ttu-id="a19ff-138">Execute o suplemento para acionar o ponto de interrup??o.</span><span class="sxs-lookup"><span data-stu-id="a19ff-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="a19ff-p108">Escolha a caixa de texto da URL na parte superior do painel QR4Office para alterar o texto. No Depurador, no painel **Pilha de Chamadas e Pontos de Interrup??o**, voc? ver? que o ponto de interrup??o est? disparado e mostra v?rias informa??es. Talvez voc? precise atualizar a ferramenta F12 para ver os resultados.</span><span class="sxs-lookup"><span data-stu-id="a19ff-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![Depurador com resultados do ponto de interrup??o disparado](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="a19ff-143">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="a19ff-143">See also</span></span>

- [<span data-ttu-id="a19ff-144">Inspecionar executando JavaScript com o Depurador</span><span class="sxs-lookup"><span data-stu-id="a19ff-144">Inspect running JavaScript with the Debugger</span></span>](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [<span data-ttu-id="a19ff-145">Usando as ferramentas de desenvolvedor F12</span><span class="sxs-lookup"><span data-stu-id="a19ff-145">Using the F12 developer tools</span></span>](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
