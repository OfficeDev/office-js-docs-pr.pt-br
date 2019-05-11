---
title: Anexar um depurador do painel de tarefas
description: ''
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 03926ea18963b98f44702f7213dd1768e9924265
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952275"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="715f8-102">Anexar um depurador do painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="715f8-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="715f8-p101">No Office 2016 no Windows, Build 77xx.xxxx ou posterior, é possível anexar o depurador do painel de tarefas. O recurso de anexar o depurador anexará diretamente o depurador ao processo correto do Internet Explorer. É possível anexar um depurador independentemente de você estar utilizando Yeoman Generator, Visual Studio Code, Node.js, Angular ou outra ferramenta.</span><span class="sxs-lookup"><span data-stu-id="715f8-p101">In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="715f8-106">Para iniciar a ferramenta **Anexar Depurador**, escolha o canto superior direito do painel de tarefas para ativar o menu **Personalidade** (conforme mostrado no círculo vermelho na imagem a seguir).</span><span class="sxs-lookup"><span data-stu-id="715f8-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="715f8-p102">Atualmente, a única ferramenta de depurador é o [Visual Studio 2015](https://www.visualstudio.com/downloads/) com a [Atualização 3](https://msdn.microsoft.com/library/mt752379.aspx) ou posterior. Se você não instalou o Visual Studio, selecionar a opção **Anexar Depurador** não resultará em nenhuma ação.</span><span class="sxs-lookup"><span data-stu-id="715f8-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="715f8-p103">Só é possível depurar o JavaScript do lado do cliente com a ferramenta **Anexar Depurador**. Para depurar o código do lado do servidor, como com um servidor Node.js, há várias opções. Confira informações sobre como depurar com o Visual Studio Code em [Depuração do Node.js no VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Se você não estiver usando o Visual Studio Code, pesquise por "depurar Node.js" ou "depurar {nome do servidor}".</span><span class="sxs-lookup"><span data-stu-id="715f8-p103">You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![Captura de tela do menu Anexar Depurador](../images/attach-debugger.png)

<span data-ttu-id="715f8-p104">Selecione **Anexar Depurador**. Isso inicia a caixa de diálogo **Depurador Just-In-Time do Visual Studio**, conforme mostrado na imagem a seguir.</span><span class="sxs-lookup"><span data-stu-id="715f8-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![Captura de tela da caixa de diálogo Depurador JIT do Visual Studio](../images/visual-studio-debugger.png)

<span data-ttu-id="715f8-p105">No Visual Studio, você verá os arquivos de código no **Gerenciador de Soluções**.   Você pode definir pontos de interrupção na linha de código que deseja depurar no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="715f8-p105">In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="715f8-119">Se você não vir o menu Personalidade, é possível depurar o suplemento com o Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="715f8-119">If you don't see the Personality menu, you can debug your add-in using Visual Studio.</span></span> <span data-ttu-id="715f8-120">Certifique-se de que o suplemento do painel tarefas esteja aberto no Office e, em seguida, siga estas etapas:</span><span class="sxs-lookup"><span data-stu-id="715f8-120">Ensure your task pane add-in is open in Office, and then follow these steps:</span></span>

> 1. <span data-ttu-id="715f8-121">No Visual Studio, escolha **DEPURAR** > **Anexar ao Processo**.</span><span class="sxs-lookup"><span data-stu-id="715f8-121">In Visual Studio, choose **DEBUG** > **Attach to Process**.</span></span>
> 2. <span data-ttu-id="715f8-122">Em **Anexar ao Processo**, escolha todos os processos Iexplore.exe disponíveis e, em seguida, selecione o botão **Anexar**.</span><span class="sxs-lookup"><span data-stu-id="715f8-122">In **Attach to Process**, choose all of the available Iexplore.exe processes, and then choose the **Attach** button.</span></span>

<span data-ttu-id="715f8-123">Veja mais informações sobre depuração no Visual Studio, em:</span><span class="sxs-lookup"><span data-stu-id="715f8-123">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="715f8-124">Para iniciar e usar o Explorador do DOM no Visual Studio, confira a Dica 4 na seção [Dicas e Truques](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) da publicação [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) (Criar aplicativos atraentes para o Office usando os novos modelos de projeto) do blog.</span><span class="sxs-lookup"><span data-stu-id="715f8-124">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="715f8-125">Para definir pontos de interrupção, confira [Usar Pontos de Interrupção](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span><span class="sxs-lookup"><span data-stu-id="715f8-125">To set breakpoints, see [Using Breakpoints](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span></span>
-   <span data-ttu-id="715f8-126">Para usar o F12, confira o artigo [Usando as ferramentas de desenvolvedor F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="715f8-126">To use F12, see [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>

## <a name="see-also"></a><span data-ttu-id="715f8-127">Confira também</span><span class="sxs-lookup"><span data-stu-id="715f8-127">See also</span></span>

- [<span data-ttu-id="715f8-128">Criar e depurar suplementos do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="715f8-128">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="715f8-129">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="715f8-129">Publish your Office Add-in</span></span>](../publish/publish.md)
