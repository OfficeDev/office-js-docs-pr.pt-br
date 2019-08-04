---
title: Realizar sideload de suplementos do Office no Office na Web para teste
description: Testar o suplemento do Office no Office na web através de sideloading
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: 6a61a8bfb4860ac31803c40d8ecea1b550f79368
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575601"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="d980c-103">Realizar sideload de suplementos do Office no Office na Web para teste</span><span class="sxs-lookup"><span data-stu-id="d980c-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="d980c-104">Você pode instalar um suplemento do Office para teste usando sideloading, sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="d980c-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="d980c-105">O sideloading pode ser realizado no Office 365 ou no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="d980c-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="d980c-106">O procedimento é ligeiramente diferente nas duas plataformas.</span><span class="sxs-lookup"><span data-stu-id="d980c-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="d980c-107">Quando você realiza o sideload de um suplemento, o manifesto do suplemento é armazenado localmente do navegador e, portanto, se você limpar o cache do navegador ou alternar para um navegador diferente, precisará realizar o sideload do suplemento novamente.</span><span class="sxs-lookup"><span data-stu-id="d980c-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="d980c-p102">A realização do sideload como descrito neste artigo tem suporte no Word, no Excel e no PowerPoint. Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="d980c-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="d980c-110">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento no Office na Web ou para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d980c-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="d980c-111">Realizar sideload de um suplemento do Office no Office na Web</span><span class="sxs-lookup"><span data-stu-id="d980c-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="d980c-112">Abra o [Microsoft Office na Web](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="d980c-112">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="d980c-113">Em **Comece a usar os aplicativos online agora**, escolha **Excel**, **Word** ou **PowerPoint** e abra um novo documento.</span><span class="sxs-lookup"><span data-stu-id="d980c-113">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="d980c-114">Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="d980c-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="d980c-115">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="d980c-115">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="d980c-117">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="d980c-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

6. <span data-ttu-id="d980c-p103">Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.</span><span class="sxs-lookup"><span data-stu-id="d980c-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="d980c-122">Para testar o suplemento do Office com o Microsoft Edge, são necessárias duas etapas de configuração:</span><span class="sxs-lookup"><span data-stu-id="d980c-122">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="d980c-123">Em um prompt de comando do Windows, execute a seguinte linha: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="d980c-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="d980c-124">Digite “**about:flags**” na barra de pesquisa do Microsoft Edge para exibir as opções de Configurações do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="d980c-124">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="d980c-125">Verifique a opção “**Permitir loopback do localhost**” e reinicie o Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="d980c-125">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![A opção “Permitir loopback do localhost” do Microsoft Edge com a caixa marcada.](../images/allow-localhost-loopback.png)


## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="d980c-127">Realizar sideload de um suplemento do Office no Office 365</span><span class="sxs-lookup"><span data-stu-id="d980c-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="d980c-128">Entre em sua conta do Office 365.</span><span class="sxs-lookup"><span data-stu-id="d980c-128">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="d980c-129">Abra o inicializador de aplicativos à esquerda da barra de ferramentas, selecione  **Excel**, **Word** ou **PowerPoint** e crie um novo documento.</span><span class="sxs-lookup"><span data-stu-id="d980c-129">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="d980c-130">As etapas 3 a 6 são as mesmas da seção anterior **Realize sideload para um suplemento do Office no Office na Web**. </span><span class="sxs-lookup"><span data-stu-id="d980c-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office Online**.</span></span>


## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="d980c-131">Sideload de um suplemento usando o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d980c-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="d980c-132">Se estiver usando o Visual Studio para desenvolver o suplemento, o processo de sideload é semelhante.</span><span class="sxs-lookup"><span data-stu-id="d980c-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="d980c-133">A única diferença é que você deve atualizar o valor do elemento **SourceURL** no manifesto para incluir a URL completa em que o suplemento for implantado.</span><span class="sxs-lookup"><span data-stu-id="d980c-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="d980c-134">Embora você possa realizar o sideload de suplementos do Visual Studio para o Office na Web, não é possível depurá-los no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d980c-134">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="d980c-135">Para depurar você precisará usar as ferramentas de depuração do navegador.</span><span class="sxs-lookup"><span data-stu-id="d980c-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="d980c-136">Para saber mais, confira [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="d980c-136">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="d980c-137">No Visual Studio, abra a janela **Propriedades** escolhendo **Modo de exibição** -> **Janela de propriedades**.</span><span class="sxs-lookup"><span data-stu-id="d980c-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="d980c-138">No **Gerenciador de Soluções**, selecione o projeto Web.</span><span class="sxs-lookup"><span data-stu-id="d980c-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="d980c-139">Isso exibirá as propriedades para o projeto na janela **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="d980c-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="d980c-140">Na janela Propriedades, copie a **URL de SSL**.</span><span class="sxs-lookup"><span data-stu-id="d980c-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="d980c-141">No projeto de suplemento, abra o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="d980c-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="d980c-142">Certifique-se de que você está editando o XML do código-fonte.</span><span class="sxs-lookup"><span data-stu-id="d980c-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="d980c-143">Para alguns tipos de projeto o Visual Studio abrirá o modo de exibição de visualização do XML que não funcionará para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="d980c-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="d980c-144">Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou.</span><span class="sxs-lookup"><span data-stu-id="d980c-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="d980c-145">Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão muito similares a `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="d980c-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="d980c-146">Salve o arquivo XML.</span><span class="sxs-lookup"><span data-stu-id="d980c-146">Save the XML file.</span></span>
7. <span data-ttu-id="d980c-147">Clique com botão direito do mouse no projeto Web e escolha **Depurar** -> **Iniciar nova instância**.</span><span class="sxs-lookup"><span data-stu-id="d980c-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="d980c-148">Isso executará o projeto Web sem iniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="d980c-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="d980c-149">No Office na Web, realize o sideload do suplemento usando as etapas descritas anteriormente em [Sideload de um suplemento do Office no Office na Web](#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="d980c-149">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-on-the-web).</span></span>
