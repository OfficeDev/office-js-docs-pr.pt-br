---
title: Realizar sideload de suplementos do Office no Office Online para teste
description: Testar o suplemento do Office no Office Online através de sideloading
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 8870e955ca30c4a3b35f2b51e0e16a3ee634960d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871714"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a><span data-ttu-id="3d3e0-103">Realizar sideload de suplementos do Office no Office Online para teste</span><span class="sxs-lookup"><span data-stu-id="3d3e0-103">Sideload Office Add-ins in Office Online for testing</span></span>

<span data-ttu-id="3d3e0-104">Você pode instalar um suplemento do Office para teste usando sideloading, sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="3d3e0-105">O sideloading pode ser feito no Office 365 ou no Office Online.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-105">Sideloading can be done in either Office 365 or Office Online.</span></span> <span data-ttu-id="3d3e0-106">O procedimento é ligeiramente diferente nas duas plataformas.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-106">The procedure is slightly different for the two platforms.</span></span> 

<span data-ttu-id="3d3e0-107">Quando você realiza o sideload de um suplemento, o manifesto do suplemento é armazenado localmente do navegador e, portanto, se você limpar o cache do navegador ou alternar para um navegador diferente, precisará realizar o sideload do suplemento novamente.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>


> [!NOTE]
> <span data-ttu-id="3d3e0-p102">A realização do sideload como descrito neste artigo tem suporte no Word, no Excel e no PowerPoint. Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="3d3e0-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

<span data-ttu-id="3d3e0-110">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento no Office para área de trabalho ou no Office Online.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-110">The following video walks you through the process of sideloading your add-in in Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="3d3e0-111">Realizar sideload de um suplemento do Office no Office 365</span><span class="sxs-lookup"><span data-stu-id="3d3e0-111">Sideload an Office Add-in in Office 365</span></span>


1. <span data-ttu-id="3d3e0-112">Entre em sua conta do Office 365.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-112">Sign in to your Office 365 account.</span></span>
    
2. <span data-ttu-id="3d3e0-113">Abra o inicializador de aplicativos à esquerda da barra de ferramentas, selecione  **Excel**, **Word** ou **PowerPoint** e crie um novo documento.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-113">Open the App Launcher on the left end of the toolbar and select  **Excel**,  **Word**, or  **PowerPoint**, and then create a new document.</span></span>
    
3. <span data-ttu-id="3d3e0-114">Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-114">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="3d3e0-115">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MINHA ORGANIZAÇÃO** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-115">On the  **Office Add-ins** dialog, select the **MY ORGANIZATION** tab, and then **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplemento do Office tem o link  "Carregar Meu Suplemento" perto do canto superior esquerdo.](../images/office-add-ins.png)

5.  <span data-ttu-id="3d3e0-117">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

6. <span data-ttu-id="3d3e0-p103">Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in the pane should appear.</span></span>
    

## <a name="sideload-an-office-add-in-in-office-online"></a><span data-ttu-id="3d3e0-122">Realizar sideload de um suplemento do Office no Office Online</span><span class="sxs-lookup"><span data-stu-id="3d3e0-122">Sideload an Office Add-in in Office Online</span></span>


1. <span data-ttu-id="3d3e0-123">Abra o [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="3d3e0-123">Open [Microsoft Office Online](https://office.live.com/).</span></span>
    
2. <span data-ttu-id="3d3e0-124">Em **Comece a usar os aplicativos online agora**, escolha **Excel**, **Word** ou **PowerPoint** e abra um novo documento.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-124">In  **Get started with the online apps now**, choose  **Excel**,  **Word**, or  **PowerPoint**; and then open a new document.</span></span>
    
3. <span data-ttu-id="3d3e0-125">Abra a guia **Inserir** na faixa de opções e, na seção **Suplementos**, escolha **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-125">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
    
4. <span data-ttu-id="3d3e0-126">Na caixa de diálogo **Suplementos do Office**, selecione a guia **MEUS SUPLEMENTOS**, escolha **Gerenciar Meus Suplementos** e **Carregar Meu Suplemento**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-126">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="3d3e0-128">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-128">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

6. <span data-ttu-id="3d3e0-p104">Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-p104">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="3d3e0-133">Para testar o suplemento do Office com o Microsoft Edge, são necessárias duas etapas de configuração:</span><span class="sxs-lookup"><span data-stu-id="3d3e0-133">To test your Office Add-in with Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="3d3e0-134">Em um prompt de comando do Windows, execute a seguinte linha: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="3d3e0-134">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="3d3e0-135">Digite “**sobre:sinalizadores**” na barra de pesquisa do Edge para exibir as opções de Configurações do Desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-135">Enter “**about:flags**” in the Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="3d3e0-136">Verifique a opção “**Permitir loopback do localhost**” e reinicie o Edge.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-136">Check the “**Allow localhost loopback**” option and restart Edge.</span></span>

>    ![A opção “Permitir loopback do localhost” do Edge com a caixa marcada.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="3d3e0-138">Sideload de um suplemento usando o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="3d3e0-138">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="3d3e0-139">Se estiver usando o Visual Studio para desenvolver o suplemento, o processo de sideload é semelhante.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-139">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="3d3e0-140">A única diferença é que você deve atualizar o valor do elemento **SourceURL** no manifesto para incluir a URL completa em que o suplemento for implantado.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-140">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="3d3e0-141">Embora você possa fazer o sideload de suplementos do Visual Studio para o Office Online, não é possível depurá-los no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-141">Although you can sideload add-ins from Visual Studio to Office Online, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="3d3e0-142">Para depurar você precisará usar as ferramentas de depuração do navegador.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-142">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="3d3e0-143">Para saber mais, confira [Depurar suplementos no Office Online](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="3d3e0-143">For more information, see [Debug add-ins in Office Online](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="3d3e0-144">No Visual Studio, abra a janela **Propriedades** escolhendo **Modo de exibição** -> **Janela de propriedades**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-144">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="3d3e0-145">No **Gerenciador de Soluções**, selecione o projeto Web.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-145">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="3d3e0-146">Isso exibirá as propriedades para o projeto na janela **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-146">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="3d3e0-147">Na janela Propriedades, copie a **URL de SSL**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-147">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="3d3e0-148">No projeto de suplemento, abra o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-148">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="3d3e0-149">Certifique-se de que você está editando o XML do código-fonte.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-149">Be sure you are editing the source XML.</span></span> <span data-ttu-id="3d3e0-150">Para alguns tipos de projeto o Visual Studio abrirá o modo de exibição de visualização do XML que não funcionará para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-150">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="3d3e0-151">Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-151">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="3d3e0-152">Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão muito similares a `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-152">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="3d3e0-153">Salve o arquivo XML.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-153">Save the XML file.</span></span>
7. <span data-ttu-id="3d3e0-154">Clique com botão direito do mouse no projeto Web e escolha **Depurar** -> **Iniciar nova instância**.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-154">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="3d3e0-155">Isso executará o projeto Web sem iniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="3d3e0-155">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="3d3e0-156">No Office Online, faça o sideload do suplemento usando as etapas descritas anteriormente em [Sideload de um suplemento do Office no Office Online](#sideload-an-office-add-in-in-office-online).</span><span class="sxs-lookup"><span data-stu-id="3d3e0-156">From Office Online, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office Online](#sideload-an-office-add-in-in-office-online).</span></span>
