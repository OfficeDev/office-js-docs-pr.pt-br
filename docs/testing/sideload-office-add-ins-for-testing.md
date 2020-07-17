---
title: Realizar sideload de suplementos do Office no Office na Web para teste
description: Testar o suplemento do Office no Office na web através de sideloading
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 60b7e4f1d598e4f5ec09307d58294f54123112ad
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094117"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a><span data-ttu-id="d49c8-103">Realizar sideload de suplementos do Office no Office na Web para teste</span><span class="sxs-lookup"><span data-stu-id="d49c8-103">Sideload Office Add-ins in Office on the web for testing</span></span>

<span data-ttu-id="d49c8-104">Você pode instalar um suplemento do Office para teste usando sideloading, sem precisar primeiro colocá-lo em um catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="d49c8-104">You can install an Office Add-in for testing without having to first put it in an add-in catalog by using sideloading.</span></span> <span data-ttu-id="d49c8-105">O Sideload pode ser feito no Microsoft 365 ou no Office na Web.</span><span class="sxs-lookup"><span data-stu-id="d49c8-105">Sideloading can be done in either Microsoft 365 or Office on the web.</span></span> <span data-ttu-id="d49c8-106">O procedimento é ligeiramente diferente nas duas plataformas.</span><span class="sxs-lookup"><span data-stu-id="d49c8-106">The procedure is slightly different for the two platforms.</span></span>

<span data-ttu-id="d49c8-107">Quando você realiza o sideload de um suplemento, o manifesto do suplemento é armazenado localmente do navegador e, portanto, se você limpar o cache do navegador ou alternar para um navegador diferente, precisará realizar o sideload do suplemento novamente.</span><span class="sxs-lookup"><span data-stu-id="d49c8-107">When you sideload an add-in, the add-in manifest is stored in the browser's local storage, so if you clear the browser's cache, or switch to a different browser, you have to sideload the add-in again.</span></span>

> [!NOTE]
> <span data-ttu-id="d49c8-p102">A realização do sideload como descrito neste artigo tem suporte no Word, no Excel e no PowerPoint. Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="d49c8-p102">Sideloading as described in this article is supported on Word, Excel, and PowerPoint. To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

<span data-ttu-id="d49c8-110">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento no Office na Web ou para área de trabalho.</span><span class="sxs-lookup"><span data-stu-id="d49c8-110">The following video walks you through the process of sideloading your add-in in Office on the web or desktop.</span></span>

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a><span data-ttu-id="d49c8-111">Realizar sideload de um suplemento do Office no Office na Web</span><span class="sxs-lookup"><span data-stu-id="d49c8-111">Sideload an Office Add-in in Office on the web</span></span>

1. <span data-ttu-id="d49c8-112">Abra o [Microsoft Office na Web](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="d49c8-112">Open [Microsoft Office on the web](https://office.live.com/).</span></span>

2. <span data-ttu-id="d49c8-113">Em introdução **aos aplicativos online agora**, escolha **Excel**, **Word**ou **PowerPoint**; e, em seguida, abra um novo documento.</span><span class="sxs-lookup"><span data-stu-id="d49c8-113">In **Get started with the online apps now**, choose **Excel**, **Word**, or **PowerPoint**; and then open a new document.</span></span>

3. <span data-ttu-id="d49c8-114">Abra a guia **Inserir** na faixa de opções e, na seção **suplementos** , escolha **suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-114">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>

4. <span data-ttu-id="d49c8-115">Na caixa de diálogo **suplementos do Office** , selecione a guia **meus suplementos** , escolha **gerenciar meus suplementos**e, em seguida, **carregar meu suplemento**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-115">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>

    ![A caixa de diálogo Suplementos do Office com um menu suspenso "Gerenciar meus suplementos" no canto superior direito e abaixo o menu suspenso com a opção "Carregar meu suplemento"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="d49c8-117">**Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-117">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>

    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

6. <span data-ttu-id="d49c8-p103">Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.</span><span class="sxs-lookup"><span data-stu-id="d49c8-p103">Verify that your add-in is installed. For example, if it is an add-in command, it should appear on either the ribbon or the context menu. If it is a task pane add-in, the pane should appear.</span></span>

> [!NOTE]
><span data-ttu-id="d49c8-122">Para testar o suplemento do Office com o Microsoft Edge, são necessárias duas etapas de configuração:</span><span class="sxs-lookup"><span data-stu-id="d49c8-122">To test your Office Add-in with Microsoft Edge, two configuration steps are required:</span></span> 
>
> - <span data-ttu-id="d49c8-123">Em um prompt de comando do Windows, execute a seguinte linha: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span><span class="sxs-lookup"><span data-stu-id="d49c8-123">In a Windows Command Prompt, run the following line: `CheckNetIsolation LoopbackExempt -a -n="Microsoft.MicrosoftEdge_8wekyb3d8bbwe"`</span></span>
>
> - <span data-ttu-id="d49c8-124">Digite "**about: flags**" na barra de pesquisa do Microsoft Edge para exibir as opções de configurações do desenvolvedor.</span><span class="sxs-lookup"><span data-stu-id="d49c8-124">Enter "**about:flags**" in the Microsoft Edge search bar to bring up the Developer Settings options.</span></span>  <span data-ttu-id="d49c8-125">Marque a opção "**permitir auto-retorno de localhost**" e reinicie o Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="d49c8-125">Check the "**Allow localhost loopback**" option and restart Microsoft Edge.</span></span>

>    ![A opção “Permitir loopback do localhost” do Microsoft Edge com a caixa marcada.](../images/allow-localhost-loopback.png)

## <a name="sideload-an-office-add-in-in-office-365"></a><span data-ttu-id="d49c8-127">Realizar sideload de um suplemento do Office no Office 365</span><span class="sxs-lookup"><span data-stu-id="d49c8-127">Sideload an Office Add-in in Office 365</span></span>

1. <span data-ttu-id="d49c8-128">Entre em sua conta do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="d49c8-128">Sign in to your Microsoft 365 account.</span></span>

2. <span data-ttu-id="d49c8-129">Abra o inicializador de aplicativos na extremidade esquerda da barra de ferramentas e selecione **Excel**, **Word**ou **PowerPoint**e, em seguida, crie um novo documento.</span><span class="sxs-lookup"><span data-stu-id="d49c8-129">Open the App Launcher on the left end of the toolbar and select **Excel**, **Word**, or **PowerPoint**, and then create a new document.</span></span>

3. <span data-ttu-id="d49c8-130">As etapas 3 a 6 são as mesmas da seção anterior **Realize sideload para um suplemento do Office no Office na Web**. </span><span class="sxs-lookup"><span data-stu-id="d49c8-130">Steps 3 - 6 are the same as in the preceding section **Sideload an Office Add-in in Office on the web**.</span></span>

## <a name="sideload-an-add-in-when-using-visual-studio"></a><span data-ttu-id="d49c8-131">Sideload de um suplemento usando o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d49c8-131">Sideload an add-in when using Visual Studio</span></span>

<span data-ttu-id="d49c8-132">Se estiver usando o Visual Studio para desenvolver o suplemento, o processo de sideload é semelhante.</span><span class="sxs-lookup"><span data-stu-id="d49c8-132">If you're using Visual Studio to develop your add-in, the process to sideload is similar.</span></span> <span data-ttu-id="d49c8-133">A única diferença é que você deve atualizar o valor do elemento **SourceURL** no manifesto para incluir a URL completa em que o suplemento for implantado.</span><span class="sxs-lookup"><span data-stu-id="d49c8-133">The only difference is that you must update the value of the **SourceURL** element in your manifest to include the full URL where the add-in is deployed.</span></span>

> [!NOTE]
> <span data-ttu-id="d49c8-134">Embora você possa realizar o sideload de suplementos do Visual Studio para o Office na Web, não é possível depurá-los no Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="d49c8-134">Although you can sideload add-ins from Visual Studio to Office on the web, you cannot debug them from Visual Studio.</span></span> <span data-ttu-id="d49c8-135">Para depurar você precisará usar as ferramentas de depuração do navegador.</span><span class="sxs-lookup"><span data-stu-id="d49c8-135">To debug you will need to use the browser debugging tools.</span></span> <span data-ttu-id="d49c8-136">Para saber mais, confira [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md).</span><span class="sxs-lookup"><span data-stu-id="d49c8-136">For more information, see [Debug add-ins in Office on the web](debug-add-ins-in-office-online.md).</span></span>

1. <span data-ttu-id="d49c8-137">No Visual Studio, abra a janela **Propriedades** escolhendo **Modo de exibição** -> **Janela de propriedades**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-137">In Visual Studio, show the **Properties** window by choosing **View** -> **Properties Window**.</span></span>
2. <span data-ttu-id="d49c8-138">No **Gerenciador de Soluções**, selecione o projeto Web.</span><span class="sxs-lookup"><span data-stu-id="d49c8-138">In the **Solution Explorer**, select the web project.</span></span> <span data-ttu-id="d49c8-139">Isso exibirá as propriedades para o projeto na janela **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-139">This will display properties for the project in the **Properties** window.</span></span>
3. <span data-ttu-id="d49c8-140">Na janela Propriedades, copie a **URL de SSL**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-140">In the Properties window, copy the **SSL URL**.</span></span>
4. <span data-ttu-id="d49c8-141">No projeto de suplemento, abra o arquivo XML do manifesto.</span><span class="sxs-lookup"><span data-stu-id="d49c8-141">In the add-in project, open the manifest XML file.</span></span> <span data-ttu-id="d49c8-142">Certifique-se de que você está editando o XML do código-fonte.</span><span class="sxs-lookup"><span data-stu-id="d49c8-142">Be sure you are editing the source XML.</span></span> <span data-ttu-id="d49c8-143">Para alguns tipos de projeto o Visual Studio abrirá o modo de exibição de visualização do XML que não funcionará para a próxima etapa.</span><span class="sxs-lookup"><span data-stu-id="d49c8-143">For some project types Visual Studio will open a visual view of the XML which will not work for the next step.</span></span>
5. <span data-ttu-id="d49c8-144">Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou.</span><span class="sxs-lookup"><span data-stu-id="d49c8-144">Search and replace all instances of **~remoteAppUrl/** with the SSL URL you just copied.</span></span> <span data-ttu-id="d49c8-145">Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão muito similares a `https://localhost:44300/Home.html`.</span><span class="sxs-lookup"><span data-stu-id="d49c8-145">You will see several replacements depending on the project type, and the new URLs will appear similar to `https://localhost:44300/Home.html`.</span></span>
6. <span data-ttu-id="d49c8-146">Salve o arquivo XML.</span><span class="sxs-lookup"><span data-stu-id="d49c8-146">Save the XML file.</span></span>
7. <span data-ttu-id="d49c8-147">Clique com botão direito do mouse no projeto Web e escolha **Depurar** -> **Iniciar nova instância**.</span><span class="sxs-lookup"><span data-stu-id="d49c8-147">Right click the web project and choose **Debug** -> **Start new instance**.</span></span> <span data-ttu-id="d49c8-148">Isso executará o projeto Web sem iniciar o Office.</span><span class="sxs-lookup"><span data-stu-id="d49c8-148">This will run the web project without launching Office.</span></span>
8. <span data-ttu-id="d49c8-149">No Office na Web, realize o sideload do suplemento usando as etapas descritas anteriormente em [Sideload de um suplemento do Office no Office na Web](#sideload-an-office-add-in-in-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="d49c8-149">From Office on the web, sideload the add-in using steps previously described in [Sideload an Office Add-in in Office on the web](#sideload-an-office-add-in-in-office-on-the-web).</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="d49c8-150">Remover um suplemento do suplementos foi feito</span><span class="sxs-lookup"><span data-stu-id="d49c8-150">Remove a sideloaded add-in</span></span>

<span data-ttu-id="d49c8-151">Você pode remover um suplemento suplementos foi feito anteriormente limpando o cache do navegador.</span><span class="sxs-lookup"><span data-stu-id="d49c8-151">You can remove a previously sideloaded add-in by clearing your browser's cache.</span></span> <span data-ttu-id="d49c8-152">Além disso, se você fizer alterações no manifesto do suplemento (por exemplo, atualizar nomes de arquivo de ícones ou texto de comandos de suplemento), talvez seja necessário limpar o cache e, em seguida, resideload o suplemento usando o manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="d49c8-152">Additionally, if you make changes to your add-in's manifest (for example, update file names of icons or text of add-in commands), you may need to clear the cache and then re-sideload the add-in using updated manifest.</span></span> <span data-ttu-id="d49c8-153">Isso permitirá que o Office processe o suplemento conforme descrito no manifesto atualizado.</span><span class="sxs-lookup"><span data-stu-id="d49c8-153">Doing so will allow Office to render the add-in as it's described by the updated manifest.</span></span>
