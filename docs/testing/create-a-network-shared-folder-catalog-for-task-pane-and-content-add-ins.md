---
title: Realizar sideload de suplementos do Office para teste
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 79d1bfc9332208e59e750e94a14abd6f1192ebe6
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871581"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="90bce-102">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="90bce-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="90bce-103">Você pode instalar um suplemento do Office para testá-lo em um cliente do Office em execução no Windows usando um catálogo de pasta compartilhada para publicar o manifesto em um compartilhamento de arquivos de rede.</span><span class="sxs-lookup"><span data-stu-id="90bce-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="90bce-104">Se o seu projeto de suplemento tiver sido criado com a ferramenta [ **yo office**](https://github.com/OfficeDev/generator-office), há uma maneira alternativa de realizar o sideloading que pode funcionar para você.</span><span class="sxs-lookup"><span data-stu-id="90bce-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="90bce-105">Para mais detalhes, veja [Realizar Sideload de Suplementos do Office](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="90bce-105">For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="90bce-106">Este artigo se aplica somente para testes em suplementos Word, Excel ou PowerPoint no Windows.</span><span class="sxs-lookup"><span data-stu-id="90bce-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="90bce-107">Se você deseja testar em outra plataforma ou um suplemento do Outlook, veja os tópicos seguintes para realizar o sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="90bce-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="90bce-108">Realizar sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="90bce-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="90bce-109">Sideload suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="90bce-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="90bce-110">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="90bce-110">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="90bce-111">O vídeo a seguir oferece orientações para a realização do processo de sideload no suplemento do Office para área de trabalho ou Office Online.</span><span class="sxs-lookup"><span data-stu-id="90bce-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online using a shared folder catalog.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="90bce-112">Compartilhar uma pasta</span><span class="sxs-lookup"><span data-stu-id="90bce-112">Share a folder</span></span>

1. <span data-ttu-id="90bce-113">No computador do Windows, onde você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="90bce-113">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="90bce-114">Abra o menu de contexto na pasta que você deseja usar como catálogo de pasta compartilhada (clique com o botão direito) e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="90bce-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="90bce-115">Dentro da janela de diálogo **Propriedades** abra a guia **Compartilhamento**e escolha o botão **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="90bce-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![caixa de diálogo de Propriedades de pastas com o guia de compartilhamento e o botão Compartilhamento realçado](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="90bce-117">Dentro a janela de diálogo **Acesso à rede** adicione você mesmo e quaisquer outros usuários e/ou grupos com quem você deseja compartilhar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="90bce-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="90bce-118">Você precisará de pelo menos da permissão **Leitura/Gravação** para a pasta.</span><span class="sxs-lookup"><span data-stu-id="90bce-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="90bce-119">Quando terminar de escolher as pessoas para compartilhar, escolha o botão **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="90bce-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="90bce-120">Quando você vir a confirmação **Sua pasta foi compartilhada**, anote o caminho de rede completo que é exibido imediatamente após o nome da pasta.</span><span class="sxs-lookup"><span data-stu-id="90bce-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="90bce-121">(Você precisará inserir esse valor como o **Url Catálogo** quando você [especificar a pasta compartilhada como um catálogo confiável](#specify-the-shared-folder-as-a-trusted-catalog), conforme descrito na próxima seção deste artigo.) Escolha o botão **Concluído** para fechar a janela de diálogo **Acesso à rede**.</span><span class="sxs-lookup"><span data-stu-id="90bce-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Caixa de diálogo de acesso de rede com o caminho do compartilhamento realçado](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="90bce-123">Escolha o botão **Fechar** para fechar a caixa de diálogo **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="90bce-123">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="90bce-124">Especifique a pasta compartilhada como um catálogo confiável</span><span class="sxs-lookup"><span data-stu-id="90bce-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="90bce-125">Abra um novo documento no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="90bce-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="90bce-126">Escolha a guia **Arquivo** e, então, **Opções**.</span><span class="sxs-lookup"><span data-stu-id="90bce-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="90bce-127">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="90bce-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="90bce-128">Escolha **Catálogos de Suplemento Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="90bce-128">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="90bce-129">Na caixa**Url catálogo**, digite o caminho completo da rede para a pasta que você [compartilhou](#share-a-folder) anteriormente.</span><span class="sxs-lookup"><span data-stu-id="90bce-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="90bce-130">Se você não conseguiu anotar todo o caminho de rede da pasta ao compartilhar a pasta, você pode obtê-lo na janela de diálogo **Propriedades**, conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="90bce-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![caixa de diálogo de Propriedades de pastas com o guia de compartilhamento e o caminho de rede realçado](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="90bce-132">Depois de inserir o caminho de de rede completo da pasta na caixa **Url catálogo**, escolha o botão **Adicionar Catálogo**.</span><span class="sxs-lookup"><span data-stu-id="90bce-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="90bce-133">Selecione a caixa de seleção **Mostrar no Menu** no novo item adicionado e, em seguida, escolha o botão **Ok** para fechar a janela de diálogo **Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="90bce-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Caixa de diálogo Central de confiabilidade com catálogo selecionado](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="90bce-135">Escolha o botão **OK** para fechar a janela de diálogo **Opções do Word**.</span><span class="sxs-lookup"><span data-stu-id="90bce-135">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="90bce-136">Feche e abra novamente o aplicativo do Office para que as alterações tenham efeito.</span><span class="sxs-lookup"><span data-stu-id="90bce-136">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="90bce-137">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="90bce-137">Sideload your add-in</span></span>


1. <span data-ttu-id="90bce-138">Coloque o arquivo de manifesto XML de qualquer suplemento que você esteja testando no catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="90bce-138">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="90bce-139">Observe que você implanta o próprio aplicativo Web em um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="90bce-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="90bce-140">Não deixe de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="90bce-140">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="90bce-141">No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="90bce-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="90bce-142">Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="90bce-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="90bce-143">Selecione o nome do suplemento e escolha **OK** para inseri-lo.</span><span class="sxs-lookup"><span data-stu-id="90bce-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="90bce-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="90bce-144">See also</span></span>

- [<span data-ttu-id="90bce-145">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="90bce-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="90bce-146">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="90bce-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
