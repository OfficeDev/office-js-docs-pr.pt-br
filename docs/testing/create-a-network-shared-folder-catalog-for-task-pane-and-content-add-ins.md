---
title: Realizar sideload de suplementos do Office para teste
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: 19cd599ea743fc577a5139d3f278dd3f993ec5b1
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477926"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="1ea0f-102">Realizar sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="1ea0f-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="1ea0f-103">Você pode instalar um suplemento do Office para testá-lo em um cliente do Office em execução no Windows usando um catálogo de pasta compartilhada para publicar o manifesto em um compartilhamento de arquivos de rede.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-103">You can install an Office Add-in for testing in an Office client running on Windows by publishing the manifest to a network file share (instructions below).</span></span>

> [!NOTE]
> <span data-ttu-id="1ea0f-104">Se o projeto de suplemento tiver sido criado com uma versão suficientemente recente do [Gerador Yeoman para Suplementos do Office](https://github.com/OfficeDev/generator-office), o suplemento realizará sideload automaticamente no cliente de desktop do Office ao executar o `npm start`.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-104">If your add-in project was created with a sufficiently recent version of the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), the add-in will automatically sideload in the Office desktop client when you run `npm start`.</span></span>

<span data-ttu-id="1ea0f-105">Este artigo se aplica somente para testes de suplementos do Word, Excel, PowerPoint e Project no Windows.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-105">This article applies only to testing Word, Excel, PowerPoint, and Project add-ins on Windows.</span></span> <span data-ttu-id="1ea0f-106">Se você deseja testar em outra plataforma ou um suplemento do Outlook, veja os tópicos seguintes para realizar o sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="1ea0f-106">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="1ea0f-107">Realizar sideload de suplementos do Office no Office na Web para teste</span><span class="sxs-lookup"><span data-stu-id="1ea0f-107">Sideload Office Add-ins in Office on the web for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="1ea0f-108">Sideload suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="1ea0f-108">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="1ea0f-109">Realizar sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="1ea0f-109">Sideload Outlook add-ins for testing</span></span>](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

<span data-ttu-id="1ea0f-110">O vídeo a seguir oferece orientações para a realização do processo de sideload no suplemento do Office na Web ou para área de trabalho usando um catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online using a shared folder catalog.</span></span>  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a><span data-ttu-id="1ea0f-111">Compartilhar uma pasta</span><span class="sxs-lookup"><span data-stu-id="1ea0f-111">Share a folder</span></span>

1. <span data-ttu-id="1ea0f-112">No computador do Windows, onde você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-112">In File Explorer on the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="1ea0f-113">Abra o menu de contexto na pasta que você deseja usar como catálogo de pasta compartilhada (clique com o botão direito) e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-113">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="1ea0f-114">Dentro da janela de diálogo **Propriedades** abra a guia **Compartilhamento**e escolha o botão **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-114">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![caixa de diálogo de Propriedades de pastas com o guia de compartilhamento e o botão Compartilhamento realçado](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="1ea0f-116">Dentro a janela de diálogo **Acesso à rede** adicione você mesmo e quaisquer outros usuários e/ou grupos com quem você deseja compartilhar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-116">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="1ea0f-117">Você precisará de pelo menos da permissão **Leitura/Gravação** para a pasta.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-117">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="1ea0f-118">Quando terminar de escolher as pessoas para compartilhar, escolha o botão **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-118">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="1ea0f-119">Quando você vir a confirmação **Sua pasta foi compartilhada**, anote o caminho de rede completo que é exibido imediatamente após o nome da pasta.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-119">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="1ea0f-120">(Você precisará inserir esse valor como o **Url Catálogo** quando você [especificar a pasta compartilhada como um catálogo confiável](#specify-the-shared-folder-as-a-trusted-catalog), conforme descrito na próxima seção deste artigo.) Escolha o botão **Concluído** para fechar a janela de diálogo **Acesso à rede**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-120">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Caixa de diálogo de acesso de rede com o caminho do compartilhamento realçado](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="1ea0f-122">Escolha o botão **Fechar** para fechar a caixa de diálogo **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-122">Choose the **Close** button to close the **Properties** dialog window.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="1ea0f-123">Especifique a pasta compartilhada como um catálogo confiável</span><span class="sxs-lookup"><span data-stu-id="1ea0f-123">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="1ea0f-124">Abra um novo documento no Excel, no Word, no PowerPoint ou no Project.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-124">Open a new document in Excel, Word, PowerPoint, or Project.</span></span>
    
2. <span data-ttu-id="1ea0f-125">Escolha a guia **Arquivo** e, então, **Opções**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-125">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="1ea0f-126">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-126">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="1ea0f-127">Escolha **Catálogos de Suplemento Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-127">Choose **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="1ea0f-128">Na caixa**Url catálogo**, digite o caminho completo da rede para a pasta que você [compartilhou](#share-a-folder) anteriormente.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-128">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="1ea0f-129">Se você não conseguiu anotar todo o caminho de rede da pasta ao compartilhar a pasta, você pode obtê-lo na janela de diálogo **Propriedades**, conforme mostrado na captura de tela a seguir.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-129">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![caixa de diálogo de Propriedades de pastas com o guia de compartilhamento e o caminho de rede realçado](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="1ea0f-131">Depois de inserir o caminho de de rede completo da pasta na caixa **Url catálogo**, escolha o botão **Adicionar Catálogo**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-131">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="1ea0f-132">Selecione a caixa de seleção **Mostrar no Menu** no novo item adicionado e, em seguida, escolha o botão **Ok** para fechar a janela de diálogo **Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-132">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Caixa de diálogo Central de confiabilidade com catálogo selecionado](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="1ea0f-134">Escolha o botão **OK** para fechar a janela de diálogo **Opções do Word**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-134">Choose the **OK** button to close the **Word Options** dialog window.</span></span>

9. <span data-ttu-id="1ea0f-135">Feche e abra novamente o aplicativo do Office para que as alterações tenham efeito.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-135">Close and reopen the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="1ea0f-136">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="1ea0f-136">Sideload your add-in</span></span>


1. <span data-ttu-id="1ea0f-137">Coloque o arquivo de manifesto XML de qualquer suplemento que você esteja testando no catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-137">Put the manifest XML file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="1ea0f-138">Observe que você implanta o próprio aplicativo Web em um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-138">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="1ea0f-139">Não deixe de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-139">Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="1ea0f-140">No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-140">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span> <span data-ttu-id="1ea0f-141">No Project, selecione **Meus Suplementos** na guia **Projeto** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-141">In Project, select **My Add-ins** on the **Project** tab of the ribbon.</span></span> 

3. <span data-ttu-id="1ea0f-142">Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="1ea0f-143">Selecione o nome do suplemento e escolha **Adicionar** para inseri-lo.</span><span class="sxs-lookup"><span data-stu-id="1ea0f-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="1ea0f-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="1ea0f-144">See also</span></span>

- [<span data-ttu-id="1ea0f-145">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="1ea0f-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="1ea0f-146">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="1ea0f-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
