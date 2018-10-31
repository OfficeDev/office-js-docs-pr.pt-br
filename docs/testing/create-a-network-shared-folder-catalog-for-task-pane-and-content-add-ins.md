---
title: Fazer sideload de suplementos do Office para teste
description: ''
ms.date: 10/17/2018
ms.openlocfilehash: 6ee8e4e9a2413b34cb8991b09d61e16888a0e6a6
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640019"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="f1327-102">Fazer sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="f1327-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="f1327-103">Você pode instalar um suplemento do Office para teste em um cliente do Office no Windows publicando o manifesto em um compartilhamento de arquivos na rede (instruções abaixo).</span><span class="sxs-lookup"><span data-stu-id="f1327-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="f1327-p101">Se seu projeto de suplemento foi criado com a [ferramenta **yo office**](https://github.com/OfficeDev/generator-office), há uma maneira alternativa de fazer sideload que pode servir para você. Para obter mais detalhes, consulte [Fazer sideload de suplementos do Office usando o comando de sideload](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="f1327-p101">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you. For details, see [Sideload Office Add-ins using the sideload command](sideload-office-addin-using-sideload-command.md).</span></span>

<span data-ttu-id="f1327-p102">Este artigo se aplica somente para testar suplementos do Word, PowerPoint ou Excel no Windows. Se você deseja testar em outra plataforma ou deseja testar um suplemento do Outlook, consulte um dos seguintes tópicos para fazer sideload de seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="f1327-p102">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows. If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="f1327-108">Fazer sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="f1327-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="f1327-109">Sideload dos suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="f1327-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="f1327-110">Fazer sideload de suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="f1327-110">Sideload Outlook add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="f1327-111">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento na área de trabalho do Office ou no Office Online usando um catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="f1327-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="f1327-112">Compartilhar uma pasta</span><span class="sxs-lookup"><span data-stu-id="f1327-112">Share a folder</span></span>

1. <span data-ttu-id="f1327-113">No Explorador de Arquivos no computador do Windows em que você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="f1327-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="f1327-114">Abra o menu de contexto para a pasta que você deseja usar como seu catálogo de pasta compartilhada (clique com o botão direito do mouse na pasta) e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="f1327-114">Open the context menu for the folder you want to use as your shared folder catalog (right-click the folder) and choose **Properties**.</span></span>

3. <span data-ttu-id="f1327-115">Dentro da janela de diálogo **Propriedades** , abra a guia **Compartilhamento** e escolha o botão **Compartilhar** .</span><span class="sxs-lookup"><span data-stu-id="f1327-115">Within the **Properties** dialog window, open the **Sharing** tab and then choose the **Share** button.</span></span>

    ![caixa de diálogo Propriedades da pasta com a guia Compartilhamento e o botão Compartilhar realçados](../images/sideload-windows-properties-dialog.png)

4. <span data-ttu-id="f1327-117">Dentro da janela de diálogo **Acesso à rede**, adicione a si mesmo e quaisquer outros usuários e/ou grupos com quem você deseja compartilhar seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="f1327-117">Within the **Network access** dialog window, add yourself and any other users and/or groups with whom you want to share your add-in.</span></span> <span data-ttu-id="f1327-118">Você precisará, no mínimo, de permissão de **Leitura/Gravação** para a pasta.</span><span class="sxs-lookup"><span data-stu-id="f1327-118">You will need at least **Read/Write** permission to the folder.</span></span> <span data-ttu-id="f1327-119">Depois de concluir a seleção de pessoas com as quais fazer o compartilhamento, escolha o botão **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="f1327-119">After you have finished choosing people to share with, choose the **Share** button.</span></span>

5. <span data-ttu-id="f1327-120">Quando você vir a confirmação de que **Sua pasta está compartilhada**, anote o caminho completo de rede que é exibido imediatamente após o nome da pasta.</span><span class="sxs-lookup"><span data-stu-id="f1327-120">When you see confirmation that **Your folder is shared**, make note of the full network path that's displayed immediately following the folder name.</span></span> <span data-ttu-id="f1327-121">(Você precisará digitar esse valor como a **Url do Catálogo** quando você [especificar a pasta compartilhada como um catálogo confiável](#specify-the-shared-folder-as-a-trusted-catalog), conforme descrito na próxima seção deste artigo.) Escolha o botão **Concluído** para fechar a janela de diálogo de **Acesso à rede**.</span><span class="sxs-lookup"><span data-stu-id="f1327-121">(You will need to enter this value as the **Catalog Url** when you [specify the shared folder as a trusted catalog](#specify-the-shared-folder-as-a-trusted-catalog), as described in the next section of this article.) Choose the **Done** button to close the **Network access** dialog window.</span></span>

   ![Caixa de diálogo de acesso à rede com o caminho de compartilhamento realçado](../images/sideload-windows-network-access-dialog.png)

6. <span data-ttu-id="f1327-123">Escolha o botão **Fechar** para fechar a janela de diálogo **Propriedades** .</span><span class="sxs-lookup"><span data-stu-id="f1327-123">Choose the **Close** button to close the **Workbook Connections** dialog box.</span></span>

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="f1327-124">Especificar a pasta compartilhada como um catálogo confiável</span><span class="sxs-lookup"><span data-stu-id="f1327-124">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="f1327-125">Abra um novo documento no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f1327-125">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="f1327-126">Escolha a guia **Arquivo** e escolha **Opções**.</span><span class="sxs-lookup"><span data-stu-id="f1327-126">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="f1327-127">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="f1327-127">Choose **Trust Center**, and then choose the **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="f1327-128">Escolha **Catálogos de Suplemento Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="f1327-128">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="f1327-129">Na caixa **Url do Catálogo** , insira o caminho completo de rede para a pasta que você [compartilhou](#share-a-folder) anteriormente.</span><span class="sxs-lookup"><span data-stu-id="f1327-129">In the **Catalog Url** box, enter the full network path to the folder that you [shared](#share-a-folder) previously.</span></span> <span data-ttu-id="f1327-130">Se você não conseguiu anotar o caminho de rede completo quando você compartilhou a pasta, você pode obtê-lo da janela de diálogo **Propriedades** da pasta, conforme mostrado na seguinte captura de tela.</span><span class="sxs-lookup"><span data-stu-id="f1327-130">If you failed to note the folder's full network path when you shared the folder, you can get it from the folder's **Properties** dialog window, as shown in the following screenshot.</span></span> 

    ![diálogo Propriedades da pasta com a guia Compartilhamento e o caminho de rede realçados](../images/sideload-windows-properties-dialog-2.png)
    
6. <span data-ttu-id="f1327-132">Depois de inserir o caminho de rede completo da pasta na caixa **Url do Catálogo**, escolha o botão **Adicionar catálogo**.</span><span class="sxs-lookup"><span data-stu-id="f1327-132">After you've entered the full network path of the folder into the **Catalog Url** box, choose the **Add catalog** button.</span></span>

7. <span data-ttu-id="f1327-133">Selecione a caixa de seleção **Mostrar no Menu** referente ao item recém-adicionado e escolha o botão **OK** para fechar a janela de diálogo **Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="f1327-133">Select the **Show in Menu** check box for the newly-added item, and then choose the **OK** button to close the **Trust Center** dialog window.</span></span> 

    ![Diálogo Central de Confiabilidade com o catálogo selecionado](../images/sideload-windows-trust-center-dialog.png)

8. <span data-ttu-id="f1327-135">Escolha o botão **OK** para fechar a janela de diálogo **Opções do Word**.</span><span class="sxs-lookup"><span data-stu-id="f1327-135">Choose the  **OK** button to close the **Internet Options** dialog box.</span></span>

9. <span data-ttu-id="f1327-136">Feche e reabra o aplicativo do Office para que as alterações tenham efeito.</span><span class="sxs-lookup"><span data-stu-id="f1327-136">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="f1327-137">Fazer o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="f1327-137">Sideload your add-in</span></span>


1. <span data-ttu-id="f1327-138">Coloque o arquivo XML de manifesto de qualquer suplemento que você está testando no catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="f1327-138">Put the manifest file of any add-in that you are testing in the shared folder catalog.</span></span> <span data-ttu-id="f1327-139">Observe que você implanta o próprio aplicativo Web em um servidor Web.</span><span class="sxs-lookup"><span data-stu-id="f1327-139">Note that you deploy the web application itself to a web server.</span></span> <span data-ttu-id="f1327-140">Certifique-se de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="f1327-140">Deploy the web application itself to a web server and specify the URL in the  **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="f1327-141">No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="f1327-141">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="f1327-142">Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="f1327-142">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="f1327-143">Selecione o nome do suplemento e escolha **OK** para inseri-lo.</span><span class="sxs-lookup"><span data-stu-id="f1327-143">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="f1327-144">Confira também</span><span class="sxs-lookup"><span data-stu-id="f1327-144">See also</span></span>

- [<span data-ttu-id="f1327-145">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="f1327-145">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="f1327-146">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="f1327-146">Publish your Office Add-in</span></span>](../publish/publish.md)
    
