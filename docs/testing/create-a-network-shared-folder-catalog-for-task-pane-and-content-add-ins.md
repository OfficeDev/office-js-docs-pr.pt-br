---
title: Fazer sideload de suplementos do Office para teste
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 94803a2c610fc869aefb6c77d53965981778e62e
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925119"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="275e0-102">Fazer sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="275e0-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="275e0-103">Você pode instalar um suplemento do Office para teste em um cliente do Office no Windows publicando o manifesto em um compartilhamento de arquivos na rede (instruções abaixo).</span><span class="sxs-lookup"><span data-stu-id="275e0-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="275e0-104">Se o seu projeto de suplemento foi criado com a ferramenta [**yo office**, existe](https://github.com/OfficeDev/generator-office) uma forma alternativa de sideload que pode funcionar para você.</span><span class="sxs-lookup"><span data-stu-id="275e0-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="275e0-105">Para mais detalhes, confira [Fazer sideload de Suplementos do Office usando o comando sideload](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="275e0-105">Sideload Office Add-ins using the sideload command</span></span>

<span data-ttu-id="275e0-106">Este artigo se aplica somente ao teste de suplementos do Word, Excel ou PowerPoint no Windows.</span><span class="sxs-lookup"><span data-stu-id="275e0-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="275e0-107">Se quiser fazer testes em outra plataforma ou se quiser testar um suplemento do Outlook, confira um dos tópicos a seguir para fazer o sideload do seu suplemento:</span><span class="sxs-lookup"><span data-stu-id="275e0-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="275e0-108">Sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="275e0-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="275e0-109">Sideload dos suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="275e0-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="275e0-110">Realizar sideload de Suplementos do Outlook para teste</span><span class="sxs-lookup"><span data-stu-id="275e0-110">Sideload Outlook Add-ins for testing</span></span>](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)


<span data-ttu-id="275e0-111">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento na área de trabalho do Office ou no Office Online usando um catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="275e0-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="275e0-112">Compartilhar uma pasta</span><span class="sxs-lookup"><span data-stu-id="275e0-112">Share a folder</span></span>

1. <span data-ttu-id="275e0-113">No computador do Windows, onde você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="275e0-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="275e0-114">Abra o menu de contexto para a pasta (com o botão direito) e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="275e0-114">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="275e0-115">Abra a guia **Compartilhamento**.</span><span class="sxs-lookup"><span data-stu-id="275e0-115">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="275e0-p103">Na página **Escolher pessoas...**, adicione a si mesmo e qualquer pessoa com quem você deseja compartilhar seu suplemento. Se todos eles forem membros de um grupo de segurança, você poderá adicionar o grupo. Você precisará de pelo menos permissão de **leitura/gravação** para a pasta.</span><span class="sxs-lookup"><span data-stu-id="275e0-p103">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="275e0-119">Escolha **Compartilhar** > **Concluído** > **Fechar**.</span><span class="sxs-lookup"><span data-stu-id="275e0-119">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="275e0-120">Especificar a pasta compartilhada como um catálogo confiável</span><span class="sxs-lookup"><span data-stu-id="275e0-120">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="275e0-121">Abra um novo documento no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="275e0-121">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="275e0-122">Escolha a guia **Arquivo** e escolha **Opções**.</span><span class="sxs-lookup"><span data-stu-id="275e0-122">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="275e0-123">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="275e0-123">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="275e0-124">Escolha **Catálogos de Suplemento Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="275e0-124">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="275e0-125">Na caixa  **URL de Catálogo**, digite o caminho de rede completo para o catálogo da pasta compartilhada e escolha **Adicionar Catálogo**.</span><span class="sxs-lookup"><span data-stu-id="275e0-125">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="275e0-126">Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="275e0-126">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="275e0-127">Feche o aplicativo do Office para que as alterações tenham efeito.</span><span class="sxs-lookup"><span data-stu-id="275e0-127">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="275e0-128">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="275e0-128">Sideload your add-in</span></span>


1. <span data-ttu-id="275e0-p104">Coloque o arquivo XML do manifesto de qualquer suplemento que você está testando no catálogo de pasta compartilhada. Observe que você implanta o próprio aplicativo Web em um servidor Web. Não deixe de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="275e0-p104">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="275e0-132">No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="275e0-132">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="275e0-133">Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="275e0-133">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="275e0-134">Selecione o nome do suplemento e escolha **OK** para inseri-lo.</span><span class="sxs-lookup"><span data-stu-id="275e0-134">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="275e0-135">Veja também</span><span class="sxs-lookup"><span data-stu-id="275e0-135">See also</span></span>

- [<span data-ttu-id="275e0-136">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="275e0-136">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="275e0-137">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="275e0-137">Publish your Office Add-in</span></span>](../publish/publish.md)
    
