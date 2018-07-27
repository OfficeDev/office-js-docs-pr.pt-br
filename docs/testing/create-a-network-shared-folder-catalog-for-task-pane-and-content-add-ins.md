---
title: Fazer sideload de suplementos do Office para teste
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e5769ef40868ec996194725d98913e61b76279bc
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/27/2018
ms.locfileid: "21270290"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="6d3b0-102">Fazer sideload de suplementos do Office para teste</span><span class="sxs-lookup"><span data-stu-id="6d3b0-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="6d3b0-103">Você pode instalar um suplemento do Office para teste em um cliente do Office em execução no Windows por meio de um dos seguintes métodos:</span><span class="sxs-lookup"><span data-stu-id="6d3b0-103">You can install an Office Add-in for testing in an Office client running on Windows by one of the following methods:</span></span>

- <span data-ttu-id="6d3b0-104">Usando um catálogo de pastas compartilhadas para publicar o manifesto em um compartilhamento de arquivos de rede (instruções abaixo)</span><span class="sxs-lookup"><span data-stu-id="6d3b0-104">Using a shared folder catalog to publish the manifest to a network file share (instructions below)</span></span>
- [<span data-ttu-id="6d3b0-105">Executando o comando "**npm run sideload**" da raiz da pasta do projeto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-105">Running the "**npm run sideload**" command from the root of the add-in project folder.</span></span>](sideload-office-addin-using-sideload-command.md) 
>[!NOTE]
><span data-ttu-id="6d3b0-106">O método "npm run sideload" funciona apenas para suplementos do Excel, Word e PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-106">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

<span data-ttu-id="6d3b0-107">Se não estiver testando um suplemento do Word, do Excel ou do PowerPoint no Windows, confira um dos tópicos a seguir para fazer sideload do suplemento:</span><span class="sxs-lookup"><span data-stu-id="6d3b0-107">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="6d3b0-108">Sideload de suplementos do Office para teste no Office Online</span><span class="sxs-lookup"><span data-stu-id="6d3b0-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="6d3b0-109">Sideload dos suplementos do Office para teste em um iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="6d3b0-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

<span data-ttu-id="6d3b0-110">O vídeo a seguir oferece orientações para o processo de sideload do seu suplemento no Office para área de trabalho ou no Office Online usando um catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="6d3b0-111">Compartilhar uma pasta</span><span class="sxs-lookup"><span data-stu-id="6d3b0-111">Share a folder</span></span>

1. <span data-ttu-id="6d3b0-112">No computador do Windows, onde você deseja hospedar seu suplemento, acesse a pasta pai ou letra da unidade da pasta que você deseja usar como seu catálogo de pasta compartilhada.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-112">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="6d3b0-113">Abra o menu de contexto para a pasta (com o botão direito) e escolha **Propriedades**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-113">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="6d3b0-114">Abra a guia **Compartilhamento**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-114">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="6d3b0-p101">Na página **Escolher pessoas...**, adicione a si mesmo e qualquer pessoa com quem você deseja compartilhar seu suplemento. Se todos eles forem membros de um grupo de segurança, você poderá adicionar o grupo. Você precisará de pelo menos permissão de **leitura/gravação** para a pasta.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-p101">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="6d3b0-118">Escolha **Compartilhar** > **Concluído** > **Fechar**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-118">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="6d3b0-119">Especificar a pasta compartilhada como um catálogo confiável</span><span class="sxs-lookup"><span data-stu-id="6d3b0-119">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="6d3b0-120">Abra um novo documento no Excel, no Word ou no PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-120">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="6d3b0-121">Escolha a guia **Arquivo** e escolha **Opções**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-121">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="6d3b0-122">Escolha **Central de Confiabilidade**, e escolha o botão **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-122">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="6d3b0-123">Escolha **Catálogos de Suplemento Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-123">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="6d3b0-124">Na caixa  **URL de Catálogo**, digite o caminho de rede completo para o catálogo da pasta compartilhada e escolha **Adicionar Catálogo**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-124">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="6d3b0-125">Selecione a caixa de seleção **Mostrar no Menu** e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-125">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="6d3b0-126">Feche o aplicativo do Office para que as alterações tenham efeito.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-126">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="6d3b0-127">Realizar o sideload do seu suplemento</span><span class="sxs-lookup"><span data-stu-id="6d3b0-127">Sideload your add-in</span></span>

1. <span data-ttu-id="6d3b0-p102">Coloque o arquivo de manifesto de qualquer suplemento que você está testando no catálogo de pasta compartilhada. Observe que você implanta o próprio aplicativo Web em um servidor Web. Não deixe de especificar a URL no elemento **SourceLocation** do arquivo de manifesto.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-p102">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="6d3b0-131">No Excel, Word ou PowerPoint, selecione **Meus Suplementos** na guia **Inserir** da faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-131">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="6d3b0-132">Escolha **PASTA COMPARTILHADA** na parte superior da caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-132">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="6d3b0-133">Selecione o nome do suplemento e escolha **OK** para inseri-lo.</span><span class="sxs-lookup"><span data-stu-id="6d3b0-133">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="6d3b0-134">Veja também</span><span class="sxs-lookup"><span data-stu-id="6d3b0-134">See also</span></span>

- [<span data-ttu-id="6d3b0-135">Validar e solucionar problemas com seu manifesto</span><span class="sxs-lookup"><span data-stu-id="6d3b0-135">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="6d3b0-136">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="6d3b0-136">Publish your Office Add-in</span></span>](../publish/publish.md)
    
