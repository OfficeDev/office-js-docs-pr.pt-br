---
title: Empacotar seu suplemento usando o Visual Studio para preparar a publicação
description: Como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2019.
ms.date: 10/14/2019
localization_priority: Priority
ms.openlocfilehash: 784741cffa0e3015caaa9c70fbb56f4b70df9462
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626961"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="2194e-103">Empacotar seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="2194e-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="2194e-104">Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="2194e-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="2194e-105">Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente.</span><span class="sxs-lookup"><span data-stu-id="2194e-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="2194e-106">Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="2194e-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="2194e-107">Para implantar seu projeto Web usando o Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="2194e-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="2194e-108">Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="2194e-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="2194e-109">Na guia **Compilar**, escolha **Publicar [Nome do seu suplemento]**.</span><span class="sxs-lookup"><span data-stu-id="2194e-109">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="2194e-110">Na janela **Escolha um destino de publicação**, escolha uma das opções de publicação para o seu destino preferido.</span><span class="sxs-lookup"><span data-stu-id="2194e-110">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="2194e-111">Cada destino de publicação exige que você inclua mais informações para começar, como um local de pasta ou uma Máquina Virtual do Azure.</span><span class="sxs-lookup"><span data-stu-id="2194e-111">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="2194e-112">Depois de especificar um local de publicação e preencher todas as informações necessárias, selecione **Publicar**</span><span class="sxs-lookup"><span data-stu-id="2194e-112">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="2194e-113">A escolha de um destino de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.</span><span class="sxs-lookup"><span data-stu-id="2194e-113">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="2194e-114">Para obter mais informações sobre as etapas de implantação de cada opção de destino de publicação, confira [Primeiro contato com a implantação no Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span><span class="sxs-lookup"><span data-stu-id="2194e-114">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="2194e-115">Para empacotar e publicar seu suplemento usando IIS, FTP ou implantação da Web usando o Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="2194e-115">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="2194e-116">Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="2194e-116">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="2194e-117">Na guia **Compilar**, escolha **Publicar [Nome do seu suplemento]**.</span><span class="sxs-lookup"><span data-stu-id="2194e-117">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="2194e-118">Na janela **Escolha um destino de publicação**, escolha **IIS, FTP, etc** e selecione **Configurar**.</span><span class="sxs-lookup"><span data-stu-id="2194e-118">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="2194e-119">Em seguida, selecione **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="2194e-119">Next, select **Publish**.</span></span>
3. <span data-ttu-id="2194e-120">Será exibido um assistente que o ajudará durante todo o processo.</span><span class="sxs-lookup"><span data-stu-id="2194e-120">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="2194e-121">Verifique se o método de publicação é o método preferido, como implantação da Web.</span><span class="sxs-lookup"><span data-stu-id="2194e-121">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="2194e-122">Na caixa **URL de destino**, digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e, em seguida, selecione **Avançar**.</span><span class="sxs-lookup"><span data-stu-id="2194e-122">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> <span data-ttu-id="2194e-123">Se você pretende enviar seu suplemento ao AppSource, escolha o botão **Validar conexão** para identificar problemas que possam impedir a aceitação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="2194e-123">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="2194e-124">Você deve resolver todos os problemas antes de enviar seu suplemento para a loja.</span><span class="sxs-lookup"><span data-stu-id="2194e-124">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="2194e-125">Confirme as configurações desejadas, incluindo **Opções de publicação de arquivo** e selecione **Salvar**.</span><span class="sxs-lookup"><span data-stu-id="2194e-125">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="2194e-126">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="2194e-126">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="2194e-p106">Agora é possível carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="2194e-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="2194e-130">Confira também</span><span class="sxs-lookup"><span data-stu-id="2194e-130">See also</span></span>

- [<span data-ttu-id="2194e-131">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="2194e-131">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="2194e-132">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="2194e-132">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
