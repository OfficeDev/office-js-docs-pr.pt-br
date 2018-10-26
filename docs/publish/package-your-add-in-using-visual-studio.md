---
title: Empacote seu suplemento usando o Visual Studio para preparar a publicação | Microsoft Docs
description: Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681760"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="da753-103">Empacote seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="da753-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="da753-p101">O pacote do Suplemento do Office contém um [arquivo de manifesto](../develop/add-in-manifests.md) XML que você utilizará para publicar o suplemento. Você terá que publicar os arquivos do aplicativo da Web do seu projeto separadamente. Este artigo descreve como implantar o seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="da753-p101">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in. You'll have to publish the web application files of your project separately. This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="da753-107">Para implantar seu projeto Web usando o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="da753-107">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="da753-108">Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="da753-108">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="da753-109">No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="da753-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="da753-110">A página **Publicar seu suplemento** é exibida.</span><span class="sxs-lookup"><span data-stu-id="da753-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="da753-111">Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.</span><span class="sxs-lookup"><span data-stu-id="da753-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="da753-112">Um perfil de publicação especifica o servidor de implantação, as credenciais necessárias para fazer logon no servidor, os bancos de dados a serem implantados e outras opções de implantação.</span><span class="sxs-lookup"><span data-stu-id="da753-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="da753-p102">Se você escolher **Novo...**, o assistente será exibido com a página **Criar perfil de publicação**. Você pode usar esse assistente para importar um perfil de publicação de um provedor de hospedagem de sites da Web, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no próximo procedimento.</span><span class="sxs-lookup"><span data-stu-id="da753-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="da753-115">Para obter mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, confira [Criar um Perfil de Publicação](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="da753-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="da753-116">Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.</span><span class="sxs-lookup"><span data-stu-id="da753-116">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="da753-p103">A caixa de diálogo  **Publicar** será exibida. Para obter mais informações sobre como usar esse assistente, consulte [Tutorial: Implantar um projeto Web usando o On-Click Publishing no Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="da753-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="da753-119">Para empacotar seu suplemento usando o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="da753-119">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="da753-120">Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="da753-120">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="da753-121">Na página **Publicar seu suplemento**, escolha o botão**Empacotar o suplemento**.</span><span class="sxs-lookup"><span data-stu-id="da753-121">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="da753-122">Será exibido o assistente com a página **Empacotar o suplemento**.</span><span class="sxs-lookup"><span data-stu-id="da753-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="da753-123">Na caixa  **Onde seu site da Web está hospedado?**, insira a URL do site da Web que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.</span><span class="sxs-lookup"><span data-stu-id="da753-123">In the **Where is your website hosted?** dropdown list, select or enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="da753-124">Os sites da Web do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="da753-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="da753-125">O Visual Studio gera os arquivos que você precisa para publicar seu suplemento e, em seguida, abre a pasta de saída da publicação.</span><span class="sxs-lookup"><span data-stu-id="da753-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="da753-126">Se você planeja enviar o suplemento para o AppSource, pode escolher o botão **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação do suplemento.</span><span class="sxs-lookup"><span data-stu-id="da753-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="da753-127">Você deve resolver todos os problemas antes de enviar o suplemento para o repositório.</span><span class="sxs-lookup"><span data-stu-id="da753-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="da753-p105">Agora é possível carregar o manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="da753-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="da753-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="da753-131">See also</span></span>

- [<span data-ttu-id="da753-132">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="da753-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="da753-133">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="da753-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
