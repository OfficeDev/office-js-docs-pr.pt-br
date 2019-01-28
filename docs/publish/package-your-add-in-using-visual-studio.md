---
title: Empacotar seu suplemento usando o Visual Studio para preparar a publicação | Microsoft Docs
description: Como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: a135e8e72703c3de60290a9eb7b2e03c63449124
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386432"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="7c99a-103">Empacotar seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="7c99a-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="7c99a-104">Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="7c99a-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="7c99a-105">Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente.</span><span class="sxs-lookup"><span data-stu-id="7c99a-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="7c99a-106">Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="7c99a-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="7c99a-107">Implantar seu projeto Web usando o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="7c99a-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="7c99a-108">Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="7c99a-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="7c99a-109">No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha  **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="7c99a-110">A página **Publicar seu suplemento** é exibida.</span><span class="sxs-lookup"><span data-stu-id="7c99a-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="7c99a-111">Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.</span><span class="sxs-lookup"><span data-stu-id="7c99a-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="7c99a-112">Um perfil de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.</span><span class="sxs-lookup"><span data-stu-id="7c99a-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="7c99a-113">Se você escolher **Novo ...**, o assistente é exibido com a página **Criar perfil de Publicação**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-113">If you choose  **New ...**, a wizard appears with the **Create publishing profile** page.</span></span> <span data-ttu-id="7c99a-114">Use esse assistente para importar um perfil de publicação de um site de hospedagem, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no procedimento seguinte.</span><span class="sxs-lookup"><span data-stu-id="7c99a-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="7c99a-115">Para mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, veja [Criar um Perfil de Publicação](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="7c99a-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="7c99a-116">Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-116">On the **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="7c99a-117">A caixa de diálogo **Publicar** é exibida.</span><span class="sxs-lookup"><span data-stu-id="7c99a-117">The  **Publish** dialog box appears.</span></span> <span data-ttu-id="7c99a-118">Para mais informações sobre como usar o assistente, veja [Como: implantar um Projeto Web usando a Publicação On-Click no Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="7c99a-118">For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="7c99a-119">Empacotar seu suplemento usando o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="7c99a-119">To package your add-in using Visual Studio 2017</span></span>

<span data-ttu-id="7c99a-120">Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="7c99a-120">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="7c99a-121">Na página **Publicar seu suplemento**, escolha o botão **Empacotar o suplemento**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-121">In the **Publish your add-in** page, choose the **Package the add-in** button.</span></span>
    
    <span data-ttu-id="7c99a-122">Um assistente é exibido com a página **Empacotar o suplemento**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="7c99a-123">Na lista suspensa **Onde seu site está hospedado?**, escolha ou digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.</span><span class="sxs-lookup"><span data-stu-id="7c99a-123">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="7c99a-124">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="7c99a-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="7c99a-125">O Visual Studio gera os arquivos nos quais você precisa publicar seu suplemento e, em seguida, abre a pasta de saída de publicação.</span><span class="sxs-lookup"><span data-stu-id="7c99a-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="7c99a-126">Se você pretende enviar seu suplemento ao AppSource, escolha o botão **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="7c99a-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="7c99a-127">Você deve resolver todos os problemas antes de enviar seu suplemento para a loja.</span><span class="sxs-lookup"><span data-stu-id="7c99a-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="7c99a-p105">Agora é possível carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="7c99a-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="7c99a-131">Confira também</span><span class="sxs-lookup"><span data-stu-id="7c99a-131">See also</span></span>

- [<span data-ttu-id="7c99a-132">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="7c99a-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="7c99a-133">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="7c99a-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
