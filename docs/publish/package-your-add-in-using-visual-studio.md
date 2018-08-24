---
title: Empacotar seu suplemento usando o Visual Studio para preparar a publicação
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 89f59d06ff305e0d0fd062a36f7e9f756415df45
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925245"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="51ded-102">Empacotar seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="51ded-102">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="51ded-103">Seu pacote de Suplemento do Office contém um [arquivo de manifesto XML](../develop/add-in-manifests.md) que deve ser usado para publicar o suplemento.</span><span class="sxs-lookup"><span data-stu-id="51ded-103">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="51ded-104">Você terá que publicar os arquivos do aplicativo Web do seu projeto separadamente.</span><span class="sxs-lookup"><span data-stu-id="51ded-104">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="51ded-105">Este artigo descreve como implantar seu projeto Web e empacotar seu suplemento usando o Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="51ded-105">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="51ded-106">Para implantar seu projeto Web usando o Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="51ded-106">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="51ded-107">Conclua as etapas a seguir para implantar seu projeto Web usando o Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="51ded-107">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="51ded-108">No **Gerenciador de Soluções**, abra o menu de atalho do projeto do suplemento e escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="51ded-108">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="51ded-109">A página **Publicar seu suplemento** é exibida.</span><span class="sxs-lookup"><span data-stu-id="51ded-109">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="51ded-110">Na lista suspensa **Perfil atual**, selecione um perfil ou escolha **Novo...** para criar um novo perfil.</span><span class="sxs-lookup"><span data-stu-id="51ded-110">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="51ded-111">Um perfil de publicação especifica o servidor que você está implantando, as credenciais necessárias para fazer logon no servidor, os bancos de dados para implantar e outras opções de implantação.</span><span class="sxs-lookup"><span data-stu-id="51ded-111">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="51ded-p102">Se você escolher **Novo...**, o assistente **Criar perfil de publicação** será exibido. Use esse assistente para importar um perfil de publicação de um site de hospedagem, como o Microsoft Azure, ou criar um novo perfil e adicionar seu servidor, as credenciais e outras configurações no procedimento seguinte.</span><span class="sxs-lookup"><span data-stu-id="51ded-p102">If you choose  **New ...**, the  **Create publishing profile** wizard appears. You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="51ded-114">Para mais informações sobre como importar perfis de publicação ou criar novos perfis de publicação, confira [Criar um Perfil de Publicação](http://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="51ded-114">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](http://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="51ded-115">Na página **Publicar seu suplemento**, escolha o link **Implantar seu projeto Web**.</span><span class="sxs-lookup"><span data-stu-id="51ded-115">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="51ded-p103">A caixa de diálogo  **Publicar Web** aparece. Para mais informações sobre a utilização do desse assistente, veja [Instruções: Implantar um Projeto da Web usando o On-Click Publishing no Visual Studio](http://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="51ded-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](http://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="51ded-118">Para empacotar seu suplemento usando o Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="51ded-118">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="51ded-119">Conclua as etapas a seguir para empacotar seu suplemento usando o Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="51ded-119">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="51ded-120">Na página **Publicar seu suplemento**, escolha o link **Empacotar o suplemento**.</span><span class="sxs-lookup"><span data-stu-id="51ded-120">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="51ded-121">O assistente **Publicar Suplementos do Office e do SharePoint** é exibido.</span><span class="sxs-lookup"><span data-stu-id="51ded-121">The **Publish Office and SharePoint Add-ins** wizard appears.</span></span>
    
2. <span data-ttu-id="51ded-122">Na lista suspensa **Onde seu site está hospedado?**, escolha ou digite a URL do site que hospedará os arquivos de conteúdo do seu suplemento e escolha **Concluir**.</span><span class="sxs-lookup"><span data-stu-id="51ded-122">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="51ded-p104">Você deve especificar uma URL que comece com o prefixo HTTPS para concluir este assistente. Se você quiser usar um ponto de extremidade HTTP para o site, abra o arquivo de manifesto XML em um editor de texto após criar o pacote e substitua o prefixo HTTPS do site por um prefixo HTTP.</span><span class="sxs-lookup"><span data-stu-id="51ded-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="51ded-125"> Os sites do Azure fornecem um ponto de extremidade HTTPS automaticamente.</span><span class="sxs-lookup"><span data-stu-id="51ded-125">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="51ded-126">O Visual Studio gera os arquivos nos quais você precisa publicar seu suplemento e, em seguida, abre a pasta de saída de publicação.</span><span class="sxs-lookup"><span data-stu-id="51ded-126">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="51ded-p105">Se você planeja enviar seu suplemento ao AppSource, pode escolher o link **Executar uma verificação de validação** para identificar problemas que possam impedir a aceitação de seu suplemento. Resolva todos os problemas antes de enviar seu suplemento para a loja.</span><span class="sxs-lookup"><span data-stu-id="51ded-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="51ded-p106">Agora é possível carregar seu manifesto XML no local apropriado para [publicar seu suplemento](../publish/publish.md). É possível encontrar o manifesto XML em `OfficeAppManifests` na pasta `app.publish`. Por exemplo:</span><span class="sxs-lookup"><span data-stu-id="51ded-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="51ded-132">Veja também</span><span class="sxs-lookup"><span data-stu-id="51ded-132">See also</span></span>

- [<span data-ttu-id="51ded-133">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="51ded-133">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="51ded-134">Disponibilizar suas soluções no AppSource e no Office</span><span class="sxs-lookup"><span data-stu-id="51ded-134">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
