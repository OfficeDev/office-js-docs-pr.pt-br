---
title: Hospedar um suplemento do Office no Microsoft Azure | Microsoft Docs
description: Saiba como implantar o aplicativo Web de um suplemento no Azure e realizar sideload do suplemento para testar em um aplicativo cliente do Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094229"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="1f3e5-103">Hospedar um Suplemento do Office no Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="1f3e5-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="1f3e5-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span></span> <span data-ttu-id="1f3e5-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span></span> <span data-ttu-id="1f3e5-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span></span> <span data-ttu-id="1f3e5-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="1f3e5-108">Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1f3e5-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="1f3e5-109">Prerequisites</span></span> 

1. <span data-ttu-id="1f3e5-110">Instale o [Visual Studio 2019](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1f3e5-111">Se você tiver instalado o Visual Studio 2019 anteriormente, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="1f3e5-112">Instalar o Office.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1f3e5-113">Se você ainda não tem o Office, [registre-se para fazer uma avaliação gratuita de um mês](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="1f3e5-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="1f3e5-114">Obtenha uma assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1f3e5-115">Se você ainda não tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) ou [registrar-se para uma avaliação gratuita](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="1f3e5-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="1f3e5-116">Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="1f3e5-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="1f3e5-117">Abra o Explorador de Arquivos em seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="1f3e5-118">Clique com o botão direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="1f3e5-119">Nomeie a nova pasta AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="1f3e5-120">Clique com o botão direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas específicas**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="1f3e5-121">Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="1f3e5-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span></span> <span data-ttu-id="1f3e5-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span><span class="sxs-lookup"><span data-stu-id="1f3e5-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="1f3e5-124">Etapa 2:Adicionar o compartilhamento de arquivos ao catálogo de Suplementos Confiáveis</span><span class="sxs-lookup"><span data-stu-id="1f3e5-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="1f3e5-125">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1f3e5-126">Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="1f3e5-127">Escolha **Arquivo** > **Opções**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="1f3e5-128">Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="1f3e5-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span></span> <span data-ttu-id="1f3e5-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="1f3e5-131">Marque a caixa de seleção **Mostrar no Menu**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1f3e5-132">Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um catálogo de suplementos da Web confiável, o suplemento aparece em **Pasta Compartilhada** na caixa de diálogo **Suplementos do Office** quando o usuário navega até a guia **Inserir** na faixa de opções e escolhe **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="1f3e5-133">Feche o Word.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="1f3e5-134">Etapa 3: Criar um aplicativo Web no Azure usando o portal do Azure</span><span class="sxs-lookup"><span data-stu-id="1f3e5-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="1f3e5-135">Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="1f3e5-136">Faça logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="1f3e5-137">Em **Serviços do Azure**, selecione **Aplicativos Web**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="1f3e5-138">Na página **Serviço de Aplicativo**, selecione **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="1f3e5-139">Forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="1f3e5-139">Provide this information:</span></span>

      - <span data-ttu-id="1f3e5-140">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="1f3e5-141">Choose the **Resource Group** for your site.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-141">Choose the **Resource Group** for your site.</span></span> <span data-ttu-id="1f3e5-142">If you create a new group, you also need to name it.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-142">If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="1f3e5-143">Insira um **Nome de aplicativo** exclusivo para seu site.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="1f3e5-144">O Azure verifica se o nome do site é exclusivo em todo o domínio azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="1f3e5-145">Escolha se deseja publicar usando um código ou um contêiner do docker.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="1f3e5-146">Especificar uma **Pilha de tempo de execução**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="1f3e5-147">Escolha o **SO** para seu site.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="1f3e5-148">Escolha uma **Região**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="1f3e5-149">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="1f3e5-150">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-150">Choose **Create**.</span></span>

4. <span data-ttu-id="1f3e5-151">A próxima página informa que a implantação está em andamento e quando ela é concluída.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="1f3e5-152">Quando estiver concluída, selecione **Ir ao recurso**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="1f3e5-153">Na seção **Visão geral**, escolha a URL exibida em **URL**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="1f3e5-154">O navegador será aberto e exibirá uma página da Web com a mensagem “Seu aplicativo de Serviço de Aplicativo está funcionando”.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="1f3e5-155">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="1f3e5-156">Etapa 4: Criar um Suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="1f3e5-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="1f3e5-157">Inicie o Visual Studio como um administrador.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="1f3e5-158">Escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="1f3e5-159">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="1f3e5-160">Escolha **Suplemento da Web do Word** como o tipo de projeto e, em seguida, escolha **Avançar** para aceitar as configurações padrão.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="1f3e5-161">O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="1f3e5-162">Para criar um suplemento para outro tipo de host do Office, como o Excel, repita as etapas e escolha um tipo de projeto com o host do Office desejado.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="1f3e5-163">Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure</span><span class="sxs-lookup"><span data-stu-id="1f3e5-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="1f3e5-164">Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções**, em seguida, selecione **Serviço de Aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="1f3e5-165">Right-click the web project and then choose **Publish**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-165">Right-click the web project and then choose **Publish**.</span></span> <span data-ttu-id="1f3e5-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="1f3e5-167">Na guia **Publicar**:</span><span class="sxs-lookup"><span data-stu-id="1f3e5-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="1f3e5-168">Escolha **Serviço de Aplicativo do Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="1f3e5-169">Escolha **Selecionar Existentes**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="1f3e5-170">Escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="1f3e5-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span></span> <span data-ttu-id="1f3e5-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span><span class="sxs-lookup"><span data-stu-id="1f3e5-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span></span> <span data-ttu-id="1f3e5-173">This is the current default page for the web app.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-173">This is the current default page for the web app.</span></span>

5. <span data-ttu-id="1f3e5-174">Copie a URL raiz (por exemplo:https://YourDomain.azurewebsites.net); você precisará dela ao editar o arquivo de manifesto do suplemento, mais tarde neste artigo.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="1f3e5-175">Etapa 6: Editar e implantar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="1f3e5-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="1f3e5-176">No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="1f3e5-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span></span> <span data-ttu-id="1f3e5-178">The add-in XML manifest file opens.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-178">The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="1f3e5-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span></span> <span data-ttu-id="1f3e5-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="1f3e5-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="1f3e5-181">Escolha **Arquivo** e **Salvar tudo**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="1f3e5-182">Em seguida, copie o arquivo do manifesto XML (por exemplo, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="1f3e5-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="1f3e5-183">Usando o programa **Gerenciador de Arquivos**, navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="1f3e5-184">Etapa 7: Inserir e executar o suplemento no aplicativo cliente do Office</span><span class="sxs-lookup"><span data-stu-id="1f3e5-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="1f3e5-185">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="1f3e5-186">Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="1f3e5-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span></span> <span data-ttu-id="1f3e5-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span></span> <span data-ttu-id="1f3e5-189">You should see an icon for your sample add-in.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-189">You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="1f3e5-190">Choose the icon for your add-in and then choose **Add**.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-190">Choose the icon for your add-in and then choose **Add**.</span></span> <span data-ttu-id="1f3e5-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="1f3e5-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="1f3e5-193">The add-in opens in a task pane to the right of the current document.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-193">The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="1f3e5-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span><span class="sxs-lookup"><span data-stu-id="1f3e5-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span></span> <span data-ttu-id="1f3e5-195">button in the task pane.</span><span class="sxs-lookup"><span data-stu-id="1f3e5-195">button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="1f3e5-196">Confira também</span><span class="sxs-lookup"><span data-stu-id="1f3e5-196">See also</span></span>

- [<span data-ttu-id="1f3e5-197">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="1f3e5-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="1f3e5-198">Publicar seu suplemento usando o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="1f3e5-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)
