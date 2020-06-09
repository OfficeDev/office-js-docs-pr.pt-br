---
title: Hospedar um suplemento do Office no Microsoft Azure | Microsoft Docs
description: Saiba como implantar o aplicativo Web de um suplemento no Azure e realizar sideload do suplemento para testar em um aplicativo cliente do Office.
ms.date: 10/16/2019
localization_priority: Normal
ms.openlocfilehash: a546e53d03bb08dd216c04eab9b684f651f9c5de
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612056"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="f4c77-103">Hospedar um Suplemento do Office no Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="f4c77-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="f4c77-p101">Os Suplementos do Office mais simples contêm um arquivo de manifesto XML e uma página HTML. O arquivo de manifesto XML descreve características do suplemento, como seu nome, quais aplicativos clientes do Office podem ser executados e a URL da página HTML do suplemento. A página HTML está contida em um aplicativo Web com o qual os usuários interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Você pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="f4c77-108">Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="f4c77-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f4c77-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="f4c77-109">Prerequisites</span></span> 

1. <span data-ttu-id="f4c77-110">Instale o [Visual Studio 2019](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4c77-111">Se você tiver instalado o Visual Studio 2019 anteriormente, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="f4c77-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="f4c77-112">Instalar o Office.</span><span class="sxs-lookup"><span data-stu-id="f4c77-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4c77-113">Se você ainda não tem o Office, [registre-se para fazer uma avaliação gratuita de um mês](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="f4c77-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="f4c77-114">Obtenha uma assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="f4c77-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4c77-115">Se você ainda não tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) ou [registrar-se para uma avaliação gratuita](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="f4c77-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="f4c77-116">Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="f4c77-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="f4c77-117">Abra o Explorador de Arquivos em seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f4c77-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="f4c77-118">Clique com o botão direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="f4c77-119">Nomeie a nova pasta AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="f4c77-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="f4c77-120">Clique com o botão direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas específicas**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="f4c77-121">Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="f4c77-p102">Nesta explicação passo a passo, você está usando um compartilhamento de arquivos local como um catálogo confiável onde armazenará o arquivo de manifesto XML do suplemento. Em um cenário real, em vez disso, é possível optar por [implantar o arquivo de manifesto XML a um catálogo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span><span class="sxs-lookup"><span data-stu-id="f4c77-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="f4c77-124">Etapa 2:Adicionar o compartilhamento de arquivos ao catálogo de Suplementos Confiáveis</span><span class="sxs-lookup"><span data-stu-id="f4c77-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="f4c77-125">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="f4c77-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4c77-126">Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="f4c77-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="f4c77-127">Escolha **Arquivo** > **Opções**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="f4c77-128">Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="f4c77-p103">Na caixa de diálogo **Central de Confiabilidade**, escolha **Catálogos de Suplementos Confiáveis**. Digite o caminho UNC (convenção universal de nomenclatura) para o compartilhamento de arquivos que você criou anteriormente como a **URL do Catálogo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar catálogo**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="f4c77-131">Marque a caixa de seleção **Mostrar no Menu**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4c77-132">Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um catálogo de suplementos da Web confiável, o suplemento aparece em **Pasta Compartilhada** na caixa de diálogo **Suplementos do Office** quando o usuário navega até a guia **Inserir** na faixa de opções e escolhe **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="f4c77-133">Feche o Word.</span><span class="sxs-lookup"><span data-stu-id="f4c77-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="f4c77-134">Etapa 3: Criar um aplicativo Web no Azure usando o portal do Azure</span><span class="sxs-lookup"><span data-stu-id="f4c77-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="f4c77-135">Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="f4c77-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="f4c77-136">Faça logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.</span><span class="sxs-lookup"><span data-stu-id="f4c77-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="f4c77-137">Em **Serviços do Azure**, selecione **Aplicativos Web**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="f4c77-138">Na página **Serviço de Aplicativo**, selecione **Adicionar**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="f4c77-139">Forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="f4c77-139">Provide this information:</span></span>

      - <span data-ttu-id="f4c77-140">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="f4c77-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="f4c77-p105">Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p105">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="f4c77-143">Insira um **Nome de aplicativo** exclusivo para seu site.</span><span class="sxs-lookup"><span data-stu-id="f4c77-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="f4c77-144">O Azure verifica se o nome do site é exclusivo em todo o domínio azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="f4c77-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="f4c77-145">Escolha se deseja publicar usando um código ou um contêiner do docker.</span><span class="sxs-lookup"><span data-stu-id="f4c77-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="f4c77-146">Especificar uma **Pilha de tempo de execução**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="f4c77-147">Escolha o **SO** para seu site.</span><span class="sxs-lookup"><span data-stu-id="f4c77-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="f4c77-148">Escolha uma **Região**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="f4c77-149">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="f4c77-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="f4c77-150">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-150">Choose **Create**.</span></span>

4. <span data-ttu-id="f4c77-151">A próxima página informa que a implantação está em andamento e quando ela é concluída.</span><span class="sxs-lookup"><span data-stu-id="f4c77-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="f4c77-152">Quando estiver concluída, selecione **Ir ao recurso**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="f4c77-153">Na seção **Visão geral**, escolha a URL exibida em **URL**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="f4c77-154">O navegador será aberto e exibirá uma página da Web com a mensagem “Seu aplicativo de Serviço de Aplicativo está funcionando”.</span><span class="sxs-lookup"><span data-stu-id="f4c77-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="f4c77-155">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="f4c77-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="f4c77-156">Etapa 4: Criar um Suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="f4c77-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="f4c77-157">Inicie o Visual Studio como um administrador.</span><span class="sxs-lookup"><span data-stu-id="f4c77-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="f4c77-158">Escolha **Criar um novo projeto**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="f4c77-159">Usando a caixa de pesquisa, insira **suplemento**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="f4c77-160">Escolha **Suplemento da Web do Word** como o tipo de projeto e, em seguida, escolha **Avançar** para aceitar as configurações padrão.</span><span class="sxs-lookup"><span data-stu-id="f4c77-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="f4c77-161">O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web.</span><span class="sxs-lookup"><span data-stu-id="f4c77-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="f4c77-162">Para criar um suplemento para outro tipo de host do Office, como o Excel, repita as etapas e escolha um tipo de projeto com o host do Office desejado.</span><span class="sxs-lookup"><span data-stu-id="f4c77-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="f4c77-163">Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure</span><span class="sxs-lookup"><span data-stu-id="f4c77-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="f4c77-164">Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções**, em seguida, selecione **Serviço de Aplicativo**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="f4c77-p110">Clique com botão direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web contém arquivos do aplicativo Web do suplemento do Office, portanto, esse é o projeto que você publica no Azure.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p110">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="f4c77-167">Na guia **Publicar**:</span><span class="sxs-lookup"><span data-stu-id="f4c77-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="f4c77-168">Escolha **Serviço de Aplicativo do Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="f4c77-169">Escolha **Selecionar Existentes**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="f4c77-170">Escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="f4c77-p111">O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrirá e mostrará uma página da Web com o texto "Seu aplicativo de Serviço de Aplicativo foi criado." Esta é a página padrão atual do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p111">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

5. <span data-ttu-id="f4c77-174">Copie a URL raiz (por exemplo:https://YourDomain.azurewebsites.net); você precisará dela ao editar o arquivo de manifesto do suplemento, mais tarde neste artigo.</span><span class="sxs-lookup"><span data-stu-id="f4c77-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="f4c77-175">Etapa 6: Editar e implantar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="f4c77-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="f4c77-176">No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.</span><span class="sxs-lookup"><span data-stu-id="f4c77-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="f4c77-p112">Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo do manifesto XML do suplemento é aberto.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p112">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="f4c77-p113">No arquivo de manifesto XML, localize e substitua todas as instâncias de "~remoteAppUrl" pela URL raiz do aplicativo Web do suplemento no Azure. Esta é a URL que você copiou anteriormente depois que publicou o aplicativo Web do suplemento no Azure (por exemplo: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="f4c77-p113">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="f4c77-181">Escolha **Arquivo** e **Salvar tudo**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="f4c77-182">Em seguida, copie o arquivo do manifesto XML (por exemplo, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="f4c77-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="f4c77-183">Usando o programa **Gerenciador de Arquivos**, navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.</span><span class="sxs-lookup"><span data-stu-id="f4c77-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="f4c77-184">Etapa 7: Inserir e executar o suplemento no aplicativo cliente do Office</span><span class="sxs-lookup"><span data-stu-id="f4c77-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="f4c77-185">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="f4c77-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="f4c77-186">Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f4c77-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="f4c77-p115">Na caixa de diálogo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um catálogo de suplementos confiáveis (na [Etapa 2: adicionar o compartilhamento de arquivos ao catálogo de suplementos confiáveis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de diálogo. Você deve ver um ícone de seu suplemento de exemplo.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p115">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="f4c77-p116">Escolha o ícone para seu suplemento e escolha **Adicionar**. Um botão **Mostrar Painel de Tarefas** para seu suplemento é adicionado à faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p116">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="f4c77-p117">Na faixa de opções da guia **Página Inicial**, escolha o botão **Mostrar Painel de Tarefas**. O suplemento é aberto em um painel de tarefas à direita do documento atual.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p117">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="f4c77-p118">Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o botão **Realçar!** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f4c77-p118">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="f4c77-196">Confira também</span><span class="sxs-lookup"><span data-stu-id="f4c77-196">See also</span></span>

- [<span data-ttu-id="f4c77-197">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="f4c77-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="f4c77-198">Publicar seu suplemento usando o Visual Studio</span><span class="sxs-lookup"><span data-stu-id="f4c77-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)
