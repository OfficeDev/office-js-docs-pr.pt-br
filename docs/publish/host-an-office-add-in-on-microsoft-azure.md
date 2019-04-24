---
title: Hospedar um suplemento do Office no Microsoft Azure | Microsoft Docs
description: Saiba como implantar o aplicativo Web de um suplemento no Azure e realizar sideload do suplemento para testar em um aplicativo cliente do Office.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 5db98ca65aac019a027592a442f427ee3b6126f1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451091"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="f2308-103">Hospedar um Suplemento do Office no Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="f2308-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="f2308-p101">Os Suplementos do Office mais simples contêm um arquivo de manifesto XML e uma página HTML. O arquivo de manifesto XML descreve características do suplemento, como seu nome, quais aplicativos clientes do Office podem ser executados e a URL da página HTML do suplemento. A página HTML está contida em um aplicativo Web com o qual os usuários interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Você pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="f2308-108">Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="f2308-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f2308-109">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="f2308-109">Prerequisites</span></span> 

1. <span data-ttu-id="f2308-110">Instale o [Visual Studio 2017](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.</span><span class="sxs-lookup"><span data-stu-id="f2308-110">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-111">Se você tiver instalado o Visual Studio 2017 anteriormente, [use o Instalador do Visual Studio](/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="f2308-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="f2308-112">Instalar o Office.</span><span class="sxs-lookup"><span data-stu-id="f2308-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-113">Se você ainda não tem o Office, [registre-se para fazer uma avaliação gratuita de um mês](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="f2308-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="f2308-114">Obtenha uma assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-115">Se você ainda não tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do Visual Studio](https://azure.microsoft.com/pt-BR/pricing/member-offers/visual-studio-subscriptions/) ou [registrar-se para uma avaliação gratuita](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="f2308-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pt-BR/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="f2308-116">Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="f2308-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="f2308-117">Abra o Explorador de Arquivos em seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="f2308-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="f2308-118">Clique com o botão direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.</span><span class="sxs-lookup"><span data-stu-id="f2308-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="f2308-119">Nomeie a nova pasta AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="f2308-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="f2308-120">Clique com o botão direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas específicas**.</span><span class="sxs-lookup"><span data-stu-id="f2308-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="f2308-121">Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="f2308-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="f2308-p102">Nesta explicação passo a passo, você está usando um compartilhamento de arquivos local como um catálogo confiável onde armazenará o arquivo de manifesto XML do suplemento. Em um cenário real, em vez disso, é possível optar por [implantar o arquivo de manifesto XML a um catálogo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="f2308-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="f2308-124">Etapa 2:Adicionar o compartilhamento de arquivos ao catálogo de Suplementos Confiáveis</span><span class="sxs-lookup"><span data-stu-id="f2308-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="f2308-125">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="f2308-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-126">Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="f2308-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="f2308-127">Escolha **Arquivo** > **Opções**.</span><span class="sxs-lookup"><span data-stu-id="f2308-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="f2308-128">Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="f2308-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="f2308-p103">Na caixa de diálogo **Central de Confiabilidade**, escolha **Catálogos de Suplementos Confiáveis**. Digite o caminho UNC (convenção universal de nomenclatura) para o compartilhamento de arquivos que você criou anteriormente como a **URL do Catálogo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar catálogo**.</span><span class="sxs-lookup"><span data-stu-id="f2308-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="f2308-131">Marque a caixa de seleção **Mostrar no Menu**.</span><span class="sxs-lookup"><span data-stu-id="f2308-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-132">Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um catálogo de suplementos da Web confiável, o suplemento aparece em **Pasta Compartilhada** na caixa de diálogo **Suplementos do Office** quando o usuário navega até a guia **Inserir** na faixa de opções e escolhe **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f2308-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="f2308-133">Feche o Word.</span><span class="sxs-lookup"><span data-stu-id="f2308-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="f2308-134">Etapa 3: Criar um aplicativo Web no Azure</span><span class="sxs-lookup"><span data-stu-id="f2308-134">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="f2308-135">Crie um aplicativo Web vazio no Azure usando [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou o [portal do Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="f2308-135">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="f2308-136">Usar o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="f2308-136">Using Visual Studio 2017</span></span>

<span data-ttu-id="f2308-137">Para criar o aplicativo Web usando o Visual Studio 2017, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="f2308-137">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="f2308-p104">No Visual Studio, no menu **Exibir**, escolha **Gerenciador de Servidores**. Clique com o botão direito do mouse em **Azure** e escolha **Conectar-se à assinatura do Microsoft Azure**. Siga as instruções para se conectar à sua assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>

2. <span data-ttu-id="f2308-141">No Visual Studio, no **Gerenciador de Servidores**, expanda **Azure**, clique com o botão direito do mouse em **Serviço de Aplicativo** e escolha **Criar novo aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="f2308-141">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>

3. <span data-ttu-id="f2308-142">Na caixa de diálogo **Criar Serviço de Aplicativo**, forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="f2308-142">In the **Create App Service** dialog box, provide this information:</span></span>

      - <span data-ttu-id="f2308-p105">Insira um **Nome do Aplicativo Web** exclusivo para seu site. O Azure verifica se o nome do site é exclusivo em todo o domínio azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="f2308-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="f2308-145">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="f2308-145">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="f2308-p106">Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="f2308-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="f2308-p107">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site. Se você criar um novo plano, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="f2308-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>

      - <span data-ttu-id="f2308-150">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="f2308-150">Choose **Create**.</span></span>

    <span data-ttu-id="f2308-151">O novo aplicativo Web aparece no **Gerenciador de Servidores** em **Azure** >> **Serviço de Aplicativo** >> (o grupo de recursos escolhido).</span><span class="sxs-lookup"><span data-stu-id="f2308-151">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>

4. <span data-ttu-id="f2308-p108">Clique com o botão direito do mouse no novo aplicativo Web e escolha **Exibir no Navegador**. O navegador será aberto e exibirá uma página da Web com a mensagem "Seu aplicativo de Serviço de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="f2308-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>

5. <span data-ttu-id="f2308-154">Na barra de endereços do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="f2308-154">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="f2308-155">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="f2308-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

### <a name="using-the-azure-portal"></a><span data-ttu-id="f2308-156">Usar o portal do Azure</span><span class="sxs-lookup"><span data-stu-id="f2308-156">Using the Azure portal</span></span>

<span data-ttu-id="f2308-157">Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="f2308-157">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="f2308-158">Faça logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-158">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="f2308-159">Escolha **Novo** > **Web + Celular** > **Aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="f2308-159">Choose **New** > **Web + Mobile** > **Web App**.</span></span>

3. <span data-ttu-id="f2308-160">Na caixa de diálogo **Criar Aplicativo Web**, forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="f2308-160">In the **Web App Create** dialog box, provide this information:</span></span>

      - <span data-ttu-id="f2308-p109">Insira um **Nome de aplicativo** exclusivo para seu site. O Azure verifica se o nome do site é exclusivo em todo o domínio azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="f2308-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="f2308-163">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="f2308-163">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="f2308-p110">Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="f2308-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="f2308-166">Escolha o **SO** para seu site.</span><span class="sxs-lookup"><span data-stu-id="f2308-166">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="f2308-p111">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site. Se você criar um novo plano, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="f2308-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>

      - <span data-ttu-id="f2308-169">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="f2308-169">Choose **Create**.</span></span>

4. <span data-ttu-id="f2308-170">Escolha **Notificações** (o ícone de sino localizado na borda superior do portal do Azure) e, em seguida, escolha a notificação **Implantações bem-sucedidas** para abrir a página **Visão geral** no portal do Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-170">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f2308-171">A notificação será alterada de **Implantação em andamento** para **Implantações bem-sucedidas** quando a implantação do site for concluída.</span><span class="sxs-lookup"><span data-stu-id="f2308-171">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="f2308-p112">Na seção **Fundamentos** da página **Visão geral** do site no portal do Azure, escolha a URL exibida em **URL**. O navegador será aberto e exibirá uma página da Web com a mensagem "Seu aplicativo de Serviço de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="f2308-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 

6. <span data-ttu-id="f2308-174">Na barra de endereços do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="f2308-174">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="f2308-175">Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="f2308-175">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="f2308-176">Etapa 4: Criar um Suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="f2308-176">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="f2308-177">Inicie o Visual Studio como um administrador.</span><span class="sxs-lookup"><span data-stu-id="f2308-177">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="f2308-178">Escolha **Arquivo** > **Novo** > **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="f2308-178">Choose **File** > **New** > **Project**.</span></span>

3. <span data-ttu-id="f2308-179">Em **Modelos**, expanda **Visual C#** (ou **Visual Basic**), expanda **Office/SharePoint** e escolha **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f2308-179">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>

4. <span data-ttu-id="f2308-180">Escolha **Suplemento da Web do Word** e escolha **OK** para aceitar as configurações padrão.</span><span class="sxs-lookup"><span data-stu-id="f2308-180">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>

<span data-ttu-id="f2308-181">O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web.</span><span class="sxs-lookup"><span data-stu-id="f2308-181">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="f2308-182">Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure</span><span class="sxs-lookup"><span data-stu-id="f2308-182">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="f2308-183">Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções** a fim de ver ambos os projetos para a solução.</span><span class="sxs-lookup"><span data-stu-id="f2308-183">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>

2. <span data-ttu-id="f2308-p113">Clique com botão direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web contém arquivos do aplicativo Web do suplemento do Office, portanto, esse é o projeto que você publica no Azure.</span><span class="sxs-lookup"><span data-stu-id="f2308-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="f2308-186">Na guia **Publicar**:</span><span class="sxs-lookup"><span data-stu-id="f2308-186">On the **Publish** tab:</span></span>

      - <span data-ttu-id="f2308-187">Escolha **Serviço de Aplicativo do Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="f2308-187">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="f2308-188">Escolha **Selecionar Existentes**.</span><span class="sxs-lookup"><span data-stu-id="f2308-188">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="f2308-189">Escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="f2308-189">Choose **Publish**.</span></span>

4. <span data-ttu-id="f2308-190">Na caixa de diálogo **Serviço de Aplicativo**, localize e escolha o aplicativo Web que você criou na [Etapa 3: criar um aplicativo Web no Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="f2308-190">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="f2308-p114">O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrirá e mostrará uma página da Web com o texto "Seu aplicativo de Serviço de Aplicativo foi criado." Esta é a página padrão atual do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="f2308-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

 <span data-ttu-id="f2308-p115">Para ver a página da Web do seu suplemento, altere a URL para que ela use HTTPS e especifique o caminho da página HTML do seu suplemento (por exemplo: https://YourDomain.azurewebsites.net/Home.html). Isso confirma que o aplicativo Web do suplemento já está hospedado no Azure. Copie a URL raiz (por exemplo: https://YourDomain.azurewebsites.net); você precisará dela ao editar o arquivo de manifesto do suplemento mais tarde neste artigo.</span><span class="sxs-lookup"><span data-stu-id="f2308-p115">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html). This confirms that your add-in's web app is now hosted on Azure. Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="f2308-197">Etapa 6: Editar e implantar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="f2308-197">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="f2308-198">No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.</span><span class="sxs-lookup"><span data-stu-id="f2308-198">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="f2308-p116">Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo do manifesto XML do suplemento é aberto.</span><span class="sxs-lookup"><span data-stu-id="f2308-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="f2308-p117">No arquivo de manifesto XML, localize e substitua todas as instâncias de "~remoteAppUrl" pela URL raiz do aplicativo Web do suplemento no Azure. Esta é a URL que você copiou anteriormente depois que publicou o aplicativo Web do suplemento no Azure (por exemplo: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="f2308-p117">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="f2308-p118">Escolha **Arquivo** e, em seguida, **Salvar Tudo**. Feche o arquivo do manifesto XML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="f2308-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>

5. <span data-ttu-id="f2308-205">No **Gerenciador de Soluções**, clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir Pasta no Gerenciador de Arquivos**.</span><span class="sxs-lookup"><span data-stu-id="f2308-205">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>

6. <span data-ttu-id="f2308-206">Copie o arquivo de manifesto XML do suplemento (por exemplo, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="f2308-206">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 

7. <span data-ttu-id="f2308-207">Navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.</span><span class="sxs-lookup"><span data-stu-id="f2308-207">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="f2308-208">Etapa 7: Inserir e executar o suplemento no aplicativo cliente do Office</span><span class="sxs-lookup"><span data-stu-id="f2308-208">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="f2308-209">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="f2308-209">Start Word and create a document.</span></span>

2. <span data-ttu-id="f2308-210">Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="f2308-210">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="f2308-p119">Na caixa de diálogo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um catálogo de suplementos confiáveis (na [Etapa 2: adicionar o compartilhamento de arquivos ao catálogo de suplementos confiáveis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de diálogo. Você deve ver um ícone de seu suplemento de exemplo.</span><span class="sxs-lookup"><span data-stu-id="f2308-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="f2308-p120">Escolha o ícone para seu suplemento e escolha **Adicionar**. Um botão **Mostrar Painel de Tarefas** para seu suplemento é adicionado à faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="f2308-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="f2308-p121">Na faixa de opções da guia **Página Inicial**, escolha o botão **Mostrar Painel de Tarefas**. O suplemento é aberto em um painel de tarefas à direita do documento atual.</span><span class="sxs-lookup"><span data-stu-id="f2308-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="f2308-p122">Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o botão **Realçar!** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="f2308-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="f2308-220">Confira também</span><span class="sxs-lookup"><span data-stu-id="f2308-220">See also</span></span>

- [<span data-ttu-id="f2308-221">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="f2308-221">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="f2308-222">Empacotar seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="f2308-222">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
