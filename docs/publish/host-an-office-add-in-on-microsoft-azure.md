---
title: Hospedar um Suplemento do Office no Microsoft Azure
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 62fc3c6dc212efc47493f2bcb3a994fb4db6a752
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945562"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="9d2ec-102">Hospedar um Suplemento do Office no Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="9d2ec-102">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="9d2ec-p101">Os Suplementos do Office mais simples contêm um arquivo de manifesto XML e uma página HTML. O arquivo de manifesto XML descreve características do suplemento, como seu nome, quais aplicativos clientes do Office podem ser executados e a URL da página HTML do suplemento. A página HTML está contida em um aplicativo Web com o qual os usuários interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Você pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="9d2ec-107">Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-107">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9d2ec-108">Pré-requisitos</span><span class="sxs-lookup"><span data-stu-id="9d2ec-108">Prerequisites</span></span> 

1. <span data-ttu-id="9d2ec-109">Instale o [Visual Studio 2017](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-109">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d2ec-110">Se você tiver instalado o Visual Studio 2017 anteriormente, [use o Instalador do Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="9d2ec-111">Instalar o Office.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-111">Install Office.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="9d2ec-112">Se você ainda não tem o Office, [registre-se para fazer uma avaliação gratuita de um mês](http://office.microsoft.com/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-112">If you don't already have Office 2016, you can [register for a free 1-month trial](http://office.microsoft.com/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span></span>

3.  <span data-ttu-id="9d2ec-113">Obtenha uma assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-113">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="9d2ec-114">Se você ainda não tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do MSDN](http://www.windowsazure.com/pricing/member-offers/msdn-benefits/) ou [registrar-se gratuitamente para uma avaliação gratuita](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-114">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="9d2ec-115">Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="9d2ec-115">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="9d2ec-116">Abra o Explorador de Arquivos em seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-116">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="9d2ec-117">Clique com o botão direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-117">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="9d2ec-118">Nomeie a nova pasta AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-118">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="9d2ec-119">Clique com o botão direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas específicas**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-119">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="9d2ec-120">Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-120">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="9d2ec-p102">Nesta explicação passo a passo, você está usando um compartilhamento de arquivos local como um catálogo confiável onde armazenará o arquivo de manifesto XML do suplemento. Em um cenário real, em vez disso, é possível optar por [implantar o arquivo de manifesto XML a um catálogo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="9d2ec-123">Etapa 2:adicionar o compartilhamento de arquivos ao catálogo de suplementos confiáveis</span><span class="sxs-lookup"><span data-stu-id="9d2ec-123">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="9d2ec-124">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-124">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d2ec-125">Embora este exemplo use o Word, é possível usar qualquer aplicativo do Office que dê suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-125">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="9d2ec-126">Escolha **Arquivo**  >  **Opções**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-126">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="9d2ec-127">Na caixa de diálogo **Opções do Word**, escolha **Central de Confiabilidade**, depois **Configurações da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-127">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="9d2ec-p103">Na caixa de diálogo **Central de Confiabilidade**, escolha **Catálogos de Suplementos Confiáveis**. Digite o caminho UNC (convenção universal de nomenclatura) para o compartilhamento de arquivos que você criou anteriormente como a **URL do Catálogo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar catálogo**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="9d2ec-130">Marque a caixa de seleção **Mostrar no Menu**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-130">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="9d2ec-131">Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um catálogo de suplementos da Web confiável, o suplemento aparece em **Pasta Compartilhada** na caixa de diálogo **Suplementos do Office** quando o usuário navega até a guia **Inserir** na faixa de opções e escolhe **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-131">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="9d2ec-132">Feche o Word.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-132">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="9d2ec-133">Etapa 3: criar um aplicativo Web no Azure</span><span class="sxs-lookup"><span data-stu-id="9d2ec-133">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="9d2ec-134">Crie um aplicativo Web vazio no Azure usando [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou o [portal do Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-134">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="9d2ec-135">Usar o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="9d2ec-135">Using Visual Studio 2017</span></span>

<span data-ttu-id="9d2ec-136">Para criar o aplicativo Web usando o Visual Studio 2017, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-136">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="9d2ec-p104">No Visual Studio, no menu **Exibir**, escolha **Gerenciador de Servidores**. Clique com o botão direito do mouse em **Azure** e escolha **Conectar-se à assinatura do Microsoft Azure**. Siga as instruções para se conectar à sua assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="9d2ec-140">No Visual Studio, no **Gerenciador de Servidores**, expanda **Azure**, clique com o botão direito do mouse em **Serviço de Aplicativo** e escolha **Criar novo aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-140">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="9d2ec-141">Na caixa de diálogo **Criar Serviço de Aplicativo**, forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="9d2ec-141">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="9d2ec-p105">Insira um **Nome do Aplicativo Web** exclusivo para seu site. O Azure verifica se o nome do site é exclusivo em todo o domínio azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="9d2ec-144">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-144">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="9d2ec-p106">Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="9d2ec-p107">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site. Se você criar um novo plano, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="9d2ec-149">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-149">Choose **Create**.</span></span>

    <span data-ttu-id="9d2ec-150">O novo aplicativo Web aparece no **Gerenciador de Servidores** em **Azure** >> **Serviço de Aplicativo** >> (o grupo de recursos escolhido).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-150">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="9d2ec-p108">Clique com o botão direito do mouse no novo aplicativo Web e escolha **Exibir no Navegador**. O navegador será aberto e exibirá uma página da Web com a mensagem "Seu aplicativo de Serviço de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="9d2ec-153">Na barra de endereços do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-153">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="9d2ec-154"> Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-154">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="9d2ec-155">Usar o portal do Azure</span><span class="sxs-lookup"><span data-stu-id="9d2ec-155">Using the Azure portal</span></span>

<span data-ttu-id="9d2ec-156">Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-156">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="9d2ec-157">Faça logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-157">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="9d2ec-158">Escolha **Novo** > **Web + Celular** > **Aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-158">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="9d2ec-159">Na caixa de diálogo **Criar Aplicativo Web**, forneça estas informações:</span><span class="sxs-lookup"><span data-stu-id="9d2ec-159">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="9d2ec-p109">Insira um **Nome de aplicativo** exclusivo para seu site. O Azure verifica se o nome do site é exclusivo em todo o domínio azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="9d2ec-162">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-162">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="9d2ec-p110">Escolha o **Grupo de Recursos** para seu site. Se você criar um novo grupo, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="9d2ec-165">Escolha o **SO** para seu site.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-165">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="9d2ec-p111">Escolha o **Plano do Serviço de Aplicativo** a ser usado para criar esse site. Se você criar um novo plano, também precisará dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="9d2ec-168">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-168">Choose **Create**.</span></span>

4. <span data-ttu-id="9d2ec-169">Escolha **Notificações** (o ícone de sino localizado na borda superior do portal do Azure) e, em seguida, escolha a notificação **Implantações bem-sucedidas** para abrir a página **Visão geral** no portal do Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-169">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="9d2ec-170">A notificação será alterada de **Implantação em andamento** para **Implantações bem-sucedidas** quando a implantação do site for concluída.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-170">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="9d2ec-p112">Na seção **Fundamentos** da página **Visão geral** do site no portal do Azure, escolha a URL exibida em **URL**. O navegador será aberto e exibirá uma página da Web com a mensagem "Seu aplicativo de Serviço de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="9d2ec-173">Na barra de endereços do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-173">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="9d2ec-174"> Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-174">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="9d2ec-175">Etapa 4: criar um suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="9d2ec-175">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="9d2ec-176">Inicie o Visual Studio como um administrador.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-176">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="9d2ec-177">Escolha **Arquivo** > **Novo** > **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-177">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="9d2ec-178">Em **Modelos**, expanda **Visual C#** (ou **Visual Basic**), expanda **Office/SharePoint** e escolha **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-178">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="9d2ec-179">Escolha **Suplemento da Web do Word** e escolha **OK** para aceitar as configurações padrão.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-179">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="9d2ec-180">O Visual Studio cria um suplemento básico do Word que você pode publicar como está, sem fazer alterações no projeto da Web.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-180">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="9d2ec-181">Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure</span><span class="sxs-lookup"><span data-stu-id="9d2ec-181">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="9d2ec-182">Com seu projeto de suplemento aberto no Visual Studio, expanda o nó da solução no **Gerenciador de Soluções** a fim de ver ambos os projetos para a solução.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-182">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="9d2ec-p113">Clique com botão direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web contém arquivos do aplicativo Web do suplemento do Office, portanto, esse é o projeto que você publica no Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="9d2ec-185">Na guia **Publicar**:</span><span class="sxs-lookup"><span data-stu-id="9d2ec-185">On the **Publish** tab:</span></span>

      - <span data-ttu-id="9d2ec-186">Escolha **Serviço de Aplicativo do Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-186">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="9d2ec-187">Escolha **Selecionar Existentes**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-187">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="9d2ec-188">Escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-188">Choose **Publish**.</span></span> 

6. <span data-ttu-id="9d2ec-189">Na caixa de diálogo **Serviço de Aplicativo**, localize e escolha o aplicativo Web que você criou na [Etapa 3: criar um aplicativo Web no Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-189">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="9d2ec-p114">O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrirá e mostrará uma página da Web com o texto "Seu aplicativo de Serviço de Aplicativo foi criado." Esta é a página padrão atual do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="9d2ec-193">Para ver a página da Web do seu suplemento, altere o URL para que ele use HTTPS e especifique o caminho da página HTML do seu suplemento (por exemplo: https://YourDomain.azurewebsites.net/Home.html).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-193">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="9d2ec-194">Isso confirma que o aplicativo Web do seu suplemento está hospedado no Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-194">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="9d2ec-195">Copie a URL raiz (por exemplo: https://YourDomain.azurewebsites.net); você precisará dela ao editar o arquivo de manifesto do suplemento mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-195">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="9d2ec-196">Etapa 6: editar e implantar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="9d2ec-196">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="9d2ec-197">No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Soluções**, expanda a solução para que ambos os projetos sejam exibidos.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-197">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="9d2ec-p116">Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo de manifesto XML do suplemento é aberto.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="9d2ec-200">No arquivo de manifesto XML, localizar e substituir todas as instâncias de "~ remoteAppUrl" com a URL raiz do aplicativo web do suplemento no Azure.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-200">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="9d2ec-201">Esse é o URL que você copiou anteriormente depois de publicar o aplicativo Web do suplemento no Azure (por exemplo: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-201">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="9d2ec-p118">Escolha **Arquivo** e **Salvar tudo**. Feche o arquivo de manifesto XML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="9d2ec-204">No **Gerenciador de Soluções**, clique com o botão direito do mouse na pasta do manifesto e escolha **Abrir Pasta no Gerenciador de Arquivos**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-204">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="9d2ec-205">Copie o arquivo de manifesto XML do suplemento (por exemplo, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="9d2ec-205">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="9d2ec-206">Navegue até o compartilhamento de arquivos de rede que você criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-206">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="9d2ec-207">Etapa 7: inserir e executar o suplemento no aplicativo cliente do Office</span><span class="sxs-lookup"><span data-stu-id="9d2ec-207">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="9d2ec-208">Inicie o Word e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-208">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="9d2ec-209">Na faixa de opções, escolha **Inserir** > **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-209">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="9d2ec-p119">Na caixa de diálogo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um catálogo de suplementos confiáveis (na [Etapa 2: adicionar o compartilhamento de arquivos ao catálogo de suplementos confiáveis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de diálogo. Você deve ver um ícone de seu suplemento de exemplo.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="9d2ec-p120">Escolha o ícone para seu suplemento e escolha **Adicionar**. Um botão **Mostrar Painel de Tarefas** para seu suplemento é adicionado à faixa de opções.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="9d2ec-p121">Na faixa de opções da guia **Página Inicial**, escolha o botão **Mostrar Painel de Tarefas**. O suplemento é aberto em um painel de tarefas à direita do documento atual.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="9d2ec-p122">Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o botão **Realçar!** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="9d2ec-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="9d2ec-219">Veja também</span><span class="sxs-lookup"><span data-stu-id="9d2ec-219">See also</span></span>

- [<span data-ttu-id="9d2ec-220">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="9d2ec-220">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="9d2ec-221">Empacotar seu suplemento usando o Visual Studio para preparar a publicação</span><span class="sxs-lookup"><span data-stu-id="9d2ec-221">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
