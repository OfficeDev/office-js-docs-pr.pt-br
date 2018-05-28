---
title: Hospedar um Suplemento do Office no Microsoft Azure
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: f0d6a5a10d2ce0620b42566be03e2d36f8a922f2
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="1a626-102">Hospedar um Suplemento do Office no Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="1a626-102">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="1a626-p101">Os Suplementos do Office mais simples cont?m um arquivo de manifesto XML e uma p?gina HTML. O arquivo de manifesto XML descreve caracter?sticas do suplemento, como seu nome, quais aplicativos clientes do Office podem ser executados e a URL da p?gina HTML do suplemento. A p?gina HTML est? contida em um aplicativo Web com o qual os usu?rios interagem quando instalam e executam seu suplemento dentro de um aplicativo cliente do Office. Voc? pode hospedar o aplicativo Web de um suplemento do Office em qualquer plataforma de hospedagem Web, incluindo o Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="1a626-107">Este artigo descreve como implantar o aplicativo Web de um suplemento no Azure e [realizar sideload do suplemento](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) para teste em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="1a626-107">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1a626-108">Pr?-requisitos</span><span class="sxs-lookup"><span data-stu-id="1a626-108">Prerequisites</span></span> 

1. <span data-ttu-id="1a626-109">Instale o [Visual Studio 2017](https://www.visualstudio.com/downloads) e opte por incluir a carga de trabalho de **desenvolvimento do Azure**.</span><span class="sxs-lookup"><span data-stu-id="1a626-109">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a626-110">Se voc? tiver instalado o Visual Studio 2017 anteriormente, [use o Instalador do Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) para garantir que a carga de trabalho de **desenvolvimento do Azure** esteja instalada.</span><span class="sxs-lookup"><span data-stu-id="1a626-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="1a626-111">Instale o Office 2016.</span><span class="sxs-lookup"><span data-stu-id="1a626-111">Install Office 2016.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="1a626-112">Se voc? ainda n?o tem o Office 2016, [registre-se para fazer uma avalia??o gratuita de um m?s](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span><span class="sxs-lookup"><span data-stu-id="1a626-112">If you don't already have Office 2016, you can [register for a free 1-month trial](http://office.microsoft.com/en-us/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span></span>

3.  <span data-ttu-id="1a626-113">Obtenha uma assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-113">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="1a626-114">Se voc? ainda n?o tem uma assinatura do Azure, pode [obter uma como parte da sua assinatura do MSDN](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) ou [registrar-se gratuitamente para uma avalia??o gratuita](https://azure.microsoft.com/en-us/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="1a626-114">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/en-us/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/en-us/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="1a626-115">Etapa 1: criar uma pasta compartilhada para hospedar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="1a626-115">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="1a626-116">Abra o Explorador de Arquivos em seu computador de desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="1a626-116">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="1a626-117">Clique com o bot?o direito do mouse na unidade C:\ e escolha **Novo** > **Pasta**.</span><span class="sxs-lookup"><span data-stu-id="1a626-117">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="1a626-118">Nomeie a nova pasta AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="1a626-118">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="1a626-119">Clique com o bot?o direito do mouse na pasta AddinManifests e escolha **Compartilhar com** > **Pessoas espec?ficas**.</span><span class="sxs-lookup"><span data-stu-id="1a626-119">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="1a626-120">Em **Compartilhamento de Arquivos**, selecione a seta suspensa e escolha **Todos** > **Adicionar** > **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="1a626-120">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="1a626-p102">Nesta explica??o passo a passo, voc? est? usando um compartilhamento de arquivos local como um cat?logo confi?vel onde armazenar? o arquivo de manifesto XML do suplemento. Em um cen?rio real, em vez disso, ? poss?vel optar por [implantar o arquivo de manifesto XML a um cat?logo do SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) ou [publicar o suplemento no AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="1a626-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/en-us/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="1a626-123">Etapa 2:adicionar o compartilhamento de arquivos ao cat?logo de suplementos confi?veis</span><span class="sxs-lookup"><span data-stu-id="1a626-123">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="1a626-124">Inicie o Word 2016 e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="1a626-124">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a626-125">Embora este exemplo use o Word 2016, ? poss?vel usar qualquer aplicativo do Office que d? suporte a Suplementos do Office, como Excel, Outlook, PowerPoint ou Project 2016.</span><span class="sxs-lookup"><span data-stu-id="1a626-125">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="1a626-126">Escolha **Arquivo**  >  **Op??es**.</span><span class="sxs-lookup"><span data-stu-id="1a626-126">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="1a626-127">Na caixa de di?logo **Op??es do Word**, escolha **Central de Confiabilidade**, depois **Configura??es da Central de Confiabilidade**.</span><span class="sxs-lookup"><span data-stu-id="1a626-127">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="1a626-p103">Na caixa de di?logo **Central de Confiabilidade**, escolha **Cat?logos de Suplementos Confi?veis**. Digite o caminho UNC (conven??o universal de nomenclatura) para o compartilhamento de arquivos que voc? criou anteriormente como a **URL do Cat?logo**. Por exemplo, \\\NomedoseuComputador\AddinManifests. Em seguida, escolha **Adicionar cat?logo**.</span><span class="sxs-lookup"><span data-stu-id="1a626-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="1a626-130">Marque a caixa de sele??o **Mostrar no Menu**.</span><span class="sxs-lookup"><span data-stu-id="1a626-130">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="1a626-131">Ao armazenar um arquivo de manifesto XML de suplemento em um compartilhamento especificado como um cat?logo de suplementos da Web confi?vel, o suplemento aparece em **Pasta Compartilhada** na caixa de di?logo **Suplementos do Office** quando o usu?rio navega at? a guia **Inserir** na faixa de op??es e escolhe **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1a626-131">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="1a626-132">Feche o Word 2016.</span><span class="sxs-lookup"><span data-stu-id="1a626-132">Close Word 2016.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="1a626-133">Etapa 3: criar um aplicativo Web no Azure</span><span class="sxs-lookup"><span data-stu-id="1a626-133">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="1a626-134">Crie um aplicativo Web vazio no Azure usando [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) ou o [portal do Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="1a626-134">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="1a626-135">Usar o Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="1a626-135">Using Visual Studio 2017</span></span>

<span data-ttu-id="1a626-136">Para criar o aplicativo Web usando o Visual Studio 2017, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="1a626-136">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="1a626-p104">No Visual Studio, no menu **Exibir**, escolha **Gerenciador de Servidores**. Clique com o bot?o direito do mouse em **Azure** e escolha **Conectar-se ? assinatura do Microsoft Azure**. Siga as instru??es para se conectar ? sua assinatura do Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="1a626-140">No Visual Studio, no **Gerenciador de Servidores**, expanda **Azure**, clique com o bot?o direito do mouse em **Servi?o de Aplicativo** e escolha **Criar novo aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="1a626-140">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="1a626-141">Na caixa de di?logo **Criar Servi?o de Aplicativo**, forne?a estas informa??es:</span><span class="sxs-lookup"><span data-stu-id="1a626-141">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="1a626-p105">Insira um **Nome do Aplicativo Web** exclusivo para seu site. O Azure verifica se o nome do site ? exclusivo em todo o dom?nio azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="1a626-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="1a626-144">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="1a626-144">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="1a626-p106">Escolha o **Grupo de Recursos** para seu site. Se voc? criar um novo grupo, tamb?m precisar? dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="1a626-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="1a626-p107">Escolha o **Plano do Servi?o de Aplicativo** a ser usado para criar esse site. Se voc? criar um novo plano, tamb?m precisar? dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="1a626-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="1a626-149">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="1a626-149">Choose **Create**.</span></span>

    <span data-ttu-id="1a626-150">O novo aplicativo Web aparece no **Gerenciador de Servidores** em **Azure** >> **Servi?o de Aplicativo** >> (o grupo de recursos escolhido).</span><span class="sxs-lookup"><span data-stu-id="1a626-150">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="1a626-p108">Clique com o bot?o direito do mouse no novo aplicativo Web e escolha **Exibir no Navegador**. O navegador ser? aberto e exibir? uma p?gina da Web com a mensagem "Seu aplicativo de Servi?o de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="1a626-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="1a626-153">Na barra de endere?os do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="1a626-153">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="1a626-154"> Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="1a626-154">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="1a626-155">Usar o portal do Azure</span><span class="sxs-lookup"><span data-stu-id="1a626-155">Using the Azure portal</span></span>

<span data-ttu-id="1a626-156">Para criar o aplicativo Web usando o portal do Azure, realize as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="1a626-156">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="1a626-157">Fa?a logon no [portal do Azure](https://portal.azure.com/) usando suas credenciais do Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-157">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="1a626-158">Escolha **Novo** > **Web + Celular** > **Aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="1a626-158">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="1a626-159">Na caixa de di?logo **Criar Aplicativo Web**, forne?a estas informa??es:</span><span class="sxs-lookup"><span data-stu-id="1a626-159">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="1a626-p109">Insira um **Nome de aplicativo** exclusivo para seu site. O Azure verifica se o nome do site ? exclusivo em todo o dom?nio azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="1a626-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="1a626-162">Escolha a **Assinatura** a ser usada para criar esse site.</span><span class="sxs-lookup"><span data-stu-id="1a626-162">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="1a626-p110">Escolha o **Grupo de Recursos** para seu site. Se voc? criar um novo grupo, tamb?m precisar? dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="1a626-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="1a626-165">Escolha o **SO** para seu site.</span><span class="sxs-lookup"><span data-stu-id="1a626-165">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="1a626-p111">Escolha o **Plano do Servi?o de Aplicativo** a ser usado para criar esse site. Se voc? criar um novo plano, tamb?m precisar? dar um nome a ele.</span><span class="sxs-lookup"><span data-stu-id="1a626-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="1a626-168">Escolha **Criar**.</span><span class="sxs-lookup"><span data-stu-id="1a626-168">Choose **Create**.</span></span>

4. <span data-ttu-id="1a626-169">Escolha **Notifica??es** (o ?cone de sino localizado na borda superior do portal do Azure) e, em seguida, escolha a notifica??o **Implanta??es bem-sucedidas** para abrir a p?gina **Vis?o geral** no portal do Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-169">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a626-170">A notifica??o ser? alterada de **Implanta??o em andamento** para **Implanta??es bem-sucedidas** quando a implanta??o do site for conclu?da.</span><span class="sxs-lookup"><span data-stu-id="1a626-170">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="1a626-p112">Na se??o **Fundamentos** da p?gina **Vis?o geral** do site no portal do Azure, escolha a URL exibida em **URL**. O navegador ser? aberto e exibir? uma p?gina da Web com a mensagem "Seu aplicativo de Servi?o de Aplicativo foi criado".</span><span class="sxs-lookup"><span data-stu-id="1a626-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="1a626-173">Na barra de endere?os do navegador, altere a URL do aplicativo Web para que ela use HTTPS e pressione **Enter** para confirmar se o protocolo HTTPS foi habilitado.</span><span class="sxs-lookup"><span data-stu-id="1a626-173">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="1a626-174"> Os sites do Azure fornecem automaticamente um ponto de extremidade HTTPS.</span><span class="sxs-lookup"><span data-stu-id="1a626-174">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="1a626-175">Etapa 4: criar um suplemento do Office no Visual Studio</span><span class="sxs-lookup"><span data-stu-id="1a626-175">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="1a626-176">Inicie o Visual Studio como um administrador.</span><span class="sxs-lookup"><span data-stu-id="1a626-176">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="1a626-177">Escolha **Arquivo** > **Novo** > **Projeto**.</span><span class="sxs-lookup"><span data-stu-id="1a626-177">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="1a626-178">Em **Modelos**, expanda **Visual C#** (ou **Visual Basic**), expanda **Office/SharePoint** e escolha **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1a626-178">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="1a626-179">Escolha **Suplemento da Web do Word** e escolha **OK** para aceitar as configura??es padr?o.</span><span class="sxs-lookup"><span data-stu-id="1a626-179">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="1a626-180">O Visual Studio cria um suplemento b?sico do Word que voc? pode publicar como est?, sem fazer altera??es no projeto da Web.</span><span class="sxs-lookup"><span data-stu-id="1a626-180">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="1a626-181">Etapa 5: publicar seu aplicativo Web do suplemento do Office no Azure</span><span class="sxs-lookup"><span data-stu-id="1a626-181">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="1a626-182">Com seu projeto de suplemento aberto no Visual Studio, expanda o n? da solu??o no **Gerenciador de Solu??es** a fim de ver ambos os projetos para a solu??o.</span><span class="sxs-lookup"><span data-stu-id="1a626-182">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="1a626-p113">Clique com bot?o direito do mouse no projeto da Web e escolha **Publicar**. O projeto da Web cont?m arquivos do aplicativo Web do suplemento do Office, portanto, esse ? o projeto que voc? publica no Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="1a626-185">Na guia **Publicar**:</span><span class="sxs-lookup"><span data-stu-id="1a626-185">On the **Publish** tab:</span></span>

      - <span data-ttu-id="1a626-186">Escolha **Servi?o de Aplicativo do Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="1a626-186">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="1a626-187">Escolha **Selecionar Existentes**.</span><span class="sxs-lookup"><span data-stu-id="1a626-187">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="1a626-188">Escolha **Publicar**.</span><span class="sxs-lookup"><span data-stu-id="1a626-188">Choose **Publish**.</span></span> 

6. <span data-ttu-id="1a626-189">Na caixa de di?logo **Servi?o de Aplicativo**, localize e escolha o aplicativo Web que voc? criou na [Etapa 3: criar um aplicativo Web no Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) e, em seguida, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="1a626-189">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="1a626-p114">O Visual Studio publica o projeto da Web de seu Suplemento do Office no seu aplicativo Web do Azure. Quando o Visual Studio terminar de publicar o projeto da Web, o navegador abrir? e mostrar? uma p?gina da Web com o texto "Seu aplicativo de Servi?o de Aplicativo foi criado." Esta ? a p?gina padr?o atual do aplicativo Web.</span><span class="sxs-lookup"><span data-stu-id="1a626-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="1a626-193">Para ver a p?gina da Web do seu suplemento, altere o URL para que ele use HTTPS e especifique o caminho da p?gina HTML do seu suplemento (por exemplo: https://YourDomain.azurewebsites.net/Home.html).</span><span class="sxs-lookup"><span data-stu-id="1a626-193">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="1a626-194">Isso confirma que o aplicativo Web do seu suplemento est? hospedado no Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-194">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="1a626-195">Copie o URL raiz (por exemplo: https://YourDomain.azurewebsites.net); voc? precisar? dele ao editar o arquivo de manifesto do suplemento mais adiante neste artigo.</span><span class="sxs-lookup"><span data-stu-id="1a626-195">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="1a626-196">Etapa 6: editar e implantar o arquivo de manifesto XML do suplemento</span><span class="sxs-lookup"><span data-stu-id="1a626-196">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="1a626-197">No Visual Studio, com o suplemento do Office de exemplo aberto no **Gerenciador de Solu??es**, expanda a solu??o para que ambos os projetos sejam exibidos.</span><span class="sxs-lookup"><span data-stu-id="1a626-197">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="1a626-p116">Expanda o projeto do Suplemento do Office (por exemplo, WordWebAddIn), clique com o bot?o direito do mouse na pasta do manifesto e escolha **Abrir**. O arquivo de manifesto XML do suplemento ? aberto.</span><span class="sxs-lookup"><span data-stu-id="1a626-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="1a626-200">No arquivo de manifesto XML, localize e substitua todas as inst?ncias de "~ remoteAppUrl" pelo URL raiz do aplicativo Web do suplemento no Azure.</span><span class="sxs-lookup"><span data-stu-id="1a626-200">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="1a626-201">Esse ? o URL que voc? copiou anteriormente depois de publicar o aplicativo Web do suplemento no Azure (por exemplo: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="1a626-201">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="1a626-p118">Escolha **Arquivo** e **Salvar tudo**. Feche o arquivo de manifesto XML do suplemento.</span><span class="sxs-lookup"><span data-stu-id="1a626-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="1a626-204">No **Gerenciador de Solu??es**, clique com o bot?o direito do mouse na pasta do manifesto e escolha **Abrir Pasta no Gerenciador de Arquivos**.</span><span class="sxs-lookup"><span data-stu-id="1a626-204">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="1a626-205">Copie o arquivo de manifesto XML do suplemento (por exemplo, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="1a626-205">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="1a626-206">Navegue at? o compartilhamento de arquivos de rede que voc? criou na [Etapa 1: criar uma pasta compartilhada](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) e cole o arquivo de manifesto na pasta.</span><span class="sxs-lookup"><span data-stu-id="1a626-206">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="1a626-207">Etapa 7: inserir e executar o suplemento no aplicativo cliente do Office</span><span class="sxs-lookup"><span data-stu-id="1a626-207">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="1a626-208">Inicie o Word 2016 e crie um documento.</span><span class="sxs-lookup"><span data-stu-id="1a626-208">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="1a626-209">Na faixa de op??es, escolha **Inserir** > **Meus Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="1a626-209">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="1a626-p119">Na caixa de di?logo **Suplementos do Office**, escolha **PASTA COMPARTILHADA**. O Word examina a pasta listada como um cat?logo de suplementos confi?veis (na [Etapa 2: adicionar o compartilhamento de arquivos ao cat?logo de suplementos confi?veis](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) e mostre os suplementos na caixa de di?logo. Voc? deve ver um ?cone de seu suplemento de exemplo.</span><span class="sxs-lookup"><span data-stu-id="1a626-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="1a626-p120">Escolha o ?cone para seu suplemento e escolha **Adicionar**. Um bot?o **Mostrar Painel de Tarefas** para seu suplemento ? adicionado ? faixa de op??es.</span><span class="sxs-lookup"><span data-stu-id="1a626-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="1a626-p121">Na faixa de op??es da guia **P?gina Inicial**, escolha o bot?o **Mostrar Painel de Tarefas**. O suplemento ? aberto em um painel de tarefas ? direita do documento atual.</span><span class="sxs-lookup"><span data-stu-id="1a626-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="1a626-p122">Para verificar se o suplemento funciona, selecione algum texto no documento e escolha o bot?o **Real?ar!** no painel de tarefas.</span><span class="sxs-lookup"><span data-stu-id="1a626-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="1a626-219">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="1a626-219">See also</span></span>

- [<span data-ttu-id="1a626-220">Publicar seu Suplemento do Office</span><span class="sxs-lookup"><span data-stu-id="1a626-220">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="1a626-221">Empacotar seu suplemento usando o Visual Studio para preparar a publica??o</span><span class="sxs-lookup"><span data-stu-id="1a626-221">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
