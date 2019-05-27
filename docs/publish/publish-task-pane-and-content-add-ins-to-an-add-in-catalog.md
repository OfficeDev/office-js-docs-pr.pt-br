---
title: Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint
description: Para disponibilizar os Suplementos do Office para os usuários na organização, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em nas organizações deles.
ms.date: 05/22/2019
localization_priority: Priority
ms.openlocfilehash: bffbf3e83a2e6d8d0c63252c27ba54826611f78b
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432240"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="deaa4-103">Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint</span><span class="sxs-lookup"><span data-stu-id="deaa4-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="deaa4-p101">Um catálogo de suplementos é um conjunto de sites dedicado em um aplicativo Web do SharePoint ou em locatário do SharePoint Online que hospeda bibliotecas de documentos para Suplementos do SharePoint e do Office. Para disponibilizar suplementos do Office nas empresas, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em suas organizações. Quando um administrador registra um catálogo de suplementos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface de usuário em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="deaa4-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="deaa4-106">Os catálogos de suplementos no SharePoint não são compatíveis com recursos de suplementos implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="deaa4-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="deaa4-107">Se você está direcionando para um ambiente híbrido ou de nuvem, recomendamos [usar a Implantação Centralizada por meio do Centro de Administração do Office 365](../publish/centralized-deployment.md) para publicar os suplementos.</span><span class="sxs-lookup"><span data-stu-id="deaa4-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="deaa4-108">Catálogos do SharePoint não são compatíveis com o Office para Mac.</span><span class="sxs-lookup"><span data-stu-id="deaa4-108">SharePoint catalogs are not supported for Office for Mac.</span></span> <span data-ttu-id="deaa4-109">Para implantar Suplementos do Office em clientes do Mac, envie-os para a [AppSource](/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="deaa4-109">To deploy Office Add-ins to Mac clients, you must submit them to [AppSource](/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="create-an-add-in-catalog"></a><span data-ttu-id="deaa4-110">Criação de um catálogo de suplementos</span><span class="sxs-lookup"><span data-stu-id="deaa4-110">Create an add-in catalog</span></span>

<span data-ttu-id="deaa4-111">Conclua as etapas em uma das seções a seguir para criar um catálogo de suplementos no SharePoint ou no Office 365.</span><span class="sxs-lookup"><span data-stu-id="deaa4-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-create-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="deaa4-112">Criar um catálogo de suplementos no SharePoint local.</span><span class="sxs-lookup"><span data-stu-id="deaa4-112">To set up an add-in catalog for on-premises SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="deaa4-113">A IU no SharePoint local ainda se refere aos suplementos como **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="deaa4-114">Acesse o **Site da Administração Central**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-114">Browse to the  **Central Administration Site**.</span></span>

2. <span data-ttu-id="deaa4-115">No painel de tarefas à esquerda, escolha os  **Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-115">In the left task pane, choose  **Apps**.</span></span>

3. <span data-ttu-id="deaa4-116">Na página**Aplicativos**, em **Gerenciamento de Aplicativos**, escolha  **Gerenciar Catálogo de Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>

4. <span data-ttu-id="deaa4-117">Na página**Gerenciar Catálogo de Aplicativos**, verifique se você tem o aplicativo web correto selecionado no **Seletor de Aplicativo Web**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-117">On the  **Manage App Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>

5. <span data-ttu-id="deaa4-118">Escolha  **Exibir configurações do site**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-118">Choose  **View site settings**.</span></span>

6. <span data-ttu-id="deaa4-119">Na página **Configurações do Site**, escolha **Administradores de conjunto de sites** para especificar os administradores de conjunto de sites e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>

7. <span data-ttu-id="deaa4-120">Para conceder permissões de site aos usuários, escolha **Permissões de Site** e **Conceder Permissões**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>

8. <span data-ttu-id="deaa4-121">Na caixa de diálogo **Compartilhar "Site do Catálogo de Aplicativos"**, especifique um ou mais usuários do site, defina as permissões apropriadas, defina outras opções se for o caso e escolha **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>

9. <span data-ttu-id="deaa4-122">Para adicionar suplementos ao catálogo de Suplementos do Office, escolha **Aplicativos do Office**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-122">To add an add-in to the Office Add-ins add-in catalog, choose **Apps for Office**.</span></span>

### <a name="to-create-an-app-catalog-on-office-365"></a><span data-ttu-id="deaa4-123">Criar um catálogo de aplicativos no Office 365</span><span class="sxs-lookup"><span data-stu-id="deaa4-123">To create an app catalog on Office 365</span></span>

<span data-ttu-id="deaa4-124">Mesmo que o SharePoint nomeie um catálogo de "aplicativo", é possível registrar os Suplementos do Office no catálogo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-124">Even though SharePoint names the catalog an "app" catalog, you can register Office Add-ins in the app catalog.</span></span>

1. <span data-ttu-id="deaa4-125">Vá para o centro de administração do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="deaa4-125">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="deaa4-126">Para saber mais sobre como encontrar o centro de administração, confira [Centro de administração do Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="deaa4-126">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>

2. <span data-ttu-id="deaa4-127">Na página do centro de administração do Microsoft 365, expanda a lista dos **Centros de administração** e selecione**SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-127">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="deaa4-128">Use o centro de administração do SharePoint Clássico para criar o catálogo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-128">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="deaa4-129">Se você estiver no novo centro de administração do SharePoint, escolha **Centro de administração do SharePoint clássico** no painel esquerdo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-129">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>

3. <span data-ttu-id="deaa4-130">No painel de tarefas à esquerda, escolha  **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-130">In the left task pane, choose  **Apps**.</span></span>

4. <span data-ttu-id="deaa4-131">Na página **aplicativos**, escolha **Catálogo de Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-131">On the **apps** page, select **App Catalog**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="deaa4-132">Se um catálogo de aplicativos já foi criado e exibido nesta página, você poderá ignorar o restante dessas etapas e ir para a próxima seção deste artigo para publicar o suplemento no catálogo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-132">If an app catalog is already created and appears on this page, then you can skip the rest of these steps and go to the next section of this article to publish your add-in to the catalog.</span></span>

5. <span data-ttu-id="deaa4-133">Na página **Site do Catálogo de Aplicativo**, escolha **OK** para aceitar a opção padrão e criar um novo site de catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="deaa4-133">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>

6. <span data-ttu-id="deaa4-134">Na página **Criar Conjunto de Sites do Catálogo de Aplicativos**, especifique o título do seu site de Catálogo de Aplicativos.</span><span class="sxs-lookup"><span data-stu-id="deaa4-134">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>

7. <span data-ttu-id="deaa4-135">Especifique o **Endereço do site da Web**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-135">Specify the web site address.</span></span>

8. <span data-ttu-id="deaa4-136">Especifique um **administrador**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-136">Specify an **Administrator**.</span></span>

9. <span data-ttu-id="deaa4-137">Defina a **Cota de Recursos de Servidor** como 0 (zero).</span><span class="sxs-lookup"><span data-stu-id="deaa4-137">Set the  **Server Resource Quota** to 0 (zero).</span></span> <span data-ttu-id="deaa4-138">(A cota de recursos de servidor está relacionada à limitação das soluções de área restrita com mau desempenho, mas não instala soluções de área restrita no seu site de catálogo de aplicativos.)</span><span class="sxs-lookup"><span data-stu-id="deaa4-138">(The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>

10. <span data-ttu-id="deaa4-139">Escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-139">Choose **OK**.</span></span>

<span data-ttu-id="deaa4-140">O catálogo de aplicativos foi criado.</span><span class="sxs-lookup"><span data-stu-id="deaa4-140">The app catalog is now created.</span></span>

## <a name="publish-an-add-in-to-an-app-catalog"></a><span data-ttu-id="deaa4-141">Publicar um suplemento em um catálogo de aplicativos</span><span class="sxs-lookup"><span data-stu-id="deaa4-141">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="deaa4-142">Para publicar um suplemento em um catálogo de aplicativo existente, conclua as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="deaa4-142">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="deaa4-143">Vá para o centro de administração do Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="deaa4-143">Go to the Microsoft 365 admin center.</span></span> <span data-ttu-id="deaa4-144">Para saber mais sobre como encontrar o centro de administração, confira [Centro de administração do Microsoft 365](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="deaa4-144">For information on how to find the admin center, see [About the Microsoft 365 admin center](https://docs.microsoft.com/office365/admin/admin-overview/about-the-admin-center).</span></span>
2. <span data-ttu-id="deaa4-145">Na página do centro de administração do Microsoft 365, expanda a lista dos **Centros de administração** e selecione**SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-145">On the Microsoft 365 admin center page, expand the list of **Admin centers**, and then choose **SharePoint**.</span></span>
    > [!NOTE]
    > <span data-ttu-id="deaa4-146">Use o centro de administração do SharePoint Clássico para criar o catálogo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-146">You need to use the Classic SharePoint admin center to create the catalog.</span></span> <span data-ttu-id="deaa4-147">Se você estiver no novo centro de administração do SharePoint, escolha **Centro de administração do SharePoint clássico** no painel esquerdo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-147">If you are in the new SharePoint admin center, choose **Classic SharePoint admin center** in the left pane.</span></span>
3. <span data-ttu-id="deaa4-148">No painel de tarefas à esquerda, escolha  **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-148">In the left task pane, choose  **Apps**.</span></span>
4. <span data-ttu-id="deaa4-149">Na página **aplicativos**, escolha **Catálogo de Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-149">On the **apps** page, select **App Catalog**.</span></span>
5. <span data-ttu-id="deaa4-150">Escolha **Distribuir aplicativos para o Office**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-150">Choose **Distribute apps for Office**.</span></span>
6. <span data-ttu-id="deaa4-151">Na página **Aplicativos do Office**, escolha **Novo**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-151">In the **Apps for Office** page, choose **New**.</span></span>
7. <span data-ttu-id="deaa4-152">Na caixa de diálogo **Adicionar um documento**, selecione o botão **Escolher Arquivos**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-152">In the **Add a document** dialog, select the **Choose Files** button.</span></span>
8. <span data-ttu-id="deaa4-153">Localize e especifique o arquivo [manifesto](../develop/add-in-manifests.md) para carregar e escolha **Abrir**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-153">Locate and specify the [manifest](../develop/add-in-manifests.md) file to upload and choose **Open**.</span></span>
9. <span data-ttu-id="deaa4-154">Na caixa de diálogo **Adicionar um documento**, escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-154">In the **Add a document** dialog box, choose **OK**.</span></span>

    <span data-ttu-id="deaa4-p108">Suplementos de conteúdo e de painel de tarefas neste catálogo agora ficam disponíveis na caixa de diálogo **Suplementos do Office**. Para acessá-los, escolha **Meus Suplementos** na guia **Inserir** e, em seguida, escolha **MINHA ORGANIZAÇÃO**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-p108">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="deaa4-157">Experiência do usuário final com o catálogo de suplementos</span><span class="sxs-lookup"><span data-stu-id="deaa4-157">End user experience with the add-in catalog</span></span>

<span data-ttu-id="deaa4-158">Os usuários finais podem acessar o catálogo de suplementos em um aplicativo do Office realizando as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="deaa4-158">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="deaa4-159">Em um aplicativo do Office, vá para **Arquivo** > **Opções** > **Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Suplementos Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-159">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>

2. <span data-ttu-id="deaa4-160">Especifique a URL do _conjunto de sites do SharePoint pai_ do catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="deaa4-160">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 

    <span data-ttu-id="deaa4-161">Por exemplo, se a URL do catálogo de Suplementos do Office é:</span><span class="sxs-lookup"><span data-stu-id="deaa4-161">For example, if the URL of the Office Add-ins catalog is:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`

    <span data-ttu-id="deaa4-162">Especifique somente a URL do conjunto de sites pai:</span><span class="sxs-lookup"><span data-stu-id="deaa4-162">Specify just the URL of the parent site collection:</span></span>

    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`

3. <span data-ttu-id="deaa4-p109">Feche e reabra o aplicativo do Office. O catálogo de suplementos estará disponível na caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="deaa4-p109">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="deaa4-165">Como alternativa, um administrador pode especificar um catálogo de Suplementos do Office no SharePoint usando as políticas de grupo.</span><span class="sxs-lookup"><span data-stu-id="deaa4-165">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="deaa4-166">Para saber mais, veja a seção [Usar uma Política de Grupo para gerenciar como os usuários podem instalar e usar os Suplementos do Office](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="deaa4-166">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span></span>
