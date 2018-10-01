---
title: Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint
description: Para tornar os suplementos do Office acessíveis aos usuários da organização, os administradores podem carregar arquivos de manifesto de suplementos do Office para o catálogo de suplementos da sua organização.
ms.date: 01/23/2018
ms.openlocfilehash: 5ba6a54c4540f79c65082cd7de3b76f300831341
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348118"
---
# <a name="publish-task-pane-and-content-add-ins-to-a-sharepoint-catalog"></a><span data-ttu-id="6da1f-103">Publicar suplementos de conteúdo e de painel de tarefas em um catálogo do SharePoint</span><span class="sxs-lookup"><span data-stu-id="6da1f-103">Publish task pane and content add-ins to a SharePoint catalog</span></span>

<span data-ttu-id="6da1f-p101">Um catálogo de suplementos é um conjunto de sites dedicado em um aplicativo Web do SharePoint ou em locatário do SharePoint Online que hospeda bibliotecas de documentos para Suplementos do SharePoint e do Office. Para disponibilizar suplementos do Office nas empresas, os administradores podem carregar arquivos de manifesto de Suplementos do Office no catálogo de suplementos para uso em suas organizações. Quando um administrador registra um catálogo de suplementos como um catálogo confiável, os usuários podem inserir o suplemento a partir da interface de usuário em um aplicativo cliente do Office.</span><span class="sxs-lookup"><span data-stu-id="6da1f-p101">An add-in catalog is a dedicated site collection in a SharePoint web application or SharePoint Online tenancy that hosts document libraries for Office and SharePoint Add-ins. To make Office Add-ins accessible to users within their organization, administrators can upload Office Add-ins manifest files to the add-in catalog for their organization. When an administrator registers an add-in catalog as a trusted catalog, users can insert the add-in from the insertion UI in an Office client application.</span></span>

> [!IMPORTANT]
> - <span data-ttu-id="6da1f-106">Os catálogos de suplementos no SharePoint não são compatíveis com recursos de suplementos implementados no nó `VersionOverrides` do [manifesto do suplemento](../develop/add-in-manifests.md), como comandos de suplemento.</span><span class="sxs-lookup"><span data-stu-id="6da1f-106">Add-in catalogs on SharePoint do not support add-in features that are implemented in the `VersionOverrides` node of the [add-in manifest](../develop/add-in-manifests.md), such as add-in commands.</span></span>
> - <span data-ttu-id="6da1f-107">Se você está direcionando para um ambiente híbrido ou de nuvem, recomendamos [usar a Implantação Centralizada por meio do Centro de Administração do Office 365](../publish/centralized-deployment.md) para publicar os suplementos.</span><span class="sxs-lookup"><span data-stu-id="6da1f-107">If you’re targeting a cloud or hybrid environment, we recommend that you [use Centralized Deployment via the Office 365 admin center](../publish/centralized-deployment.md) to publish your add-ins.</span></span>
> - <span data-ttu-id="6da1f-108">Catálogos do SharePoint não são compatíveis com o Office para Mac.</span><span class="sxs-lookup"><span data-stu-id="6da1f-108">SharePoint catalogs are not supported for Office 2016 for Mac.</span></span> <span data-ttu-id="6da1f-109">Para implantar suplementos do Office para clientes Mac, você deve enviá-los para o [AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="6da1f-109">To deploy Office Add-ins to Mac clients, you must submit them to the [Office Store](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>   

## <a name="set-up-an-add-in-catalog"></a><span data-ttu-id="6da1f-110">Configurar um catálogo de suplementos</span><span class="sxs-lookup"><span data-stu-id="6da1f-110">Set up an add-in catalog</span></span>

<span data-ttu-id="6da1f-111">Conclua as etapas em uma das seções a seguir para configurar um catálogo de suplementos no SharePoint ou no Office 365.</span><span class="sxs-lookup"><span data-stu-id="6da1f-111">Complete the steps in one of the following sections to set up an add-in catalog on SharePoint or on Office 365.</span></span>

### <a name="to-set-up-an-add-in-catalog-for-on-premises-sharepoint"></a><span data-ttu-id="6da1f-112">Para configurar um catálogo de suplementos do SharePoint local</span><span class="sxs-lookup"><span data-stu-id="6da1f-112">To set up an add-in catalog on SharePoint</span></span>

> [!NOTE]
> <span data-ttu-id="6da1f-113">A interface do usuário no SharePoint local ainda se refere aos suplementos como **aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-113">The UI in on-premises SharePoint still refers to add-ins as **apps**.</span></span>

1. <span data-ttu-id="6da1f-114">Navegue até o **Site da Administração Central**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-114">Browse to the  **Central Administration Site**.</span></span>
    
2. <span data-ttu-id="6da1f-115">No painel de tarefas esquerdo, escolha **Aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-115">In the left task pane, choose  **add-ins**.</span></span>
    
3. <span data-ttu-id="6da1f-116">Na página **Aplicativos**, em **Gerenciamento de aplicativos**, escolha **Gerenciar o catálogo de aplicativos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-116">On the  **Apps** page, under **App Management**, choose  **Manage App Catalog**.</span></span>
    
4. <span data-ttu-id="6da1f-117">Na página\*\* Gerenciar Catálogo de Suplementos\*\*, verifique se você tem o aplicativo da Web correto selecionado no **Seletor de Aplicativo da Web**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-117">On the  **Manage Add-in Catalog** page, make sure you have the right web application selected in the **Web Application Selector**.</span></span>
    
5. <span data-ttu-id="6da1f-118">Escolha **Exibir configurações do site**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-118">Choose  **View site settings**.</span></span>
    
6. <span data-ttu-id="6da1f-119">Na página **Configurações do Site**, escolha **Administradores de conjunto de sites** para especificar os administradores de conjunto de sites e escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-119">On the  **Site Settings** page, choose **Site collection administrators** to specify the site collection administrators, and then choose **OK**.</span></span>
    
7. <span data-ttu-id="6da1f-120">Para conceder permissões de site aos usuários, escolha **Permissões de Site** e **Conceder Permissões**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-120">To grant site permissions to users, choose  **Site Permissions**, and then choose  **Grant Permissions**.</span></span>
    
8. <span data-ttu-id="6da1f-121">Na caixa de diálogo **Compartilhar "Site do Catálogo de Aplicativos"**, especifique um ou mais usuários do site, defina as permissões apropriadas, defina outras opções, se for o caso, e escolha **Compartilhar**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-121">In the  **Share 'App Catalog Site'** dialog box, specify one or more site users, set the appropriate permissions for them, optionally set other options, and then choose **Share**.</span></span>
    
9. <span data-ttu-id="6da1f-122">Para adicionar suplementos ao catálogo de Suplementos do Office, escolha **Aplicativos do Office**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-122">To add an add-in to the Office Add-ins add-in catalog, choose **Office Add-ins**.</span></span>

### <a name="to-set-up-an-add-in-catalog-on-office-365"></a><span data-ttu-id="6da1f-123">Para configurar um catálogo de suplementos no Office 365</span><span class="sxs-lookup"><span data-stu-id="6da1f-123">To set up an add-in catalog on Office 365</span></span>

1. <span data-ttu-id="6da1f-124">Na página do Centro de Administração do Office 365, escolha **Administrador** e **SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-124">On the Office 365 admin center page, choose  **Admin**, and then choose  **SharePoint**.</span></span>
    
2. <span data-ttu-id="6da1f-125">No painel de tarefas à esquerda, escolha **suplementos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-125">In the left task pane, choose  **add-ins**.</span></span>
    
3. <span data-ttu-id="6da1f-126">Na página **suplementos**, escolha **Catálogo de Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-126">On the  **add-ins** page, choose **Add-in Catalog**.</span></span>
    
4. <span data-ttu-id="6da1f-127">Na página **Site do Catálogo de Suplementos**, escolha **OK** para aceitar a opção padrão e criar um novo site de catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="6da1f-127">On the  **Add-in Catalog Site** page, choose **OK** to accept the default option and create a new add-in catalog site.</span></span>
    
5. <span data-ttu-id="6da1f-128">Na página **Criar Conjunto de Sites de Catálogo de Suplementos**, especifique o título do seu site de Catálogo de Suplementos.</span><span class="sxs-lookup"><span data-stu-id="6da1f-128">On the  **Create Add-in Catalog Site Collection** page, specify the title of your Add-in Catalog site.</span></span>
    
6. <span data-ttu-id="6da1f-129">Especifique o endereço do site da Web.</span><span class="sxs-lookup"><span data-stu-id="6da1f-129">Specify the web site address.</span></span>
    
7. <span data-ttu-id="6da1f-p103">Defina a **cota de armazenamento** como o menor valor possível (atualmente 110). Você só instalará pacotes de suplementos neste conjunto de sites e eles são muito pequenos.</span><span class="sxs-lookup"><span data-stu-id="6da1f-p103">Set the  **Storage Quota** to the lowest possible value (currently 110). You will only be installing add-in packages on this site collection and they are very small.</span></span>
    
8. <span data-ttu-id="6da1f-p104">Defina a **Cota de Recursos de Servidor** como 0 (zero). (A cota de recursos de servidor está relacionada à limitação das soluções de área restrita com mau desempenho, mas você não vai instalar soluções de área restrita no seu site de catálogo de suplementos.)</span><span class="sxs-lookup"><span data-stu-id="6da1f-p104">Set the  **Server Resource Quota** to 0 (zero). (The server resource quota is related to throttling poorly performing sandboxed solutions, but you won't be installing any sandboxed solutions on your add-in catalog site.)</span></span>
    
9. <span data-ttu-id="6da1f-134">Escolha **OK**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-134">Choose  **OK**.</span></span>
    
10. <span data-ttu-id="6da1f-p105">Para adicionar um suplemento ao Site do Catálogo de Suplementos, navegue até o site que acabou de criar. No painel de navegação à esquerda, escolha **Suplementos do Office** e, para carregar um arquivo de manifesto do suplemento do Office, escolha **novo suplemento**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-p105">To add an add-in to the Add-in Catalog Site, browse to the site you have just created. In the left navigation pane, choose  **Office Add-ins**, and then, to upload an Office Add-in manifest file, choose  **new add-in**.</span></span>

## <a name="publish-an-add-in-to-an-add-in-catalog"></a><span data-ttu-id="6da1f-137">Publicar um suplemento em um catálogo de suplementos</span><span class="sxs-lookup"><span data-stu-id="6da1f-137">Publish an add-in to an add-in catalog</span></span>

<span data-ttu-id="6da1f-138">Para publicar um suplemento em um catálogo suplementos, conclua as etapas a seguir.</span><span class="sxs-lookup"><span data-stu-id="6da1f-138">To publish an add-in to an add-in catalog, complete the following steps.</span></span>

1. <span data-ttu-id="6da1f-139">Navegue até o catálogo de suplementos:</span><span class="sxs-lookup"><span data-stu-id="6da1f-139">Browse to the add-in catalog:</span></span>

    - <span data-ttu-id="6da1f-140">Abra a página principal da Administração Central do SharePoint.</span><span class="sxs-lookup"><span data-stu-id="6da1f-140">Open the SharePoint Central Administration main page.</span></span>
    
    - <span data-ttu-id="6da1f-141">Selecione **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-141">Select  **Add-ins**.</span></span>
    
    - <span data-ttu-id="6da1f-142">Selecione **Gerenciar Catálogo de Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-142">Select  **Manage Add-in Catalog**.</span></span>
    
    - <span data-ttu-id="6da1f-143">Escolha o link fornecido e escolha **Suplementos do Office** na barra de navegação à esquerda.</span><span class="sxs-lookup"><span data-stu-id="6da1f-143">Choose the link provided, and then choose  **Office Add-ins** on the left navigation bar.</span></span>
    
2. <span data-ttu-id="6da1f-144">Escolha o link **Clique para adicionar um novo item**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-144">Choose the  **Click to add new item** link.</span></span>
    
3. <span data-ttu-id="6da1f-145">Escolha **Procurar** e especifique o [manifesto](../develop/add-in-manifests.md) para carregar.</span><span class="sxs-lookup"><span data-stu-id="6da1f-145">Choose  **Browse**, and then specify the [manifest](../develop/add-in-manifests.md) to upload.</span></span>
    
    <span data-ttu-id="6da1f-p106">Suplementos de conteúdo e de painel de tarefas neste catálogo agora estão disponíveis na caixa de diálogo **Suplementos do Office**. Para acessá-los, escolha **Meus Suplementos** na guia **Inserir** e, em seguida, escolha **MINHA ORGANIZAÇÃO**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-p106">Content and task pane add-ins in this catalog are now available from the  **Office Add-ins** dialog box. To access them, choose **My Add-ins** on the **Insert** tab, and then choose **MY ORGANIZATION**.</span></span>

## <a name="end-user-experience-with-the-add-in-catalog"></a><span data-ttu-id="6da1f-148">Experiência do usuário final com o catálogo de suplementos</span><span class="sxs-lookup"><span data-stu-id="6da1f-148">End user experience with the add-in catalog</span></span>

<span data-ttu-id="6da1f-149">Os usuários finais podem acessar o catálogo de suplementos em um aplicativo do Office realizando as seguintes etapas:</span><span class="sxs-lookup"><span data-stu-id="6da1f-149">End users can access the add-in catalog in an Office application by completing the following steps:</span></span>

1. <span data-ttu-id="6da1f-150">Em um aplicativo do Office, vá para **Arquivo** > **Opções** > **Central de Confiabilidade** > **Configurações da Central de Confiabilidade** > **Catálogos de Suplementos Confiáveis**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-150">In the Office application, go to  **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Trusted Add-in Catalogs**.</span></span>
    
2. <span data-ttu-id="6da1f-151">Especifique a URL do _conjunto de sites do SharePoint pai_ do catálogo de suplementos.</span><span class="sxs-lookup"><span data-stu-id="6da1f-151">Specify the URL of the  _parent SharePoint site collection_ of the add-in catalog.</span></span> 
    
    <span data-ttu-id="6da1f-152">Por exemplo, se a URL do catálogo de Suplementos do Office for:</span><span class="sxs-lookup"><span data-stu-id="6da1f-152">For example, if the URL of the Office Add-ins catalog is:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_ /AgaveCatalog`
    
    <span data-ttu-id="6da1f-153">Especifique somente a URL do conjunto de sites pai:</span><span class="sxs-lookup"><span data-stu-id="6da1f-153">Specify just the URL of the parent site collection:</span></span>
    
    - `https:// _domain_ /sites/ _AddinCatalogSiteCollection_`
    
3. <span data-ttu-id="6da1f-p107">Feche e reabra o aplicativo do Office. O catálogo de suplementos estará disponível na caixa de diálogo **Suplementos do Office**.</span><span class="sxs-lookup"><span data-stu-id="6da1f-p107">Close and reopen the Office application. The add-in catalog will be available in the **Office Add-ins** dialog box.</span></span>

<span data-ttu-id="6da1f-156">Como alternativa, um administrador pode especificar um catálogo de Suplementos do Office no SharePoint usando políticas de grupo.</span><span class="sxs-lookup"><span data-stu-id="6da1f-156">Alternatively, an administrator can specify an Office Add-in catalog on SharePoint by using group policy.</span></span> <span data-ttu-id="6da1f-157">Confira mais detalhes na seção [Usar uma Política de Grupo para gerenciar como usuários podem instalar e usar Suplementos do Office](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office).</span><span class="sxs-lookup"><span data-stu-id="6da1f-157">For details, see the section [Using Group Policy to manage how users can install and use Office Add-ins](https://docs.microsoft.com/previous-versions/office/office-2013-resource-kit/jj219429(v=office.15)#using-group-policy-to-manage-how-users-can-install-and-use-apps-for-office) on TechNet.</span></span>
