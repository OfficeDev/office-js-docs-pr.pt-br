---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 04/17/2019
localization_priority: Priority
ms.openlocfilehash: 9a3d09a78a7644b3b26c345ba2588a1fae59c1eb
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914267"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="36c74-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="36c74-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="36c74-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="36c74-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="36c74-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="36c74-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="36c74-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="36c74-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="36c74-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="36c74-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="36c74-107">Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.</span><span class="sxs-lookup"><span data-stu-id="36c74-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="36c74-108">Você também precisará ingressar no [programa Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="36c74-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="36c74-109">O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="36c74-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="36c74-110">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="36c74-110">Features in preview</span></span>

<span data-ttu-id="36c74-111">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="36c74-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="36c74-112">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="36c74-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="36c74-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="36c74-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="36c74-114">Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="36c74-114">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="36c74-115">Esse valor é usado para cancelar a execução de um evento.</span><span class="sxs-lookup"><span data-stu-id="36c74-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="36c74-116">**Disponível em**: Outlook na web (clássico)</span><span class="sxs-lookup"><span data-stu-id="36c74-116">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="attachments"></a><span data-ttu-id="36c74-117">Attachments</span><span class="sxs-lookup"><span data-stu-id="36c74-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="36c74-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="36c74-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="36c74-119">Adicionado um novo objeto que representa o conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="36c74-119">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="36c74-120">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-120">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="36c74-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="36c74-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="36c74-122">Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="36c74-122">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="36c74-123">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-123">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="36c74-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="36c74-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="36c74-125">Adicionar um novo método para acessar o conteúdo de um anexo específico.</span><span class="sxs-lookup"><span data-stu-id="36c74-125">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="36c74-126">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-126">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="36c74-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="36c74-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="36c74-128">Adicionado um novo método que obtém um item anexo no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="36c74-128">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="36c74-129">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-129">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="36c74-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="36c74-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="36c74-131">Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="36c74-131">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="36c74-132">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-132">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="36c74-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="36c74-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="36c74-134">Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.</span><span class="sxs-lookup"><span data-stu-id="36c74-134">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="36c74-135">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-135">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="36c74-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="36c74-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="36c74-137">Adicionado `AttachmentsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="36c74-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="36c74-138">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-138">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="categories"></a><span data-ttu-id="36c74-139">Categorias</span><span class="sxs-lookup"><span data-stu-id="36c74-139">Categories</span></span>

<span data-ttu-id="36c74-140">No Outlook, um usuário pode agrupar mensagens e compromissos usando uma categoria para codificá-los por cor.</span><span class="sxs-lookup"><span data-stu-id="36c74-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="36c74-141">O usuário define as categorias em uma lista mestra em sua caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="36c74-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="36c74-142">Ele pode, em seguida, aplicar uma ou mais categorias a um item.</span><span class="sxs-lookup"><span data-stu-id="36c74-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="36c74-143">Não há suporte para esse recurso no Outlook para iOS ou Outlook para Android.</span><span class="sxs-lookup"><span data-stu-id="36c74-143">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="36c74-144">Categories</span><span class="sxs-lookup"><span data-stu-id="36c74-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="36c74-145">Adicionou um novo objeto que representa a categoria de um item.</span><span class="sxs-lookup"><span data-stu-id="36c74-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="36c74-146">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-146">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="36c74-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="36c74-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="36c74-148">Adicionou um novo objeto que representa os detalhes de uma categoria (seu nome e cor associada).</span><span class="sxs-lookup"><span data-stu-id="36c74-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="36c74-149">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-149">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="36c74-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="36c74-150">masterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="36c74-151">Adicionou um novo objeto que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="36c74-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="36c74-152">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-152">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="36c74-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="36c74-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="36c74-154">Adicionou uma nova propriedade que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="36c74-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="36c74-155">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-155">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="36c74-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="36c74-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="36c74-157">Adicionou uma nova propriedade que representa o conjunto de categorias em um item.</span><span class="sxs-lookup"><span data-stu-id="36c74-157">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="36c74-158">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-158">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="36c74-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="36c74-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="36c74-160">Adicionou uma nova enumeração que especifica as cores disponíveis a serem associadas a categorias. </span><span class="sxs-lookup"><span data-stu-id="36c74-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="36c74-161">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-161">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="36c74-162">Acesso de representante</span><span class="sxs-lookup"><span data-stu-id="36c74-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="36c74-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="36c74-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="36c74-164">Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="36c74-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="36c74-165">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-165">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="36c74-166">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="36c74-166">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="36c74-167">Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="36c74-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="36c74-168">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-168">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="36c74-169">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="36c74-169">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="36c74-170">Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="36c74-170">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="36c74-171">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-171">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="36c74-172">Elemento manifesto SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="36c74-172">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="36c74-173">Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="36c74-173">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="36c74-174">Define se o suplemento está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="36c74-174">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="36c74-175">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-175">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="36c74-176">Local aprimorado</span><span class="sxs-lookup"><span data-stu-id="36c74-176">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="36c74-177">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="36c74-177">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="36c74-178">Adicionado um novo objeto que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="36c74-178">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="36c74-179">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-179">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="36c74-180">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="36c74-180">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="36c74-181">Adicionado um novo objeto que representa um local.</span><span class="sxs-lookup"><span data-stu-id="36c74-181">Added a new object that represents a location.</span></span> <span data-ttu-id="36c74-182">Somente leitura.</span><span class="sxs-lookup"><span data-stu-id="36c74-182">Read only.</span></span>

<span data-ttu-id="36c74-183">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-183">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="36c74-184">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="36c74-184">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="36c74-185">Adicionado um novo objeto que representa a id de um local.</span><span class="sxs-lookup"><span data-stu-id="36c74-185">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="36c74-186">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-186">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="36c74-187">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="36c74-187">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="36c74-188">Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="36c74-188">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="36c74-189">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-189">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="36c74-190">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="36c74-190">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="36c74-191">Adicionada uma nova enumeração que especifica o tipo de local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="36c74-191">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="36c74-192">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-192">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="36c74-193">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="36c74-193">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="36c74-194">Adicionado `EnhancedLocationsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="36c74-194">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="36c74-195">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-195">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="36c74-196">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="36c74-196">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="36c74-197">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="36c74-197">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="36c74-198">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="36c74-198">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="36c74-199">**Disponível em**: Outlook para Windows (Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="36c74-199">**Available in**: Outlook for Windows (Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="36c74-200">Cabeçalhos de Internet</span><span class="sxs-lookup"><span data-stu-id="36c74-200">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="36c74-201">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="36c74-201">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="36c74-202">Adicionado um novo objeto que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="36c74-202">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="36c74-203">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-203">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="36c74-204">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="36c74-204">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="36c74-205">Adicionada uma nova propriedade que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="36c74-205">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="36c74-206">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-206">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="36c74-207">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="36c74-207">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="36c74-208">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="36c74-208">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="36c74-209">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="36c74-209">Added ability to get Office theme.</span></span>

<span data-ttu-id="36c74-210">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-210">**Available in**: Outlook for Windows (Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="36c74-211">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="36c74-211">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="36c74-212">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="36c74-212">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="36c74-213">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="36c74-213">**Available in**: Outlook for Windows (Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="36c74-214">SSO</span><span class="sxs-lookup"><span data-stu-id="36c74-214">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="36c74-215">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="36c74-215">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="36c74-216">Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="36c74-216">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="36c74-217">**Disponível em**: Outlook para Windows (Office 365), Outlook para Mac (Office 365), Outlook na Web (Office 365 e Outlook.com), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="36c74-217">**Available in**: Outlook for Windows (Office 365), Outlook for Mac (Office 365), Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="36c74-218">Confira também</span><span class="sxs-lookup"><span data-stu-id="36c74-218">See also</span></span>

- [<span data-ttu-id="36c74-219">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="36c74-219">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="36c74-220">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="36c74-220">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="36c74-221">Introdução</span><span class="sxs-lookup"><span data-stu-id="36c74-221">Get started</span></span>](/outlook/add-ins/quick-start)
