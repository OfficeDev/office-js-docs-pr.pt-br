---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b46fada2fa69f3526c929a0289341f7dab5b58b8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128472"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="90d5e-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="90d5e-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="90d5e-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="90d5e-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="90d5e-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="90d5e-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="90d5e-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="90d5e-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="90d5e-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="90d5e-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="90d5e-107">Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.</span><span class="sxs-lookup"><span data-stu-id="90d5e-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="90d5e-108">Você também precisará ingressar no [programa Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="90d5e-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="90d5e-109">O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="90d5e-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="90d5e-110">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="90d5e-110">Features in preview</span></span>

<span data-ttu-id="90d5e-111">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="90d5e-111">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="90d5e-112">Anexos</span><span class="sxs-lookup"><span data-stu-id="90d5e-112">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="90d5e-113">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="90d5e-113">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="90d5e-114">Adicionado um novo objeto que representa o conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="90d5e-114">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="90d5e-115">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-115">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="90d5e-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="90d5e-116">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="90d5e-117">Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="90d5e-117">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="90d5e-118">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-118">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="90d5e-119">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-119">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="90d5e-120">Adicionar um novo método para acessar o conteúdo de um anexo específico.</span><span class="sxs-lookup"><span data-stu-id="90d5e-120">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="90d5e-121">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-121">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="90d5e-122">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-122">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="90d5e-123">Adicionado um novo método que obtém um item anexo no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="90d5e-123">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="90d5e-124">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-124">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="90d5e-125">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="90d5e-125">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="90d5e-126">Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="90d5e-126">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="90d5e-127">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-127">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="90d5e-128">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="90d5e-128">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="90d5e-129">Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.</span><span class="sxs-lookup"><span data-stu-id="90d5e-129">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="90d5e-130">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-130">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="90d5e-131">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="90d5e-131">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="90d5e-132">Adicionado `AttachmentsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="90d5e-132">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="90d5e-133">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-133">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="90d5e-134">Bloquear ao enviar</span><span class="sxs-lookup"><span data-stu-id="90d5e-134">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="90d5e-135">Event.completed</span><span class="sxs-lookup"><span data-stu-id="90d5e-135">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="90d5e-136">Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="90d5e-136">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="90d5e-137">Esse valor é usado para cancelar a execução de um evento.</span><span class="sxs-lookup"><span data-stu-id="90d5e-137">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="90d5e-138">**Disponível em**: Outlook na Web (classic)</span><span class="sxs-lookup"><span data-stu-id="90d5e-138">**Available in**: Outlook on the web (Classic)</span></span>

---

### <a name="categories"></a><span data-ttu-id="90d5e-139">Categorias</span><span class="sxs-lookup"><span data-stu-id="90d5e-139">Categories</span></span>

<span data-ttu-id="90d5e-140">No Outlook, um usuário pode agrupar mensagens e compromissos usando uma categoria para codificá-los por cor.</span><span class="sxs-lookup"><span data-stu-id="90d5e-140">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="90d5e-141">O usuário define as categorias em uma lista mestra em sua caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="90d5e-141">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="90d5e-142">Ele pode, em seguida, aplicar uma ou mais categorias a um item.</span><span class="sxs-lookup"><span data-stu-id="90d5e-142">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="90d5e-143">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="90d5e-143">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="90d5e-144">Categories</span><span class="sxs-lookup"><span data-stu-id="90d5e-144">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="90d5e-145">Adicionou um novo objeto que representa a categoria de um item.</span><span class="sxs-lookup"><span data-stu-id="90d5e-145">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="90d5e-146">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-146">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="90d5e-147">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="90d5e-147">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="90d5e-148">Adicionou um novo objeto que representa os detalhes de uma categoria (seu nome e cor associada).</span><span class="sxs-lookup"><span data-stu-id="90d5e-148">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="90d5e-149">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-149">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="90d5e-150">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="90d5e-150">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="90d5e-151">Adicionou um novo objeto que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="90d5e-151">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="90d5e-152">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-152">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="90d5e-153">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="90d5e-153">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="90d5e-154">Adicionou uma nova propriedade que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="90d5e-154">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="90d5e-155">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-155">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="90d5e-156">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="90d5e-156">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="90d5e-157">Adicionou uma nova propriedade que representa o conjunto de categorias em um item.</span><span class="sxs-lookup"><span data-stu-id="90d5e-157">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="90d5e-158">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-158">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="90d5e-159">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="90d5e-159">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="90d5e-160">Adicionou uma nova enumeração que especifica as cores disponíveis a serem associadas a categorias. </span><span class="sxs-lookup"><span data-stu-id="90d5e-160">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="90d5e-161">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-161">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="90d5e-162">Acesso de representante</span><span class="sxs-lookup"><span data-stu-id="90d5e-162">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="90d5e-163">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="90d5e-163">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="90d5e-164">Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="90d5e-164">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="90d5e-165">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-165">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="90d5e-166">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-166">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="90d5e-167">Adicionado um novo método que obtém o ID de um compromisso ou item de mensagem salvo.</span><span class="sxs-lookup"><span data-stu-id="90d5e-167">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="90d5e-168">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-168">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="90d5e-169">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-169">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="90d5e-170">Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="90d5e-170">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="90d5e-171">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-171">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="90d5e-172">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="90d5e-172">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="90d5e-173">Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="90d5e-173">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="90d5e-174">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-174">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="90d5e-175">Elemento manifesto SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="90d5e-175">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="90d5e-176">Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="90d5e-176">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="90d5e-177">Define se o suplemento está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="90d5e-177">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="90d5e-178">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-178">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="90d5e-179">Local aprimorado</span><span class="sxs-lookup"><span data-stu-id="90d5e-179">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="90d5e-180">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="90d5e-180">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="90d5e-181">Adicionado um novo objeto que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="90d5e-181">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="90d5e-182">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-182">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="90d5e-183">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="90d5e-183">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="90d5e-184">Adicionado um novo objeto que representa um local.</span><span class="sxs-lookup"><span data-stu-id="90d5e-184">Added a new object that represents a location.</span></span> <span data-ttu-id="90d5e-185">Somente leitura.</span><span class="sxs-lookup"><span data-stu-id="90d5e-185">Read only.</span></span>

<span data-ttu-id="90d5e-186">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-186">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="90d5e-187">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="90d5e-187">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="90d5e-188">Adicionado um novo objeto que representa a id de um local.</span><span class="sxs-lookup"><span data-stu-id="90d5e-188">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="90d5e-189">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-189">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="90d5e-190">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="90d5e-190">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="90d5e-191">Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="90d5e-191">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="90d5e-192">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-192">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="90d5e-193">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="90d5e-193">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="90d5e-194">Adicionada uma nova enumeração que especifica o tipo de local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="90d5e-194">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="90d5e-195">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-195">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="90d5e-196">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="90d5e-196">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="90d5e-197">Adicionado `EnhancedLocationsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="90d5e-197">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="90d5e-198">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-198">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="90d5e-199">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="90d5e-199">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="90d5e-200">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-200">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="90d5e-201">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="90d5e-201">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="90d5e-202">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="90d5e-202">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="90d5e-203">Cabeçalhos de Internet</span><span class="sxs-lookup"><span data-stu-id="90d5e-203">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="90d5e-204">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="90d5e-204">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="90d5e-205">Adicionado um novo objeto que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="90d5e-205">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="90d5e-206">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-206">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="90d5e-207">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="90d5e-207">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="90d5e-208">Adicionada uma nova propriedade que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="90d5e-208">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="90d5e-209">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-209">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="90d5e-210">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="90d5e-210">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="90d5e-211">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="90d5e-211">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="90d5e-212">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="90d5e-212">Added ability to get Office theme.</span></span>

<span data-ttu-id="90d5e-213">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-213">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="90d5e-214">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="90d5e-214">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="90d5e-215">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="90d5e-215">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="90d5e-216">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="90d5e-216">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="90d5e-217">SSO</span><span class="sxs-lookup"><span data-stu-id="90d5e-217">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="90d5e-218">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="90d5e-218">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="90d5e-219">Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="90d5e-219">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="90d5e-220">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (novo), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="90d5e-220">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (new), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="90d5e-221">Confira também</span><span class="sxs-lookup"><span data-stu-id="90d5e-221">See also</span></span>

- [<span data-ttu-id="90d5e-222">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="90d5e-222">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="90d5e-223">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="90d5e-223">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="90d5e-224">Introdução</span><span class="sxs-lookup"><span data-stu-id="90d5e-224">Get started</span></span>](/outlook/add-ins/quick-start)
