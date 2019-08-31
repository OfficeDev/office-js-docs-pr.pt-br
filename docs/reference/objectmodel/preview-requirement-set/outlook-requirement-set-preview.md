---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 08/15/2019
localization_priority: Priority
ms.openlocfilehash: aa3f46c505e8c87508699f6e84194272ee4d13bb
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696453"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="23b92-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="23b92-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="23b92-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="23b92-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="23b92-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="23b92-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="23b92-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="23b92-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="23b92-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="23b92-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="23b92-107">O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="23b92-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="23b92-108">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="23b92-108">Features in preview</span></span>

<span data-ttu-id="23b92-109">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="23b92-109">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="23b92-110">Anexos</span><span class="sxs-lookup"><span data-stu-id="23b92-110">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="23b92-111">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="23b92-111">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="23b92-112">Adicionado um novo objeto que representa o conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="23b92-112">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="23b92-113">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="23b92-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="23b92-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="23b92-115">Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="23b92-115">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="23b92-116">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-116">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="23b92-117">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-117">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="23b92-118">Adicionar um novo método para acessar o conteúdo de um anexo específico.</span><span class="sxs-lookup"><span data-stu-id="23b92-118">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="23b92-119">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-119">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="23b92-120">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-120">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="23b92-121">Adicionado um novo método que obtém um item anexo no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="23b92-121">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="23b92-122">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-122">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="23b92-123">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="23b92-123">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="23b92-124">Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="23b92-124">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="23b92-125">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="23b92-126">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="23b92-126">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="23b92-127">Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.</span><span class="sxs-lookup"><span data-stu-id="23b92-127">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="23b92-128">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="23b92-129">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="23b92-129">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="23b92-130">Adicionado evento `AttachmentsChanged` ao `Item`.</span><span class="sxs-lookup"><span data-stu-id="23b92-130">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="23b92-131">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="block-on-send"></a><span data-ttu-id="23b92-132">Bloquear ao enviar</span><span class="sxs-lookup"><span data-stu-id="23b92-132">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="23b92-133">Event.completed</span><span class="sxs-lookup"><span data-stu-id="23b92-133">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="23b92-134">Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="23b92-134">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="23b92-135">Esse valor é usado para cancelar a execução de um evento.</span><span class="sxs-lookup"><span data-stu-id="23b92-135">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="23b92-136">**Disponível em:** Outlook na Web (clássico), Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-136">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="categories"></a><span data-ttu-id="23b92-137">Categorias</span><span class="sxs-lookup"><span data-stu-id="23b92-137">Categories</span></span>

<span data-ttu-id="23b92-138">No Outlook, um usuário pode agrupar mensagens e compromissos usando uma categoria para codificá-los por cor.</span><span class="sxs-lookup"><span data-stu-id="23b92-138">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="23b92-139">O usuário define as categorias em uma lista mestra em sua caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="23b92-139">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="23b92-140">Ele pode, em seguida, aplicar uma ou mais categorias a um item.</span><span class="sxs-lookup"><span data-stu-id="23b92-140">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="23b92-141">Não há suporte para esse método no Outlook no iOS ou no Android.</span><span class="sxs-lookup"><span data-stu-id="23b92-141">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="23b92-142">Categories</span><span class="sxs-lookup"><span data-stu-id="23b92-142">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="23b92-143">Adicionou um novo objeto que representa a categoria de um item.</span><span class="sxs-lookup"><span data-stu-id="23b92-143">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="23b92-144">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="23b92-145">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="23b92-145">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="23b92-146">Adicionou um novo objeto que representa os detalhes de uma categoria (seu nome e cor associada).</span><span class="sxs-lookup"><span data-stu-id="23b92-146">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="23b92-147">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-147">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="23b92-148">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="23b92-148">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="23b92-149">Adicionou um novo objeto que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="23b92-149">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="23b92-150">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="23b92-151">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="23b92-151">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="23b92-152">Adicionou uma nova propriedade que representa a lista mestra de categorias em uma caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="23b92-152">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="23b92-153">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="23b92-154">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="23b92-154">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="23b92-155">Adicionou uma nova propriedade que representa o conjunto de categorias em um item.</span><span class="sxs-lookup"><span data-stu-id="23b92-155">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="23b92-156">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="23b92-157">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="23b92-157">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="23b92-158">Adicionou uma nova enumeração que especifica as cores disponíveis a serem associadas a categorias. </span><span class="sxs-lookup"><span data-stu-id="23b92-158">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="23b92-159">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="delegate-access"></a><span data-ttu-id="23b92-160">Acesso de representante</span><span class="sxs-lookup"><span data-stu-id="23b92-160">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="23b92-161">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="23b92-161">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="23b92-162">Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="23b92-162">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="23b92-163">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="23b92-164">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-164">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="23b92-165">Adicionado um novo método que obtém o ID de um compromisso ou item de mensagem salvo.</span><span class="sxs-lookup"><span data-stu-id="23b92-165">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="23b92-166">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-166">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="23b92-167">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-167">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="23b92-168">Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="23b92-168">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="23b92-169">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="23b92-170">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="23b92-170">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="23b92-171">Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="23b92-171">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="23b92-172">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-172">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="23b92-173">Elemento manifesto SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="23b92-173">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="23b92-174">Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="23b92-174">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="23b92-175">Define se o suplemento está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="23b92-175">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="23b92-176">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="enhanced-location"></a><span data-ttu-id="23b92-177">Local aprimorado</span><span class="sxs-lookup"><span data-stu-id="23b92-177">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="23b92-178">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="23b92-178">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="23b92-179">Adicionado um novo objeto que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="23b92-179">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="23b92-180">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="23b92-181">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="23b92-181">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="23b92-182">Adicionado um novo objeto que representa um local.</span><span class="sxs-lookup"><span data-stu-id="23b92-182">Added a new object that represents a location.</span></span> <span data-ttu-id="23b92-183">Somente leitura.</span><span class="sxs-lookup"><span data-stu-id="23b92-183">Read only.</span></span>

<span data-ttu-id="23b92-184">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="23b92-185">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="23b92-185">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="23b92-186">Adicionado um novo objeto que representa a id de um local.</span><span class="sxs-lookup"><span data-stu-id="23b92-186">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="23b92-187">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="23b92-188">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="23b92-188">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="23b92-189">Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="23b92-189">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="23b92-190">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-190">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="23b92-191">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="23b92-191">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="23b92-192">Adicionada uma nova enumeração que especifica o tipo de local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="23b92-192">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="23b92-193">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-193">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="23b92-194">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="23b92-194">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="23b92-195">Adicionado evento `EnhancedLocationsChanged` ao `Item`.</span><span class="sxs-lookup"><span data-stu-id="23b92-195">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="23b92-196">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-196">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="23b92-197">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="23b92-197">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="23b92-198">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-198">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="23b92-199">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="23b92-199">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="23b92-200">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="23b92-200">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

<br>

---

### <a name="internet-headers"></a><span data-ttu-id="23b92-201">Cabeçalhos de Internet</span><span class="sxs-lookup"><span data-stu-id="23b92-201">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="23b92-202">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="23b92-202">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="23b92-203">Adicionado um novo objeto que representa os cabeçalhos de internet personalizados de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="23b92-203">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="23b92-204">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-204">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="23b92-205">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="23b92-205">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="23b92-206">Adicionada uma nova propriedade que representa os cabeçalhos de internet personalizados de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="23b92-206">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="23b92-207">**Disponível no**: Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-207">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="office-theme"></a><span data-ttu-id="23b92-208">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="23b92-208">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="23b92-209">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="23b92-209">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="23b92-210">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="23b92-210">Added ability to get Office theme.</span></span>

<span data-ttu-id="23b92-211">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-211">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="23b92-212">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="23b92-212">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="23b92-213">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="23b92-213">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="23b92-214">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="23b92-214">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="23b92-215">SSO</span><span class="sxs-lookup"><span data-stu-id="23b92-215">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="23b92-216">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="23b92-216">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="23b92-217">Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="23b92-217">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="23b92-218">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="23b92-218">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="23b92-219">Confira também</span><span class="sxs-lookup"><span data-stu-id="23b92-219">See also</span></span>

- [<span data-ttu-id="23b92-220">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="23b92-220">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="23b92-221">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="23b92-221">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="23b92-222">Introdução</span><span class="sxs-lookup"><span data-stu-id="23b92-222">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="23b92-223">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="23b92-223">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
