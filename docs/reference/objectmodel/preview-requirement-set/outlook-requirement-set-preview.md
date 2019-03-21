---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: d24c4647116b4af56d85a434f3ece5ccf4662a39
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691164"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="e4a77-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="e4a77-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="e4a77-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="e4a77-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e4a77-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e4a77-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="e4a77-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="e4a77-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="e4a77-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="e4a77-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="e4a77-107">Os métodos e as propriedades que são apresentadas neste conjunto de requisitos devem ser testados individualmente para disponibilidade antes de usá-los.</span><span class="sxs-lookup"><span data-stu-id="e4a77-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="e4a77-108">Você também precisará ingressar no [programa Office Insider](https://products.office.com/office-insider).</span><span class="sxs-lookup"><span data-stu-id="e4a77-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="e4a77-109">O modo de visualização do conjunto de requisitos inclui todos os recursos do [Conjunto de requisitos 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span><span class="sxs-lookup"><span data-stu-id="e4a77-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="e4a77-110">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="e4a77-110">Features in preview</span></span>

<span data-ttu-id="e4a77-111">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="e4a77-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="e4a77-112">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="e4a77-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="e4a77-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="e4a77-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="e4a77-114">Adicionado um novo parâmetro opcional `options`, que é um dicionário com um valor válido `allowEvent`.</span><span class="sxs-lookup"><span data-stu-id="e4a77-114">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="e4a77-115">Esse valor é usado para cancelar a execução de um evento.</span><span class="sxs-lookup"><span data-stu-id="e4a77-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="e4a77-116">**Disponível em**: Outlook na web (clássico)</span><span class="sxs-lookup"><span data-stu-id="e4a77-116">**Available in**: Outlook on the web (Classic)</span></span>

### <a name="attachments"></a><span data-ttu-id="e4a77-117">Attachments</span><span class="sxs-lookup"><span data-stu-id="e4a77-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="e4a77-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="e4a77-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="e4a77-119">Adicionado um novo objeto que representa o conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="e4a77-119">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="e4a77-120">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-120">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="e4a77-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e4a77-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="e4a77-122">Adicionado um novo método que permite anexar um arquivo representado como uma cadeia de caracteres codificada na Base64 para uma mensagem ou um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e4a77-122">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="e4a77-123">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-123">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="e4a77-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="e4a77-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="e4a77-125">Adicionar um novo método para acessar o conteúdo de um anexo específico.</span><span class="sxs-lookup"><span data-stu-id="e4a77-125">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="e4a77-126">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-126">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="e4a77-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="e4a77-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="e4a77-128">Adicionado um novo método que obtém um item anexo no modo de redação.</span><span class="sxs-lookup"><span data-stu-id="e4a77-128">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="e4a77-129">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-129">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="e4a77-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="e4a77-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="e4a77-131">Adicionada uma nova enumeração que especifica a formatação que se aplica ao conteúdo de um anexo.</span><span class="sxs-lookup"><span data-stu-id="e4a77-131">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="e4a77-132">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-132">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="e4a77-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="e4a77-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="e4a77-134">Adicionada uma nova enumeração que especifica se um anexo foi adicionado ou removido de um item.</span><span class="sxs-lookup"><span data-stu-id="e4a77-134">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="e4a77-135">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-135">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e4a77-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="e4a77-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e4a77-137">Adicionado `AttachmentsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="e4a77-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="e4a77-138">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-138">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="delegate-access"></a><span data-ttu-id="e4a77-139">Acesso de representante</span><span class="sxs-lookup"><span data-stu-id="e4a77-139">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="e4a77-140">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="e4a77-140">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="e4a77-141">Adicionado um novo objeto que representa as propriedades de um item de compromisso ou de mensagem em uma pasta compartilhada, calendário ou caixa de correio.</span><span class="sxs-lookup"><span data-stu-id="e4a77-141">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="e4a77-142">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-142">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="e4a77-143">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e4a77-143">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="e4a77-144">Adicionado um novo método que é um objeto que representa sharedProperties de um compromisso ou item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="e4a77-144">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="e4a77-145">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-145">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="e4a77-146">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="e4a77-146">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="e4a77-147">Adicionada uma novo enumeração de sinalizador bits que especifica as permissões de representante.</span><span class="sxs-lookup"><span data-stu-id="e4a77-147">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="e4a77-148">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-148">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="e4a77-149">Elemento manifesto SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e4a77-149">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="e4a77-150">Adicionado um elemento filho ao elemento do manifesto [DesktopFormFactor](../../manifest/desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e4a77-150">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="e4a77-151">Define se o suplemento está disponível nos cenários de representante.</span><span class="sxs-lookup"><span data-stu-id="e4a77-151">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="e4a77-152">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-152">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="enhanced-location"></a><span data-ttu-id="e4a77-153">Local aprimorado</span><span class="sxs-lookup"><span data-stu-id="e4a77-153">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="e4a77-154">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e4a77-154">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="e4a77-155">Adicionado um novo objeto que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e4a77-155">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e4a77-156">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-156">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="e4a77-157">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="e4a77-157">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="e4a77-158">Adicionado um novo objeto que representa um local.</span><span class="sxs-lookup"><span data-stu-id="e4a77-158">Added a new object that represents a location.</span></span> <span data-ttu-id="e4a77-159">Somente leitura.</span><span class="sxs-lookup"><span data-stu-id="e4a77-159">Read only.</span></span>

<span data-ttu-id="e4a77-160">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-160">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="e4a77-161">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="e4a77-161">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="e4a77-162">Adicionado um novo objeto que representa a id de um local.</span><span class="sxs-lookup"><span data-stu-id="e4a77-162">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="e4a77-163">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-163">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="e4a77-164">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="e4a77-164">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="e4a77-165">Adicionada uma nova propriedade que representa o conjunto de locais em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="e4a77-165">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="e4a77-166">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-166">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="e4a77-167">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="e4a77-167">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="e4a77-168">Adicionada uma nova enumeração que especifica o tipo de local do compromisso.</span><span class="sxs-lookup"><span data-stu-id="e4a77-168">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="e4a77-169">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-169">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e4a77-170">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="e4a77-170">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e4a77-171">Adicionado `EnhancedLocationsChanged` evento `Item`.</span><span class="sxs-lookup"><span data-stu-id="e4a77-171">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="e4a77-172">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-172">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="e4a77-173">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="e4a77-173">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="e4a77-174">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e4a77-174">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="e4a77-175">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="e4a77-175">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="e4a77-176">**Disponível em**: Outlook para Windows (Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="e4a77-176">**Available in**: Office 2019 for Windows (Office 365 subscription), Outlook on the web (Classic)</span></span>

### <a name="internet-headers"></a><span data-ttu-id="e4a77-177">Cabeçalhos de Internet</span><span class="sxs-lookup"><span data-stu-id="e4a77-177">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="e4a77-178">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e4a77-178">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="e4a77-179">Adicionado um novo objeto que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="e4a77-179">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="e4a77-180">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-180">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="e4a77-181">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="e4a77-181">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="e4a77-182">Adicionada uma nova propriedade que representa os cabeçalhos de Internet de um item de mensagem.</span><span class="sxs-lookup"><span data-stu-id="e4a77-182">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="e4a77-183">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-183">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="office-theme"></a><span data-ttu-id="e4a77-184">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="e4a77-184">Office theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="e4a77-185">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="e4a77-185">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="e4a77-186">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="e4a77-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="e4a77-187">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-187">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="e4a77-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="e4a77-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="e4a77-189">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="e4a77-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="e4a77-190">**Disponível em**: Outlook para Windows (Office 365)</span><span class="sxs-lookup"><span data-stu-id="e4a77-190">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="sso"></a><span data-ttu-id="e4a77-191">SSO</span><span class="sxs-lookup"><span data-stu-id="e4a77-191">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="e4a77-192">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e4a77-192">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="e4a77-193">Foi adicionado acesso ao `getAccessTokenAsync`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="e4a77-193">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="e4a77-194">**Disponível em**: Outlook para Windows (Office 365), Outlook para Mac (Office 365), Outlook na Web (Office 365 e Outlook.com), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="e4a77-194">**Available in**: Outlook 2019 for Windows (Office 365 subscription), Outlook 2019 for Mac, Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="e4a77-195">Confira também</span><span class="sxs-lookup"><span data-stu-id="e4a77-195">See also</span></span>

- [<span data-ttu-id="e4a77-196">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="e4a77-196">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="e4a77-197">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="e4a77-197">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="e4a77-198">Introdução</span><span class="sxs-lookup"><span data-stu-id="e4a77-198">Get started</span></span>](/outlook/add-ins/quick-start)
