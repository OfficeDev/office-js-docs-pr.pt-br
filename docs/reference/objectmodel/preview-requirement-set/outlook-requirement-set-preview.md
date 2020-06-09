---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: a8026448f32d29de36684eb6a6d9fa0826de5f5b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608074"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="38138-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="38138-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="38138-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="38138-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="38138-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="38138-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="38138-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="38138-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="38138-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="38138-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="38138-108">Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="38138-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="38138-109">"Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="38138-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="38138-110">Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="38138-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="38138-111">"Solicitar acesso de visualização" é observado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="38138-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="38138-112">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="38138-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="38138-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="38138-113">Features in preview</span></span>

<span data-ttu-id="38138-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="38138-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="38138-115">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="38138-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="38138-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="38138-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="38138-117">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="38138-118">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="38138-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="38138-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="38138-120">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="38138-121">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="38138-122">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="38138-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="38138-123">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="38138-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="38138-124">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="38138-125">Office. Context. Mailbox. Item. sensibilidade</span><span class="sxs-lookup"><span data-stu-id="38138-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="38138-126">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="38138-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="38138-127">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="38138-128">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="38138-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="38138-129">Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="38138-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="38138-130">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="38138-131">Anexar ao enviar</span><span class="sxs-lookup"><span data-stu-id="38138-131">Append on send</span></span>

<span data-ttu-id="38138-132">Para saber mais sobre como usar o recurso Append-on-Send, confira [implementar anexar ao enviar em seu suplemento do Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="38138-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="38138-133">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="38138-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="38138-134">Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="38138-135">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="38138-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="38138-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="38138-137">Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="38138-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="38138-138">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-138">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="38138-139">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="38138-139">Event-based activation</span></span>

<span data-ttu-id="38138-140">Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="38138-140">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="38138-141">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="38138-141">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="38138-142">Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="38138-142">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="38138-143">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="38138-143">It configures event-based activation functionality.</span></span>

<span data-ttu-id="38138-144">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="38138-144">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="38138-145">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="38138-145">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="38138-146">`LaunchEvents`Elemento adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="38138-146">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="38138-147">Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="38138-147">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="38138-148">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="38138-148">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="38138-149">Elemento de manifesto de runtimes</span><span class="sxs-lookup"><span data-stu-id="38138-149">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="38138-150">Adicionado suporte do Outlook ao `Runtimes` elemento manifest.</span><span class="sxs-lookup"><span data-stu-id="38138-150">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="38138-151">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="38138-151">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="38138-152">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="38138-152">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="38138-153">Obter todas as propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="38138-153">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="38138-154">CustomProperties. getAll</span><span class="sxs-lookup"><span data-stu-id="38138-154">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="38138-155">Foi adicionada uma nova função ao `CustomProperties` objeto que obtém todas as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="38138-155">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="38138-156">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook no Mac (conectado à assinatura do Office 365), Outlook no Android, Outlook no Ios</span><span class="sxs-lookup"><span data-stu-id="38138-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="38138-157">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="38138-157">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="38138-158">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="38138-158">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="38138-159">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="38138-159">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="38138-160">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="38138-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="38138-161">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="38138-161">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="38138-162">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="38138-162">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="38138-163">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-163">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="38138-164">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-164">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="38138-165">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="38138-165">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="38138-166">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-166">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="38138-167">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="38138-168">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="38138-168">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="38138-169">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-169">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="38138-170">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="38138-171">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="38138-171">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="38138-172">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-172">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="38138-173">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="38138-174">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="38138-174">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="38138-175">Adição de uma nova enumeração `ComposeType` disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="38138-175">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="38138-176">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="38138-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="38138-177">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="38138-177">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="38138-178">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="38138-178">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="38138-179">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="38138-179">Added ability to get Office theme.</span></span>

<span data-ttu-id="38138-180">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-180">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="38138-181">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="38138-181">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="38138-182">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="38138-182">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="38138-183">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="38138-183">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="38138-184">SSO (logon único)</span><span class="sxs-lookup"><span data-stu-id="38138-184">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="38138-185">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="38138-185">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="38138-186">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="38138-186">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="38138-187">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="38138-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="38138-188">Confira também</span><span class="sxs-lookup"><span data-stu-id="38138-188">See also</span></span>

- [<span data-ttu-id="38138-189">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="38138-189">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="38138-190">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="38138-190">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="38138-191">Introdução</span><span class="sxs-lookup"><span data-stu-id="38138-191">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="38138-192">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="38138-192">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
