---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 05/19/2020
localization_priority: Normal
ms.openlocfilehash: 3183c81a9af99f480c2dbecc787695501380cea7
ms.sourcegitcommit: 8499a4247d1cb1e96e99c17cb520f4a8a41667e3
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/19/2020
ms.locfileid: "44292291"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="5cc66-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="5cc66-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="5cc66-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="5cc66-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5cc66-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5cc66-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="5cc66-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="5cc66-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="5cc66-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="5cc66-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="5cc66-108">Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="5cc66-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="5cc66-109">"Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="5cc66-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="5cc66-110">Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="5cc66-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="5cc66-111">"Solicitar acesso de visualização" é observado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="5cc66-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="5cc66-112">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="5cc66-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="5cc66-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="5cc66-113">Features in preview</span></span>

<span data-ttu-id="5cc66-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="5cc66-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="5cc66-115">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="5cc66-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="5cc66-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="5cc66-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="5cc66-117">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="5cc66-118">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="5cc66-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="5cc66-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="5cc66-120">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="5cc66-121">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="5cc66-122">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="5cc66-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="5cc66-123">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="5cc66-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="5cc66-124">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="5cc66-125">Office. Context. Mailbox. Item. sensibilidade</span><span class="sxs-lookup"><span data-stu-id="5cc66-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="5cc66-126">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="5cc66-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="5cc66-127">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="5cc66-128">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="5cc66-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="5cc66-129">Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="5cc66-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="5cc66-130">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="5cc66-131">Anexar ao enviar</span><span class="sxs-lookup"><span data-stu-id="5cc66-131">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="5cc66-132">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-132">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="5cc66-133">Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-133">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="5cc66-134">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="5cc66-135">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="5cc66-135">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="5cc66-136">Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="5cc66-136">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="5cc66-137">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="5cc66-138">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="5cc66-138">Event-based activation</span></span>

<span data-ttu-id="5cc66-139">Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="5cc66-139">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="5cc66-140">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="5cc66-140">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="5cc66-141">Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="5cc66-141">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="5cc66-142">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="5cc66-142">It configures event-based activation functionality.</span></span>

<span data-ttu-id="5cc66-143">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="5cc66-143">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="5cc66-144">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="5cc66-144">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="5cc66-145">`LaunchEvents`Elemento adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="5cc66-145">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="5cc66-146">Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="5cc66-146">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="5cc66-147">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="5cc66-147">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="5cc66-148">Elemento de manifesto de runtimes</span><span class="sxs-lookup"><span data-stu-id="5cc66-148">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="5cc66-149">Adicionado suporte do Outlook ao `Runtimes` elemento manifest.</span><span class="sxs-lookup"><span data-stu-id="5cc66-149">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="5cc66-150">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="5cc66-150">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="5cc66-151">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="5cc66-151">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="5cc66-152">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="5cc66-152">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="5cc66-153">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-153">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5cc66-154">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="5cc66-154">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="5cc66-155">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="5cc66-155">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="5cc66-156">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="5cc66-156">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="5cc66-157">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-157">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="5cc66-158">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-158">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="5cc66-159">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="5cc66-160">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-160">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5cc66-161">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-161">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="5cc66-162">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-162">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="5cc66-163">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-163">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="5cc66-164">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-164">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="5cc66-165">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-165">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="5cc66-166">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="5cc66-166">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="5cc66-167">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-167">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="5cc66-168">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-168">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="5cc66-169">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="5cc66-169">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="5cc66-170">Adição de uma nova enumeração `ComposeType` disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="5cc66-170">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="5cc66-171">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="5cc66-171">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="5cc66-172">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="5cc66-172">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="5cc66-173">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="5cc66-173">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="5cc66-174">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="5cc66-174">Added ability to get Office theme.</span></span>

<span data-ttu-id="5cc66-175">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-175">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="5cc66-176">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="5cc66-176">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="5cc66-177">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="5cc66-177">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="5cc66-178">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5cc66-178">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="5cc66-179">SSO (logon único)</span><span class="sxs-lookup"><span data-stu-id="5cc66-179">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="5cc66-180">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="5cc66-180">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="5cc66-181">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="5cc66-181">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="5cc66-182">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="5cc66-182">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="5cc66-183">Confira também</span><span class="sxs-lookup"><span data-stu-id="5cc66-183">See also</span></span>

- [<span data-ttu-id="5cc66-184">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="5cc66-184">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="5cc66-185">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="5cc66-185">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="5cc66-186">Introdução</span><span class="sxs-lookup"><span data-stu-id="5cc66-186">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="5cc66-187">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="5cc66-187">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
