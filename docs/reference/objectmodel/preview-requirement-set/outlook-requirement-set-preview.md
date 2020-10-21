---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: d91105e0cfbb97dc1a239e40b1c81adc4e76988b
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626593"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="69892-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="69892-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="69892-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="69892-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="69892-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="69892-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="69892-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="69892-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="69892-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="69892-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="69892-108">Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="69892-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="69892-109">"Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="69892-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="69892-110">Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="69892-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="69892-111">"Solicitar acesso de visualização" é observado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="69892-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="69892-112">O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1,9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span><span class="sxs-lookup"><span data-stu-id="69892-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="69892-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="69892-113">Features in preview</span></span>

<span data-ttu-id="69892-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="69892-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="69892-115">Ativação de suplementos em itens protegidos por IRM (gerenciamento de direitos de informação)</span><span class="sxs-lookup"><span data-stu-id="69892-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="69892-116">Agora, os suplementos podem ser ativados em itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="69892-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="69892-117">Para ativar esse recurso, um administrador de locatários precisa habilitar o `OBJMODEL` direito de uso, configurando a opção permitir política personalizada de **acesso programático** no Office.</span><span class="sxs-lookup"><span data-stu-id="69892-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="69892-118">Confira os [direitos de uso e as descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="69892-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="69892-119">**Disponível em**: Outlook no Windows, começando com a compilação 13229,10000 (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="69892-120">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="69892-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="69892-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="69892-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="69892-122">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="69892-123">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="69892-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="69892-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="69892-125">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="69892-126">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="69892-127">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="69892-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="69892-128">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="69892-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="69892-129">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="69892-130">Office. Context. Mailbox. Item. sensibilidade</span><span class="sxs-lookup"><span data-stu-id="69892-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="69892-131">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="69892-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="69892-132">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="69892-133">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="69892-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="69892-134">Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="69892-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="69892-135">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="69892-136">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="69892-136">Event-based activation</span></span>

<span data-ttu-id="69892-137">Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="69892-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="69892-138">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="69892-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="69892-139">Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="69892-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="69892-140">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="69892-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="69892-141">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-141">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="69892-142">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="69892-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="69892-143">`LaunchEvents`Elemento adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="69892-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="69892-144">Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="69892-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="69892-145">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-145">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="69892-146">Elemento de manifesto de runtimes</span><span class="sxs-lookup"><span data-stu-id="69892-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="69892-147">Adicionado suporte do Outlook ao `Runtimes` elemento manifest.</span><span class="sxs-lookup"><span data-stu-id="69892-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="69892-148">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="69892-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="69892-149">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-149">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="69892-150">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="69892-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="69892-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="69892-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="69892-152">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="69892-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="69892-153">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="69892-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="69892-154">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="69892-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="69892-155">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="69892-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="69892-156">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="69892-157">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="69892-158">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="69892-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="69892-159">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="69892-160">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="69892-161">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="69892-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="69892-162">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="69892-163">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="69892-164">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="69892-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="69892-165">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="69892-166">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="69892-167">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="69892-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="69892-168">Adição de uma nova enumeração `ComposeType` disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="69892-169">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="69892-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="69892-170">Mensagens de notificação com ações</span><span class="sxs-lookup"><span data-stu-id="69892-170">Notification messages with actions</span></span>

<span data-ttu-id="69892-171">Este recurso permite que o suplemento inclua uma mensagem de notificação com uma ação personalizada além da ação padrão de **ignorar** .</span><span class="sxs-lookup"><span data-stu-id="69892-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="69892-172">Office. NotificationMessageDetails. Actions</span><span class="sxs-lookup"><span data-stu-id="69892-172">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="69892-173">Adicionada uma nova propriedade que permite que você adicione uma `InsightMessage` notificação com uma ação personalizada.</span><span class="sxs-lookup"><span data-stu-id="69892-173">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="69892-174">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="69892-174">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="69892-175">Office. NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="69892-175">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="69892-176">Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.</span><span class="sxs-lookup"><span data-stu-id="69892-176">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="69892-177">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="69892-177">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="69892-178">Office. MailboxEnums. ActionType</span><span class="sxs-lookup"><span data-stu-id="69892-178">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="69892-179">Foi adicionada uma nova enumeração `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="69892-179">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="69892-180">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="69892-180">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="69892-181">Office. MailboxEnums. ItemNotificationMessageType. InsightMessage</span><span class="sxs-lookup"><span data-stu-id="69892-181">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="69892-182">Adicionado um novo tipo `InsightMessage` à `ItemNotificationMessageType` enumeração.</span><span class="sxs-lookup"><span data-stu-id="69892-182">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="69892-183">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="69892-183">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="69892-184">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="69892-184">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="69892-185">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="69892-185">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="69892-186">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="69892-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="69892-187">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-187">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="69892-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="69892-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="69892-189">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="69892-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="69892-190">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="69892-191">Os dados da sessão</span><span class="sxs-lookup"><span data-stu-id="69892-191">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="69892-192">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="69892-192">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="69892-193">Adicionado um novo objeto que representa os dados de sessão de um item.</span><span class="sxs-lookup"><span data-stu-id="69892-193">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="69892-194">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="69892-195">Office. Context. Mailbox. Item. sessionData</span><span class="sxs-lookup"><span data-stu-id="69892-195">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="69892-196">Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="69892-196">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="69892-197">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="69892-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="69892-198">Confira também</span><span class="sxs-lookup"><span data-stu-id="69892-198">See also</span></span>

- [<span data-ttu-id="69892-199">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="69892-199">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="69892-200">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="69892-200">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="69892-201">Introdução</span><span class="sxs-lookup"><span data-stu-id="69892-201">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="69892-202">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="69892-202">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
