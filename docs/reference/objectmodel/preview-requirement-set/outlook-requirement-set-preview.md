---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 2f83f81dcf7aa7ab0e3a48fff4279c1e08ba6286
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2020
ms.locfileid: "49612747"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9a844-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a844-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9a844-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9a844-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9a844-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9a844-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9a844-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="9a844-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9a844-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="9a844-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="9a844-108">Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="9a844-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="9a844-109">"Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="9a844-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="9a844-110">Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="9a844-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="9a844-111">"Solicitar acesso de visualização" é observado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="9a844-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="9a844-112">O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1,9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span><span class="sxs-lookup"><span data-stu-id="9a844-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9a844-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="9a844-113">Features in preview</span></span>

<span data-ttu-id="9a844-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="9a844-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="9a844-115">Ativação de suplementos em itens protegidos por IRM (gerenciamento de direitos de informação)</span><span class="sxs-lookup"><span data-stu-id="9a844-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="9a844-116">Agora, os suplementos podem ser ativados em itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="9a844-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="9a844-117">Para ativar esse recurso, um administrador de locatários precisa habilitar o `OBJMODEL` direito de uso, configurando a opção permitir política personalizada de **acesso programático** no Office.</span><span class="sxs-lookup"><span data-stu-id="9a844-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="9a844-118">Confira os [direitos de uso e as descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="9a844-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="9a844-119">**Disponível em**: Outlook no Windows, começando com a compilação 13229,10000 (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="9a844-120">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="9a844-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="9a844-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9a844-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9a844-122">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="9a844-123">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="9a844-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="9a844-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9a844-125">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="9a844-126">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="9a844-127">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9a844-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9a844-128">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="9a844-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="9a844-129">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="9a844-130">Office. Context. Mailbox. Item. sensibilidade</span><span class="sxs-lookup"><span data-stu-id="9a844-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9a844-131">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a844-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="9a844-132">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="9a844-133">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="9a844-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9a844-134">Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9a844-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="9a844-135">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="9a844-136">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="9a844-136">Event-based activation</span></span>

<span data-ttu-id="9a844-137">Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="9a844-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="9a844-138">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="9a844-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="9a844-139">Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="9a844-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="9a844-140">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="9a844-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="9a844-141">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-141">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="9a844-142">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="9a844-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="9a844-143">`LaunchEvents`Elemento adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="9a844-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="9a844-144">Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="9a844-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="9a844-145">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-145">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="9a844-146">Elemento de manifesto de runtimes</span><span class="sxs-lookup"><span data-stu-id="9a844-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="9a844-147">Adicionado suporte do Outlook ao `Runtimes` elemento manifest.</span><span class="sxs-lookup"><span data-stu-id="9a844-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="9a844-148">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="9a844-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="9a844-149">**Disponível no**: Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-149">**Available in**: Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9a844-150">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="9a844-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9a844-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9a844-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9a844-152">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="9a844-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9a844-153">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9a844-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="9a844-154">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="9a844-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="9a844-155">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="9a844-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="9a844-156">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="9a844-157">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="9a844-158">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="9a844-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9a844-159">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="9a844-160">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="9a844-161">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="9a844-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="9a844-162">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="9a844-163">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="9a844-164">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="9a844-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9a844-165">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="9a844-166">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="9a844-167">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="9a844-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9a844-168">Adição de uma nova enumeração `ComposeType` disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="9a844-169">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9a844-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="9a844-170">Mensagens de notificação com ações</span><span class="sxs-lookup"><span data-stu-id="9a844-170">Notification messages with actions</span></span>

<span data-ttu-id="9a844-171">Este recurso permite que o suplemento inclua uma mensagem de notificação com uma ação personalizada além da ação padrão de **ignorar** .</span><span class="sxs-lookup"><span data-stu-id="9a844-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="9a844-172">No Outlook moderno na Web, este recurso está disponível somente no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="9a844-173">Office. NotificationMessageDetails. Actions</span><span class="sxs-lookup"><span data-stu-id="9a844-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="9a844-174">Adicionada uma nova propriedade que permite que você adicione uma `InsightMessage` notificação com uma ação personalizada.</span><span class="sxs-lookup"><span data-stu-id="9a844-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="9a844-175">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9a844-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="9a844-176">Office. NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="9a844-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="9a844-177">Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.</span><span class="sxs-lookup"><span data-stu-id="9a844-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="9a844-178">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9a844-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="9a844-179">Office. MailboxEnums. ActionType</span><span class="sxs-lookup"><span data-stu-id="9a844-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="9a844-180">Foi adicionada uma nova enumeração `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="9a844-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="9a844-181">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9a844-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="9a844-182">Office. MailboxEnums. ItemNotificationMessageType. InsightMessage</span><span class="sxs-lookup"><span data-stu-id="9a844-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="9a844-183">Adicionado um novo tipo `InsightMessage` à `ItemNotificationMessageType` enumeração.</span><span class="sxs-lookup"><span data-stu-id="9a844-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="9a844-184">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9a844-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9a844-185">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="9a844-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9a844-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9a844-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9a844-187">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="9a844-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="9a844-188">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9a844-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9a844-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9a844-190">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="9a844-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9a844-191">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="9a844-192">Os dados da sessão</span><span class="sxs-lookup"><span data-stu-id="9a844-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="9a844-193">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="9a844-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="9a844-194">Adicionado um novo objeto que representa os dados de sessão de um item.</span><span class="sxs-lookup"><span data-stu-id="9a844-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="9a844-195">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="9a844-196">Office. Context. Mailbox. Item. sessionData</span><span class="sxs-lookup"><span data-stu-id="9a844-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9a844-197">Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9a844-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="9a844-198">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9a844-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="9a844-199">Confira também</span><span class="sxs-lookup"><span data-stu-id="9a844-199">See also</span></span>

- [<span data-ttu-id="9a844-200">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a844-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9a844-201">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9a844-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9a844-202">Introdução</span><span class="sxs-lookup"><span data-stu-id="9a844-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9a844-203">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="9a844-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
