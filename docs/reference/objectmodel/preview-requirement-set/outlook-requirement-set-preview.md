---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em visualização para os complementos do Outlook.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: 39dd1221f4dea9674c89cdaad20024ce408f8db3
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104837"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="1c66d-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="1c66d-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="1c66d-104">O subconjunto da API de complemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um complemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="1c66d-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1c66d-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="1c66d-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="1c66d-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="1c66d-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="1c66d-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="1c66d-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="1c66d-108">Você pode visualizar recursos no Outlook na Web configurando o lançamento direcionado [no locatário do Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="1c66d-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="1c66d-109">"Configurar o acesso de visualização" é notado nesta página para recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="1c66d-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="1c66d-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="1c66d-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="1c66d-111">"Solicitar acesso de visualização" é notado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="1c66d-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="1c66d-112">O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span><span class="sxs-lookup"><span data-stu-id="1c66d-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="1c66d-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="1c66d-113">Features in preview</span></span>

<span data-ttu-id="1c66d-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="1c66d-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="1c66d-115">Ativação de um complemento em itens protegidos pelo Gerenciamento de Direitos de Informação (IRM)</span><span class="sxs-lookup"><span data-stu-id="1c66d-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="1c66d-116">Os complementos agora podem ser ativados em itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="1c66d-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="1c66d-117">Para ativar esse recurso, um administrador de locatários precisa habilitar o direito de uso definindo a opção Permitir acesso `OBJMODEL` **programático** personalizado de política no Office.</span><span class="sxs-lookup"><span data-stu-id="1c66d-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="1c66d-118">Consulte [Direitos de uso e descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="1c66d-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="1c66d-119">**Disponível em:** Outlook no Windows, a partir do build 13229.10000 (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="1c66d-120">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="1c66d-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="1c66d-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="1c66d-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="1c66d-122">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="1c66d-123">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="1c66d-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="1c66d-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="1c66d-125">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="1c66d-126">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="1c66d-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="1c66d-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="1c66d-128">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="1c66d-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="1c66d-129">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="1c66d-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="1c66d-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="1c66d-131">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1c66d-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="1c66d-132">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="1c66d-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="1c66d-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="1c66d-134">Adicionada uma nova enum `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="1c66d-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="1c66d-135">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="1c66d-136">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="1c66d-136">Event-based activation</span></span>

<span data-ttu-id="1c66d-137">Adicionado suporte para a funcionalidade de ativação baseada em eventos em complementos do Outlook. Confira [Configurar seu complemento do Outlook para ativação baseada em eventos](../../../outlook/autolaunch.md) para saber mais.</span><span class="sxs-lookup"><span data-stu-id="1c66d-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="1c66d-138">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="1c66d-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="1c66d-139">Adicionado `LaunchEvent` suporte ao ponto de extensão para manifesto.</span><span class="sxs-lookup"><span data-stu-id="1c66d-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="1c66d-140">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="1c66d-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="1c66d-141">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="1c66d-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="1c66d-142">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="1c66d-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="1c66d-143">Elemento `LaunchEvents` adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="1c66d-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="1c66d-144">Ele dá suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="1c66d-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="1c66d-145">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="1c66d-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="1c66d-146">Elemento de manifesto runtimes</span><span class="sxs-lookup"><span data-stu-id="1c66d-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="1c66d-147">Adicionado suporte do Outlook ao elemento `Runtimes` de manifesto.</span><span class="sxs-lookup"><span data-stu-id="1c66d-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="1c66d-148">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="1c66d-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="1c66d-149">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="1c66d-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="1c66d-150">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="1c66d-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="1c66d-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="1c66d-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1c66d-152">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="1c66d-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="1c66d-153">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="1c66d-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="1c66d-154">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="1c66d-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="1c66d-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="1c66d-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="1c66d-156">Adicionada uma nova função ao objeto que adiciona ou substitui a `Body` assinatura no corpo do item no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="1c66d-157">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="1c66d-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="1c66d-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="1c66d-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1c66d-159">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="1c66d-160">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="1c66d-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="1c66d-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="1c66d-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="1c66d-162">Adicionada uma nova função que obtém o tipo de composição de uma mensagem no modo redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="1c66d-163">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="1c66d-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="1c66d-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="1c66d-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="1c66d-165">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no item no modo redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="1c66d-166">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="1c66d-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="1c66d-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="1c66d-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="1c66d-168">Adicionada uma nova enum `ComposeType` disponível no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="1c66d-169">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="1c66d-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="1c66d-170">Mensagens de notificação com ações</span><span class="sxs-lookup"><span data-stu-id="1c66d-170">Notification messages with actions</span></span>

<span data-ttu-id="1c66d-171">Esse recurso permite que o seu complemento inclua uma mensagem de notificação com uma ação personalizada além da ação **Padrão Descartar.**</span><span class="sxs-lookup"><span data-stu-id="1c66d-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="1c66d-172">No Outlook na Web moderno, esse recurso está disponível somente no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="1c66d-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="1c66d-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="1c66d-174">Adicionada uma nova propriedade que permite adicionar uma `InsightMessage` notificação com uma ação personalizada.</span><span class="sxs-lookup"><span data-stu-id="1c66d-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="1c66d-175">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="1c66d-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="1c66d-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="1c66d-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="1c66d-177">Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="1c66d-178">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="1c66d-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="1c66d-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="1c66d-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="1c66d-180">Adicionada uma nova enum `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="1c66d-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="1c66d-181">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="1c66d-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="1c66d-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="1c66d-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="1c66d-183">Adicionado um novo `InsightMessage` tipo à `ItemNotificationMessageType` enum.</span><span class="sxs-lookup"><span data-stu-id="1c66d-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="1c66d-184">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="1c66d-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="1c66d-185">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="1c66d-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="1c66d-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="1c66d-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="1c66d-187">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="1c66d-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="1c66d-188">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="1c66d-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="1c66d-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="1c66d-190">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="1c66d-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="1c66d-191">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="1c66d-192">Os dados da sessão</span><span class="sxs-lookup"><span data-stu-id="1c66d-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="1c66d-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="1c66d-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="1c66d-194">Adicionado um novo objeto que representa os dados de sessão de um item.</span><span class="sxs-lookup"><span data-stu-id="1c66d-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="1c66d-195">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="1c66d-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="1c66d-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="1c66d-197">Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo redação.</span><span class="sxs-lookup"><span data-stu-id="1c66d-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="1c66d-198">**Disponível em:** Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="1c66d-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="1c66d-199">Confira também</span><span class="sxs-lookup"><span data-stu-id="1c66d-199">See also</span></span>

- [<span data-ttu-id="1c66d-200">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="1c66d-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="1c66d-201">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="1c66d-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1c66d-202">Introdução</span><span class="sxs-lookup"><span data-stu-id="1c66d-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="1c66d-203">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="1c66d-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
