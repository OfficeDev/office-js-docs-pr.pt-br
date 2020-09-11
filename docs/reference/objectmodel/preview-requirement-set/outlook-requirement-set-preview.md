---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 24cad394f0f3ffb95a05a81ccb38ee4aa72a3797
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431063"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="bf795-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="bf795-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="bf795-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="bf795-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bf795-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="bf795-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="bf795-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="bf795-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="bf795-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="bf795-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="bf795-108">Você pode Visualizar recursos no Outlook na Web [Configurando a versão de destino no seu locatário do Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="bf795-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="bf795-109">"Configurar acesso de visualização" é indicado nesta página para ver os recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="bf795-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="bf795-110">Para outros recursos, talvez você possa solicitar acesso aos bits de visualização do Outlook na Web usando sua conta do Microsoft 365, concluindo e enviando [este formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="bf795-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="bf795-111">"Solicitar acesso de visualização" é observado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="bf795-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="bf795-112">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="bf795-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="bf795-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="bf795-113">Features in preview</span></span>

<span data-ttu-id="bf795-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="bf795-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="bf795-115">Ativação de suplementos em itens protegidos por IRM (gerenciamento de direitos de informação)</span><span class="sxs-lookup"><span data-stu-id="bf795-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="bf795-116">Agora, os suplementos podem ser ativados em itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="bf795-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="bf795-117">Para ativar esse recurso, um administrador de locatários precisa habilitar o `OBJMODEL` direito de uso, configurando a opção permitir política personalizada de **acesso programático** no Office.</span><span class="sxs-lookup"><span data-stu-id="bf795-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="bf795-118">Confira os [direitos de uso e as descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="bf795-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="bf795-119">**Disponível em**: Outlook no Windows, começando com a compilação 13229,10000 (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="bf795-120">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="bf795-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="bf795-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="bf795-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="bf795-122">Adicionado um novo objeto que representa a propriedade de evento de dia inteiro de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="bf795-123">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="bf795-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="bf795-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="bf795-125">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="bf795-126">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="bf795-127">Office. Context. Mailbox. Item. isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="bf795-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bf795-128">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="bf795-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="bf795-129">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="bf795-130">Office. Context. Mailbox. Item. sensibilidade</span><span class="sxs-lookup"><span data-stu-id="bf795-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bf795-131">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bf795-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="bf795-132">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="bf795-133">Office. MailboxEnums. AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="bf795-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="bf795-134">Foi adicionada uma nova enumeração `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="bf795-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="bf795-135">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="bf795-136">Acrescentar ao enviar</span><span class="sxs-lookup"><span data-stu-id="bf795-136">Append on send</span></span>

<span data-ttu-id="bf795-137">Para saber mais sobre como usar o recurso Append-on-Send, confira [implementar anexar ao enviar em seu suplemento do Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="bf795-137">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="bf795-138">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-138">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

<span data-ttu-id="bf795-139">Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-139">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="bf795-140">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="bf795-141">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="bf795-141">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="bf795-142">Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="bf795-142">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="bf795-143">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="bf795-144">Versões assíncronas de `display` APIs</span><span class="sxs-lookup"><span data-stu-id="bf795-144">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="bf795-145">Office. Context. Mailbox. displayAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-145">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="bf795-146">Foi adicionada uma nova função ao `Mailbox` objeto que exibe um compromisso existente.</span><span class="sxs-lookup"><span data-stu-id="bf795-146">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="bf795-147">Esta é a versão assíncrona do `displayAppointmentForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-147">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="bf795-148">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="bf795-149">Office. Context. Mailbox. displayMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-149">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="bf795-150">Foi adicionada uma nova função ao `Mailbox` objeto que exibe uma mensagem existente.</span><span class="sxs-lookup"><span data-stu-id="bf795-150">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="bf795-151">Esta é a versão assíncrona do `displayMessageForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-151">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="bf795-152">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-152">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="bf795-153">Office. Context. Mailbox. displayNewAppointmentFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-153">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="bf795-154">Foi adicionada uma nova função ao `Mailbox` objeto que exibe um novo formulário de compromisso.</span><span class="sxs-lookup"><span data-stu-id="bf795-154">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="bf795-155">Esta é a versão assíncrona do `displayNewAppointmentForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-155">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="bf795-156">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-156">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="bf795-157">Office. Context. Mailbox. displayNewMessageFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-157">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="bf795-158">Foi adicionada uma nova função ao `Mailbox` objeto que exibe um novo formulário de mensagem.</span><span class="sxs-lookup"><span data-stu-id="bf795-158">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="bf795-159">Esta é a versão assíncrona do `displayNewMessageForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-159">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="bf795-160">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="bf795-161">Office. Context. Mailbox. Item. displayReplyAllFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-161">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bf795-162">Foi adicionada uma nova função ao `Item` objeto que exibe o formulário "responder a todos" no modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bf795-162">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="bf795-163">Esta é a versão assíncrona do `displayReplyAllForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-163">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="bf795-164">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-164">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="bf795-165">Office. Context. Mailbox. Item. displayReplyFormAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-165">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bf795-166">Foi adicionada uma nova função ao `Item` objeto que exibe o formulário "responder" no modo de leitura.</span><span class="sxs-lookup"><span data-stu-id="bf795-166">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="bf795-167">Esta é a versão assíncrona do `displayReplyForm` método.</span><span class="sxs-lookup"><span data-stu-id="bf795-167">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="bf795-168">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-168">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="bf795-169">Ativação baseada em evento</span><span class="sxs-lookup"><span data-stu-id="bf795-169">Event-based activation</span></span>

<span data-ttu-id="bf795-170">Adicionado suporte à funcionalidade de ativação baseada em eventos em suplementos do Outlook. Confira [Configurar o suplemento do Outlook para](../../../outlook/autolaunch.md) obter mais informações sobre a ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="bf795-170">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="bf795-171">Ponto de extensão LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="bf795-171">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="bf795-172">Adicionado o `LaunchEvent` suporte a ponto de extensão ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="bf795-172">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="bf795-173">Ele configura a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="bf795-173">It configures event-based activation functionality.</span></span>

<span data-ttu-id="bf795-174">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bf795-174">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="bf795-175">Elemento de manifesto LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="bf795-175">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="bf795-176">`LaunchEvents`Elemento adicionado ao manifesto.</span><span class="sxs-lookup"><span data-stu-id="bf795-176">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="bf795-177">Ele oferece suporte à configuração da funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="bf795-177">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="bf795-178">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bf795-178">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="bf795-179">Elemento de manifesto de runtimes</span><span class="sxs-lookup"><span data-stu-id="bf795-179">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="bf795-180">Adicionado suporte do Outlook ao `Runtimes` elemento manifest.</span><span class="sxs-lookup"><span data-stu-id="bf795-180">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="bf795-181">Ele faz referência aos arquivos HTML e JavaScript necessários para a funcionalidade de ativação baseada em eventos.</span><span class="sxs-lookup"><span data-stu-id="bf795-181">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="bf795-182">**Disponível no**: Outlook na Web (moderno, [solicitar acesso de visualização](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="bf795-182">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="bf795-183">Obter todas as propriedades personalizadas</span><span class="sxs-lookup"><span data-stu-id="bf795-183">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="bf795-184">CustomProperties. getAll</span><span class="sxs-lookup"><span data-stu-id="bf795-184">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

<span data-ttu-id="bf795-185">Foi adicionada uma nova função ao `CustomProperties` objeto que obtém todas as propriedades personalizadas.</span><span class="sxs-lookup"><span data-stu-id="bf795-185">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="bf795-186">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno), Outlook no Mac (conectado a uma assinatura do Microsoft 365), Outlook no Android, Outlook no Ios</span><span class="sxs-lookup"><span data-stu-id="bf795-186">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="bf795-187">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="bf795-187">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="bf795-188">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-188">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bf795-189">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="bf795-189">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="bf795-190">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="bf795-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="bf795-191">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="bf795-191">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="bf795-192">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-192">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="bf795-193">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-193">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="bf795-194">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="bf795-195">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-195">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bf795-196">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-196">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="bf795-197">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="bf795-198">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-198">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="bf795-199">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-199">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="bf795-200">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-200">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="bf795-201">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="bf795-201">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="bf795-202">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-202">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="bf795-203">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-203">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="bf795-204">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="bf795-204">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="bf795-205">Adição de uma nova enumeração `ComposeType` disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-205">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="bf795-206">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno, [Configurar acesso de visualização](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="bf795-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="bf795-207">Mensagens de notificação com ações</span><span class="sxs-lookup"><span data-stu-id="bf795-207">Notification messages with actions</span></span>

<span data-ttu-id="bf795-208">Este recurso permite que o suplemento inclua uma mensagem de notificação com uma ação personalizada além da ação padrão de **ignorar** .</span><span class="sxs-lookup"><span data-stu-id="bf795-208">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="bf795-209">Office. NotificationMessageDetails. Actions</span><span class="sxs-lookup"><span data-stu-id="bf795-209">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="bf795-210">Adicionada uma nova propriedade que permite que você adicione uma `InsightMessage` notificação com uma ação personalizada.</span><span class="sxs-lookup"><span data-stu-id="bf795-210">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="bf795-211">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-211">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="bf795-212">Office. NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="bf795-212">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="bf795-213">Adicionado um novo objeto onde você define uma ação personalizada para sua `InsightMessage` notificação.</span><span class="sxs-lookup"><span data-stu-id="bf795-213">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="bf795-214">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-214">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="bf795-215">Office. MailboxEnums. ActionType</span><span class="sxs-lookup"><span data-stu-id="bf795-215">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="bf795-216">Foi adicionada uma nova enumeração `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="bf795-216">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="bf795-217">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-217">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="bf795-218">Office. MailboxEnums. ItemNotificationMessageType. InsightMessage</span><span class="sxs-lookup"><span data-stu-id="bf795-218">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="bf795-219">Adicionado um novo tipo `InsightMessage` à `ItemNotificationMessageType` enumeração.</span><span class="sxs-lookup"><span data-stu-id="bf795-219">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="bf795-220">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="bf795-220">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="bf795-221">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="bf795-221">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="bf795-222">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="bf795-222">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="bf795-223">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="bf795-223">Added ability to get Office theme.</span></span>

<span data-ttu-id="bf795-224">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-224">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="bf795-225">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="bf795-225">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="bf795-226">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="bf795-226">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="bf795-227">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-227">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="bf795-228">Os dados da sessão</span><span class="sxs-lookup"><span data-stu-id="bf795-228">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="bf795-229">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="bf795-229">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="bf795-230">Adicionado um novo objeto que representa os dados de sessão de um item.</span><span class="sxs-lookup"><span data-stu-id="bf795-230">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="bf795-231">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-231">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="bf795-232">Office. Context. Mailbox. Item. sessionData</span><span class="sxs-lookup"><span data-stu-id="bf795-232">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="bf795-233">Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="bf795-233">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="bf795-234">**Disponível no**: Outlook no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="bf795-234">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="bf795-235">SSO (logon único)</span><span class="sxs-lookup"><span data-stu-id="bf795-235">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="bf795-236">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="bf795-236">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="bf795-237">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="bf795-237">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="bf795-238">**Disponível em**: Outlook no Windows (conectado a uma assinatura do Microsoft 365), Outlook no Mac (conectado a uma assinatura do Microsoft 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="bf795-238">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="bf795-239">Confira também</span><span class="sxs-lookup"><span data-stu-id="bf795-239">See also</span></span>

- [<span data-ttu-id="bf795-240">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="bf795-240">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="bf795-241">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="bf795-241">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="bf795-242">Introdução</span><span class="sxs-lookup"><span data-stu-id="bf795-242">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="bf795-243">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="bf795-243">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
