---
title: Outlook conjunto de requisitos de visualização de API de complemento
description: Recursos e APIs que estão atualmente em visualização para Outlook de complementos.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 98bf56c169967ad7c994d1793afa8678d31f6892
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591055"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9dc5a-103">Outlook conjunto de requisitos de visualização de API de complemento</span><span class="sxs-lookup"><span data-stu-id="9dc5a-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="9dc5a-104">O Outlook de API de Office da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um Outlook de usuário.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9dc5a-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9dc5a-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9dc5a-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9dc5a-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="9dc5a-108">Você pode ser capaz de visualizar recursos em Outlook na Web configurando a versão direcionada em [seu locatário Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="9dc5a-109">"Configure preview access" é notado nesta página para recursos aplicáveis.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="9dc5a-110">Para outros recursos, você pode solicitar acesso aos bits de visualização para Outlook na Web usando sua conta Microsoft 365, concluindo e enviando [esse formulário](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="9dc5a-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="9dc5a-111">"Solicitar acesso de visualização" é notado nesses recursos.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="9dc5a-112">O conjunto de requisitos de visualização inclui todos os recursos do [conjunto de requisitos 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="9dc5a-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9dc5a-113">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="9dc5a-113">Features in preview</span></span>

<span data-ttu-id="9dc5a-114">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="9dc5a-115">Ativação do complemento em itens protegidos pelo IRM (Gerenciamento de Direitos de Informação)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="9dc5a-116">Os complementos agora podem ser ativados em itens protegidos por IRM.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="9dc5a-117">Para ativar esse recurso, um administrador de locatário precisa habilitar o direito de uso definindo a opção Permitir política personalizada de acesso `OBJMODEL` programático em  Office.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="9dc5a-118">Confira [Direitos de uso e descrições](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) para obter mais informações.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="9dc5a-119">**Disponível em**: Outlook no Windows, começando com a com build 13229.10000 (conectada a uma assinatura Microsoft 365 de terceiros)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="9dc5a-120">Propriedades de calendário adicionais</span><span class="sxs-lookup"><span data-stu-id="9dc5a-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="9dc5a-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9dc5a-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9dc5a-122">Adicionado um novo objeto que representa a propriedade de evento de todos os dias de um compromisso no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="9dc5a-123">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="9dc5a-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="9dc5a-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9dc5a-125">Adicionado um novo objeto que representa a sensibilidade de um compromisso no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="9dc5a-126">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="9dc5a-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9dc5a-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9dc5a-128">Adicionada uma nova propriedade que representa se um compromisso é um evento de dia inteiro.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="9dc5a-129">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="9dc5a-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="9dc5a-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9dc5a-131">Adicionada uma nova propriedade que representa a sensibilidade de um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="9dc5a-132">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="9dc5a-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="9dc5a-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9dc5a-134">Adicionado um novo número `AppointmentSensitivityType` que representa as opções de sensibilidade disponíveis em um compromisso.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="9dc5a-135">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9dc5a-136">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="9dc5a-136">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9dc5a-137">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9dc5a-137">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9dc5a-138">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="9dc5a-138">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9dc5a-139">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-139">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9dc5a-140">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="9dc5a-140">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9dc5a-141">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9dc5a-141">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9dc5a-142">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-142">Added ability to get Office theme.</span></span>

<span data-ttu-id="9dc5a-143">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9dc5a-144">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9dc5a-144">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9dc5a-145">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-145">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9dc5a-146">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-146">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="9dc5a-147">Os dados da sessão</span><span class="sxs-lookup"><span data-stu-id="9dc5a-147">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="9dc5a-148">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="9dc5a-148">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="9dc5a-149">Adicionado um novo objeto que representa os dados de sessão de um item.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-149">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="9dc5a-150">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-150">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="9dc5a-151">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="9dc5a-151">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9dc5a-152">Adicionada uma nova propriedade para gerenciar os dados de sessão de um item no modo Redação.</span><span class="sxs-lookup"><span data-stu-id="9dc5a-152">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="9dc5a-153">**Disponível em**: Outlook no Windows (conectado a uma assinatura Microsoft 365 de Microsoft 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="9dc5a-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="9dc5a-154">Confira também</span><span class="sxs-lookup"><span data-stu-id="9dc5a-154">See also</span></span>

- [<span data-ttu-id="9dc5a-155">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9dc5a-155">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9dc5a-156">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9dc5a-156">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9dc5a-157">Introdução</span><span class="sxs-lookup"><span data-stu-id="9dc5a-157">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9dc5a-158">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="9dc5a-158">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
