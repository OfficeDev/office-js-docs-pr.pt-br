---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: Recursos e APIs que estão atualmente em versão prévia para suplementos do Outlook e as APIs JavaScript do Office.
ms.date: 03/26/2020
localization_priority: Normal
ms.openlocfilehash: 55de284932a53d2226258a15c86ead4f05361c30
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978617"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="6d27f-103">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="6d27f-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="6d27f-104">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="6d27f-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6d27f-105">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="6d27f-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="6d27f-106">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="6d27f-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="6d27f-107">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="6d27f-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="6d27f-108">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="6d27f-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="6d27f-109">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="6d27f-109">Features in preview</span></span>

<span data-ttu-id="6d27f-110">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="6d27f-110">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="6d27f-111">Anexar ao enviar</span><span class="sxs-lookup"><span data-stu-id="6d27f-111">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="6d27f-112">Office. Context. Mailbox. Item. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-112">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="6d27f-113">Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-113">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="6d27f-114">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="6d27f-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="6d27f-115">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="6d27f-115">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="6d27f-116">Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="6d27f-116">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="6d27f-117">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="6d27f-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="6d27f-118">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="6d27f-118">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="6d27f-119">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-119">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d27f-120">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="6d27f-120">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="6d27f-121">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="6d27f-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="6d27f-122">Assinatura de email</span><span class="sxs-lookup"><span data-stu-id="6d27f-122">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="6d27f-123">Office. Context. Mailbox. Item. Body. setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-123">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="6d27f-124">Foi adicionada uma nova função ao `Body` objeto que adiciona ou substitui a assinatura no corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-124">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="6d27f-125">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="6d27f-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="6d27f-126">Office. Context. Mailbox. Item. disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-126">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d27f-127">Adicionada uma nova função que desabilita a assinatura do cliente para a caixa de correio de envio no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-127">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="6d27f-128">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="6d27f-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="6d27f-129">Office. Context. Mailbox. Item. getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-129">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="6d27f-130">Foi adicionada uma nova função que obtém o tipo de redação de uma mensagem no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-130">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="6d27f-131">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="6d27f-131">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="6d27f-132">Office. Context. Mailbox. Item. isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="6d27f-132">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d27f-133">Adicionada uma nova função que verifica se a assinatura do cliente está habilitada no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-133">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="6d27f-134">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="6d27f-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="6d27f-135">Office. MailboxEnums. composetype</span><span class="sxs-lookup"><span data-stu-id="6d27f-135">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="6d27f-136">Adição de uma nova `ComposeType` Enumeração disponível no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="6d27f-136">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="6d27f-137">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (moderno)</span><span class="sxs-lookup"><span data-stu-id="6d27f-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="6d27f-138">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="6d27f-138">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="6d27f-139">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="6d27f-139">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="6d27f-140">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="6d27f-140">Added ability to get Office theme.</span></span>

<span data-ttu-id="6d27f-141">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="6d27f-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="6d27f-142">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="6d27f-142">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="6d27f-143">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="6d27f-143">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="6d27f-144">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="6d27f-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="6d27f-145">SSO</span><span class="sxs-lookup"><span data-stu-id="6d27f-145">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="6d27f-146">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="6d27f-146">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="6d27f-147">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="6d27f-147">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="6d27f-148">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="6d27f-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="6d27f-149">Confira também</span><span class="sxs-lookup"><span data-stu-id="6d27f-149">See also</span></span>

- [<span data-ttu-id="6d27f-150">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="6d27f-150">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="6d27f-151">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="6d27f-151">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="6d27f-152">Introdução</span><span class="sxs-lookup"><span data-stu-id="6d27f-152">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="6d27f-153">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="6d27f-153">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
