---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: 4365dab3d8dd1ddb876536b3030926d68a89ac49
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605670"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9c6c0-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9c6c0-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9c6c0-103">O subconjunto de APIs de suplemento do Outlook da API JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9c6c0-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9c6c0-104">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9c6c0-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9c6c0-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="9c6c0-107">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="9c6c0-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9c6c0-108">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="9c6c0-108">Features in preview</span></span>

<span data-ttu-id="9c6c0-109">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-109">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="9c6c0-110">Anexar ao enviar</span><span class="sxs-lookup"><span data-stu-id="9c6c0-110">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="9c6c0-111">Office. Body. appendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="9c6c0-111">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="9c6c0-112">Foi adicionada uma nova função ao `Body` objeto que acrescenta dados ao final do corpo do item no modo de composição.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-112">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="9c6c0-113">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-113">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="9c6c0-114">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="9c6c0-114">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="9c6c0-115">Adicionado um novo elemento ao manifesto onde a `AppendOnSend` permissão estendida deve ser incluída na coleção de permissões estendidas.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-115">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="9c6c0-116">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-116">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9c6c0-117">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="9c6c0-117">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9c6c0-118">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9c6c0-118">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9c6c0-119">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="9c6c0-119">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9c6c0-120">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9c6c0-121">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="9c6c0-121">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9c6c0-122">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9c6c0-122">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9c6c0-123">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-123">Added ability to get Office theme.</span></span>

<span data-ttu-id="9c6c0-124">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9c6c0-125">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9c6c0-125">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9c6c0-126">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-126">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9c6c0-127">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="9c6c0-128">SSO</span><span class="sxs-lookup"><span data-stu-id="9c6c0-128">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="9c6c0-129">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="9c6c0-129">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="9c6c0-130">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](../../../outlook/authenticate-a-user-with-an-sso-token.md) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9c6c0-130">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="9c6c0-131">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="9c6c0-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="9c6c0-132">Confira também</span><span class="sxs-lookup"><span data-stu-id="9c6c0-132">See also</span></span>

- [<span data-ttu-id="9c6c0-133">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9c6c0-133">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9c6c0-134">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="9c6c0-134">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9c6c0-135">Introdução</span><span class="sxs-lookup"><span data-stu-id="9c6c0-135">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9c6c0-136">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="9c6c0-136">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
