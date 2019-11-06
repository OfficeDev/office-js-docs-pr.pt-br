---
title: Conjunto de requisitos do modo de visualização de API para suplementos do Outlook
description: ''
ms.date: 11/05/2019
localization_priority: Priority
ms.openlocfilehash: f52edea26eba6c10b24f14feb1e86d811642798b
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001618"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="af6bd-102">Conjunto de requisitos do modo de visualização de API para suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="af6bd-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="af6bd-103">O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.</span><span class="sxs-lookup"><span data-stu-id="af6bd-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="af6bd-104">Esta documentação destina-se a um modo de **visualização** de [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="af6bd-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="af6bd-105">Esse conjunto de requisitos ainda não está totalmente implementado e os clientes não informarão precisamente o suporte para ele.</span><span class="sxs-lookup"><span data-stu-id="af6bd-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="af6bd-106">Você não deve especificar a esse conjunto de requisitos em seu manifesto de suplemento.</span><span class="sxs-lookup"><span data-stu-id="af6bd-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="af6bd-107">O conjunto de requisitos de visualização inclui todos os recursos do [Conjunto de requisitos 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="af6bd-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="af6bd-108">Recursos no modo de visualização</span><span class="sxs-lookup"><span data-stu-id="af6bd-108">Features in preview</span></span>

<span data-ttu-id="af6bd-109">Os seguintes recursos estão no modo de visualização.</span><span class="sxs-lookup"><span data-stu-id="af6bd-109">The following features are in preview.</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="af6bd-110">Integração à mensagens acionáveis</span><span class="sxs-lookup"><span data-stu-id="af6bd-110">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="af6bd-111">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="af6bd-111">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="af6bd-112">Adicionada uma nova função que retorna os dados inicialização que são transmitidos quando o suplemento é [ativado por uma mensagem acionável](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="af6bd-112">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="af6bd-113">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="af6bd-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="af6bd-114">Tema do Office</span><span class="sxs-lookup"><span data-stu-id="af6bd-114">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="af6bd-115">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="af6bd-115">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="af6bd-116">Capacidade adicional para obter o tema do Office.</span><span class="sxs-lookup"><span data-stu-id="af6bd-116">Added ability to get Office theme.</span></span>

<span data-ttu-id="af6bd-117">**Disponível em**: Outlook no Windows (conectado a assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="af6bd-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="af6bd-118">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="af6bd-118">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="af6bd-119">Adicionado `OfficeThemeChanged` evento `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="af6bd-119">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="af6bd-120">**Disponível em**: Outlook no Windows (conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="af6bd-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="af6bd-121">SSO</span><span class="sxs-lookup"><span data-stu-id="af6bd-121">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstokenofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="af6bd-122">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="af6bd-122">OfficeRuntime.auth.getAccessToken</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="af6bd-123">Foi adicionado acesso ao `getAccessToken`, que permite que os suplementos [obtenham um token de acesso](/outlook/add-ins/authenticate-a-user-with-an-sso-token) da API do Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="af6bd-123">Added access to `getAccessToken`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="af6bd-124">**Disponível em:** Outlook no Windows (conectado à assinatura do Office 365), Outlook para Mac (conectado à assinatura do Office 365), Outlook na Web (moderno), Outlook na Web (clássico)</span><span class="sxs-lookup"><span data-stu-id="af6bd-124">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="af6bd-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="af6bd-125">See also</span></span>

- [<span data-ttu-id="af6bd-126">Suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="af6bd-126">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="af6bd-127">Exemplos de código de suplementos do Outlook</span><span class="sxs-lookup"><span data-stu-id="af6bd-127">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="af6bd-128">Introdução</span><span class="sxs-lookup"><span data-stu-id="af6bd-128">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="af6bd-129">Conjuntos de requisitos e clientes com suporte</span><span class="sxs-lookup"><span data-stu-id="af6bd-129">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
