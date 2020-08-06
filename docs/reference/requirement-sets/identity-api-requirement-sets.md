---
title: Conjuntos de requisitos da API de Identidade
description: Identity API requirements define informações para suplementos do Office.
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 05805451f17cc70597a61e55d1ecacbb81c383c5
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573214"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="3220c-103">Conjuntos de requisitos da API de Identidade</span><span class="sxs-lookup"><span data-stu-id="3220c-103">Identity API requirement sets</span></span>

<span data-ttu-id="3220c-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3220c-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3220c-107">Os suplementos do Office executam várias versões do Office.</span><span class="sxs-lookup"><span data-stu-id="3220c-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="3220c-108">A tabela a seguir lista os conjuntos de requisitos da API de Identidade, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="3220c-108">The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="3220c-109">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="3220c-109">Requirement set</span></span>  | <span data-ttu-id="3220c-110">Office 2013 ou posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="3220c-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="3220c-111">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3220c-111">(one-time purchase)</span></span> | <span data-ttu-id="3220c-112">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3220c-112">Office on Windows</span></span><br><span data-ttu-id="3220c-113">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3220c-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="3220c-114">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3220c-114">Office on iPad</span></span><br><span data-ttu-id="3220c-115">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3220c-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="3220c-116">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3220c-116">Office on Mac</span></span><br><span data-ttu-id="3220c-117">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3220c-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="3220c-118">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3220c-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="3220c-119">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="3220c-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="3220c-120">N/A</span><span class="sxs-lookup"><span data-stu-id="3220c-120">N/A</span></span> | <span data-ttu-id="3220c-121">2008 (Build 13127,20000) ou posterior</span><span class="sxs-lookup"><span data-stu-id="3220c-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="3220c-122">Em breve</span><span class="sxs-lookup"><span data-stu-id="3220c-122">Coming soon</span></span> | <span data-ttu-id="3220c-123">16,40 ou posterior</span><span class="sxs-lookup"><span data-stu-id="3220c-123">16.40 or later</span></span> | <span data-ttu-id="3220c-124">Agosto de 2020 \*</span><span class="sxs-lookup"><span data-stu-id="3220c-124">August, 2020\*</span></span> |

> <span data-ttu-id="3220c-125">\*Inicialmente, o conjunto de requisitos é suportado no Office na Web somente para documentos que são abertos a partir do SharePoint Online e do OneDrive.com.</span><span class="sxs-lookup"><span data-stu-id="3220c-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="3220c-126">O suporte para outros documentos será colocado no Office na Web mais tarde no 2020.</span><span class="sxs-lookup"><span data-stu-id="3220c-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="3220c-127">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="3220c-127">Office versions and build numbers</span></span>

<span data-ttu-id="3220c-128">Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:</span><span class="sxs-lookup"><span data-stu-id="3220c-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="3220c-129">Visão geral sobre o Servidor do Office Online</span><span class="sxs-lookup"><span data-stu-id="3220c-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3220c-130">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="3220c-130">Office Common API requirement sets</span></span>

<span data-ttu-id="3220c-131">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3220c-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="3220c-132">Visualização do IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="3220c-132">IdentityAPI Preview</span></span>

<span data-ttu-id="3220c-133">Para obter detalhes sobre essa API, consulte a versão que usa promessas em [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) ou a versão que usa retornos de chamada em [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="3220c-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="3220c-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="3220c-134">See also</span></span>

- [<span data-ttu-id="3220c-135">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="3220c-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3220c-136">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="3220c-136">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3220c-137">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3220c-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
