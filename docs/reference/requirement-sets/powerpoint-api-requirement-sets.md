---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do PowerPoint.
ms.date: 07/10/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: b2b5d4b7b5a0677812f227b6a32683c35bbf1662
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293503"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="df066-103">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="df066-103">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="df066-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="df066-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="df066-107">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos do cliente Office que oferecem suporte a esses conjuntos de requisitos e as versões de compilação ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="df066-107">The following table lists the PowerPoint requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="df066-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="df066-108">Requirement set</span></span>  |  <span data-ttu-id="df066-109">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="df066-109">Office on Windows</span></span><br><span data-ttu-id="df066-110">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="df066-110">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="df066-111">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="df066-111">Office on iPad</span></span><br><span data-ttu-id="df066-112">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="df066-112">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="df066-113">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="df066-113">Office on Mac</span></span><br><span data-ttu-id="df066-114">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="df066-114">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="df066-115">Office na Web</span><span class="sxs-lookup"><span data-stu-id="df066-115">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="df066-116">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="df066-116">PowerPointApi 1.1</span></span> | <span data-ttu-id="df066-117">Versão 1810 (Build 11001.20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="df066-117">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="df066-118">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="df066-118">2.17 or later</span></span> | <span data-ttu-id="df066-119">16.19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="df066-119">16.19 or later</span></span> | <span data-ttu-id="df066-120">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="df066-120">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="df066-121">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="df066-121">Office versions and build numbers</span></span>

<span data-ttu-id="df066-122">Para saber mais sobre as versões do Office e os números de build, confira:</span><span class="sxs-lookup"><span data-stu-id="df066-122">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="df066-123">API JavaScript do PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="df066-123">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="df066-124">A API JavaScript do PowerPoint 1.1 contém uma única API para criar uma nova apresentação.</span><span class="sxs-lookup"><span data-stu-id="df066-124">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="df066-125">Para obter detalhes sobre a API, confira [API JavaScript para o PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="df066-125">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="df066-126">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="df066-126">Runtime requirement support check</span></span>

<span data-ttu-id="df066-127">Durante o tempo de execução, os suplementos podem verificar se um determinado aplicativo oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação.</span><span class="sxs-lookup"><span data-stu-id="df066-127">At runtime, add-ins can check if a particular application supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="df066-128">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="df066-128">Manifest-based requirement support check</span></span>

<span data-ttu-id="df066-129">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar.</span><span class="sxs-lookup"><span data-stu-id="df066-129">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="df066-130">Se o aplicativo do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. elemento`Requirements`, o suplemento não será executado no aplicativo ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="df066-130">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="df066-131">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos do cliente Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="df066-131">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="df066-132">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="df066-132">Office Common API requirement sets</span></span>

<span data-ttu-id="df066-133">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="df066-133">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="df066-134">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="df066-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="df066-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="df066-135">See also</span></span>

- [<span data-ttu-id="df066-136">Documentação de Referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="df066-136">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="df066-137">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="df066-137">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="df066-138">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="df066-138">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="df066-139">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="df066-139">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
