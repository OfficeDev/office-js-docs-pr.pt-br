---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: ''
ms.date: 03/11/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: ef76077c3a2a975fae8a0dc101e8e1b42ef66094
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600694"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="b1ab7-102">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b1ab7-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="b1ab7-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b1ab7-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="b1ab7-106">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos de host do Office que oferecem suporte a esses conjuntos de requisitos e os versões de compilação ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="b1ab7-107">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ab7-107">Requirement set</span></span>  |  <span data-ttu-id="b1ab7-108">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b1ab7-108">Office on Windows</span></span><br><span data-ttu-id="b1ab7-109">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1ab7-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="b1ab7-110">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b1ab7-110">Office on iPad</span></span><br><span data-ttu-id="b1ab7-111">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1ab7-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="b1ab7-112">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b1ab7-112">Office on Mac</span></span><br><span data-ttu-id="b1ab7-113">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b1ab7-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="b1ab7-114">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b1ab7-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="b1ab7-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="b1ab7-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="b1ab7-116">Versão 1810 (Build 11001.20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="b1ab7-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="b1ab7-117">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b1ab7-117">2.17 or later</span></span> | <span data-ttu-id="b1ab7-118">16.19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b1ab7-118">16.19 or later</span></span> | <span data-ttu-id="b1ab7-119">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="b1ab7-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="b1ab7-120">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="b1ab7-120">Office versions and build numbers</span></span>

<span data-ttu-id="b1ab7-121">Para saber mais sobre as versões do Office e os números de build, confira:</span><span class="sxs-lookup"><span data-stu-id="b1ab7-121">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="b1ab7-122">API JavaScript do PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="b1ab7-122">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="b1ab7-123">A API JavaScript do PowerPoint 1.1 contém uma única API para criar uma nova apresentação.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-123">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="b1ab7-124">Para obter detalhes sobre a API, confira [API JavaScript para o PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b1ab7-124">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="b1ab7-125">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="b1ab7-125">Runtime requirement support check</span></span>

<span data-ttu-id="b1ab7-126">Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-126">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="b1ab7-127">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="b1ab7-127">Manifest-based requirement support check</span></span>

<span data-ttu-id="b1ab7-128">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-128">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="b1ab7-129">Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. `Requirements` elemento, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-129">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="b1ab7-130">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-130">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b1ab7-131">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="b1ab7-131">Office Common API requirement sets</span></span>

<span data-ttu-id="b1ab7-132">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="b1ab7-132">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="b1ab7-133">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b1ab7-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="b1ab7-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="b1ab7-134">See also</span></span>

- [<span data-ttu-id="b1ab7-135">Documentação de Referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b1ab7-135">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="b1ab7-136">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="b1ab7-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b1ab7-137">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="b1ab7-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="b1ab7-138">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b1ab7-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
