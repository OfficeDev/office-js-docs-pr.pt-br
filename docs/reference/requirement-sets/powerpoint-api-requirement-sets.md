---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Normal
ms.openlocfilehash: 4f64654a4130cc0d4bf96d9c59e364e77c808748
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35941141"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="a7979-102">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a7979-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="a7979-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="a7979-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="a7979-106">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos host do Office que dão suporte a esses conjuntos de requisitos e às versões de compilação ou data de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="a7979-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="a7979-107">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="a7979-107">Requirement set</span></span>  |  <span data-ttu-id="a7979-108">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="a7979-108">Office on Windows</span></span><br><span data-ttu-id="a7979-109">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a7979-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a7979-110">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="a7979-110">Office on iPad</span></span><br><span data-ttu-id="a7979-111">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a7979-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a7979-112">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="a7979-112">Office on Mac</span></span><br><span data-ttu-id="a7979-113">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a7979-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="a7979-114">Office na Web</span><span class="sxs-lookup"><span data-stu-id="a7979-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="a7979-115">PowerPointApi 1,1</span><span class="sxs-lookup"><span data-stu-id="a7979-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="a7979-116">Versão 1810 (Build 11001,20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="a7979-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="a7979-117">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="a7979-117">2.17 or later</span></span> | <span data-ttu-id="a7979-118">16,19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="a7979-118">16.19 or later</span></span> | <span data-ttu-id="a7979-119">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="a7979-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="a7979-120">Versões do Office e números de compilação</span><span class="sxs-lookup"><span data-stu-id="a7979-120">Office versions and build numbers</span></span>

<span data-ttu-id="a7979-121">Para obter mais informações sobre versões e números de compilação do Office, consulte:</span><span class="sxs-lookup"><span data-stu-id="a7979-121">For more information about Office versions and build numbers, see:</span></span>

- <span data-ttu-id="a7979-122">
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="a7979-122">[Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- [<span data-ttu-id="a7979-123">Qual versão do Office estou usando?</span><span class="sxs-lookup"><span data-stu-id="a7979-123">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- <span data-ttu-id="a7979-124">
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="a7979-124">[Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="a7979-125">API JavaScript do PowerPoint 1,1</span><span class="sxs-lookup"><span data-stu-id="a7979-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="a7979-126">A API JavaScript do PowerPoint 1,1 contém uma única API para criar uma nova apresentação.</span><span class="sxs-lookup"><span data-stu-id="a7979-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="a7979-127">Para obter detalhes sobre a API, consulte [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="a7979-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="a7979-128">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="a7979-128">Runtime requirement support check</span></span>

<span data-ttu-id="a7979-129">No tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, fazendo o seguinte.</span><span class="sxs-lookup"><span data-stu-id="a7979-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="a7979-130">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="a7979-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="a7979-131">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos críticos ou membros da API que seu suplemento deve usar.</span><span class="sxs-lookup"><span data-stu-id="a7979-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="a7979-132">Se o host ou a plataforma do Office não oferecer suporte aos conjuntos de requisitos ou membros `Requirements` de API especificados no elemento, o suplemento não será executado nesse host ou plataforma e não será exibido em meus suplementos.</span><span class="sxs-lookup"><span data-stu-id="a7979-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="a7979-133">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="a7979-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a7979-134">Conjuntos de requisitos da API Comum do Office</span><span class="sxs-lookup"><span data-stu-id="a7979-134">Office Common API requirement sets</span></span>

<span data-ttu-id="a7979-135">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="a7979-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="a7979-136">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a7979-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a7979-137">Confira também</span><span class="sxs-lookup"><span data-stu-id="a7979-137">See also</span></span>

- [<span data-ttu-id="a7979-138">Documentação de referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a7979-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="a7979-139">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="a7979-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a7979-140">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="a7979-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="a7979-141">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a7979-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
