---
title: Conjuntos de requisitos da API JavaScript do PowerPoint
description: ''
ms.date: 07/26/2019
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 5bba2354cabba3c3ccd4ddf38d3e03c25a32b8a9
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950954"
---
# <a name="powerpoint-javascript-api-requirement-sets"></a><span data-ttu-id="0a729-102">Conjuntos de requisitos da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0a729-102">PowerPoint JavaScript API requirement sets</span></span>

<span data-ttu-id="0a729-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="0a729-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="0a729-106">A tabela a seguir lista os conjuntos de requisitos do PowerPoint, os aplicativos de host do Office que oferecem suporte a esses conjuntos de requisitos e os versões de compilação ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="0a729-106">The following table lists the PowerPoint requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="0a729-107">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="0a729-107">Requirement set</span></span>  |  <span data-ttu-id="0a729-108">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="0a729-108">Office on Windows</span></span><br><span data-ttu-id="0a729-109">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="0a729-109">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="0a729-110">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="0a729-110">Office on iPad</span></span><br><span data-ttu-id="0a729-111">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="0a729-111">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="0a729-112">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="0a729-112">Office on Mac</span></span><br><span data-ttu-id="0a729-113">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="0a729-113">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="0a729-114">Office na Web</span><span class="sxs-lookup"><span data-stu-id="0a729-114">Office on the web</span></span> |
|:-----|-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="0a729-115">PowerPointApi 1.1</span><span class="sxs-lookup"><span data-stu-id="0a729-115">PowerPointApi 1.1</span></span> | <span data-ttu-id="0a729-116">Versão 1810 (Build 11001.20074) ou posterior</span><span class="sxs-lookup"><span data-stu-id="0a729-116">Version 1810 (Build 11001.20074) or later</span></span> | <span data-ttu-id="0a729-117">2.17 ou posterior</span><span class="sxs-lookup"><span data-stu-id="0a729-117">2.17 or later</span></span> | <span data-ttu-id="0a729-118">16.19 ou posterior</span><span class="sxs-lookup"><span data-stu-id="0a729-118">16.19 or later</span></span> | <span data-ttu-id="0a729-119">Outubro de 2018</span><span class="sxs-lookup"><span data-stu-id="0a729-119">October 2018</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="0a729-120">Versões do Office e números de build</span><span class="sxs-lookup"><span data-stu-id="0a729-120">Office versions and build numbers</span></span>

<span data-ttu-id="0a729-121">Para saber mais sobre as versões do Office e os números de build, confira:</span><span class="sxs-lookup"><span data-stu-id="0a729-121">For more information about Office versions and build numbers, see:</span></span>

- <span data-ttu-id="0a729-122">
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="0a729-122">[Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- [<span data-ttu-id="0a729-123">Qual versão do Office estou usando?</span><span class="sxs-lookup"><span data-stu-id="0a729-123">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- <span data-ttu-id="0a729-124">
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="0a729-124">[Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>

## <a name="powerpoint-javascript-api-11"></a><span data-ttu-id="0a729-125">API JavaScript do PowerPoint 1.1</span><span class="sxs-lookup"><span data-stu-id="0a729-125">PowerPoint JavaScript API 1.1</span></span>

<span data-ttu-id="0a729-126">A API JavaScript do PowerPoint 1.1 contém uma única API para criar uma nova apresentação.</span><span class="sxs-lookup"><span data-stu-id="0a729-126">PowerPoint JavaScript API 1.1 contains a single API to create a new presentation.</span></span> <span data-ttu-id="0a729-127">Para obter detalhes sobre a API, confira [API JavaScript para o PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="0a729-127">For details about the API, see [JavaScript API for PowerPoint](../../powerpoint/powerpoint-add-ins.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="0a729-128">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="0a729-128">Runtime requirement support check</span></span>

<span data-ttu-id="0a729-129">Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação.</span><span class="sxs-lookup"><span data-stu-id="0a729-129">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="0a729-130">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="0a729-130">Manifest-based requirement support check</span></span>

<span data-ttu-id="0a729-131">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar.</span><span class="sxs-lookup"><span data-stu-id="0a729-131">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="0a729-132">Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. `Requirements` elemento, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="0a729-132">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="0a729-133">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="0a729-133">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="0a729-134">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="0a729-134">Office Common API requirement sets</span></span>

<span data-ttu-id="0a729-135">A maior parte da funcionalidade do suplemento do PowerPoint vem do conjunto de APIs comuns.</span><span class="sxs-lookup"><span data-stu-id="0a729-135">Most of the PowerPoint Add-in functionality comes from the Common API set.</span></span> <span data-ttu-id="0a729-136">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="0a729-136">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0a729-137">Confira também</span><span class="sxs-lookup"><span data-stu-id="0a729-137">See also</span></span>

- [<span data-ttu-id="0a729-138">Documentação de Referência da API JavaScript do PowerPoint</span><span class="sxs-lookup"><span data-stu-id="0a729-138">PowerPoint JavaScript API reference documentation</span></span>](/javascript/api/powerpoint)
- [<span data-ttu-id="0a729-139">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="0a729-139">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="0a729-140">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="0a729-140">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="0a729-141">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="0a729-141">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
