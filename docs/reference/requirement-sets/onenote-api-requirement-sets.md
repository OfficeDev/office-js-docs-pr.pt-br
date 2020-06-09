---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do OneNote
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 1adc3554cfce5cafa94afefdb1f2a2130817288e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611341"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="913e0-103">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="913e0-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="913e0-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="913e0-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="913e0-107">A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="913e0-107">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="913e0-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="913e0-108">Requirement set</span></span>  |  <span data-ttu-id="913e0-109">Office na Web</span><span class="sxs-lookup"><span data-stu-id="913e0-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="913e0-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="913e0-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1)  | <span data-ttu-id="913e0-111">Setembro de 2016</span><span class="sxs-lookup"><span data-stu-id="913e0-111">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="913e0-112">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="913e0-112">Office Common API requirement sets</span></span>

<span data-ttu-id="913e0-113">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="913e0-113">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="913e0-114">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="913e0-114">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="913e0-115">A OneNote JavaScript API 1.1 é a primeira versão da API.</span><span class="sxs-lookup"><span data-stu-id="913e0-115">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="913e0-116">Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="913e0-116">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="913e0-117">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="913e0-117">Runtime requirement support check</span></span>

<span data-ttu-id="913e0-118">Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação.</span><span class="sxs-lookup"><span data-stu-id="913e0-118">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="913e0-119">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="913e0-119">Manifest-based requirement support check</span></span>

<span data-ttu-id="913e0-120">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar.</span><span class="sxs-lookup"><span data-stu-id="913e0-120">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="913e0-121">Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. `Requirements` elemento, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="913e0-121">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="913e0-122">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="913e0-122">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="913e0-123">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="913e0-123">Office Common API requirement sets</span></span>

<span data-ttu-id="913e0-124">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="913e0-124">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="913e0-125">Confira também</span><span class="sxs-lookup"><span data-stu-id="913e0-125">See also</span></span>

- [<span data-ttu-id="913e0-126">Documentação de Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="913e0-126">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="913e0-127">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="913e0-127">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="913e0-128">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="913e0-128">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="913e0-129">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="913e0-129">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
