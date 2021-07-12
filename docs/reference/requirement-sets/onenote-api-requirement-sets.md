---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: Saiba mais sobre os conjuntos de requisitos da API JavaScript do OneNote.
ms.date: 08/24/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: ecdb26edca54758540688ba03b1d9c1eec14e739
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350187"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="95157-103">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="95157-103">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="95157-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="95157-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="95157-107">A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos do cliente Office que oferecem suporte a esse conjunto de requisitos, e as versões de compilação ou data de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="95157-107">The following table lists the OneNote requirement sets, the Office client applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="95157-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="95157-108">Requirement set</span></span>  |  <span data-ttu-id="95157-109">Office na Web</span><span class="sxs-lookup"><span data-stu-id="95157-109">Office on the web</span></span> |
|:-----|:-----|
| [<span data-ttu-id="95157-110">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="95157-110">OneNoteApi 1.1</span></span>](/javascript/api/onenote?view=onenote-js-1.1&preserve-view=true)  | <span data-ttu-id="95157-111">Setembro de 2016</span><span class="sxs-lookup"><span data-stu-id="95157-111">September 2016</span></span> |  

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="95157-112">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="95157-112">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="95157-113">A OneNote JavaScript API 1.1 é a primeira versão da API.</span><span class="sxs-lookup"><span data-stu-id="95157-113">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="95157-114">Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](../../onenote/onenote-add-ins-programming-overview.md).</span><span class="sxs-lookup"><span data-stu-id="95157-114">For details about the API, see the [OneNote JavaScript API programming overview](../../onenote/onenote-add-ins-programming-overview.md).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="95157-115">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="95157-115">Runtime requirement support check</span></span>

<span data-ttu-id="95157-116">Durante o tempo de execução, os suplementos podem verificar se um determinado aplicativo do Office oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação:</span><span class="sxs-lookup"><span data-stu-id="95157-116">At runtime, add-ins can check if a particular Office application supports an API requirement set by doing the following:</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="95157-117">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="95157-117">Manifest-based requirement support check</span></span>

<span data-ttu-id="95157-118">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar.</span><span class="sxs-lookup"><span data-stu-id="95157-118">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="95157-119">Se o aplicativo do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no. elemento`Requirements`, o suplemento não será executado no aplicativo ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="95157-119">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="95157-120">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos do cliente Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="95157-120">The following code example shows an add-in that loads in all Office client applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="95157-121">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="95157-121">Office Common API requirement sets</span></span>

<span data-ttu-id="95157-122">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="95157-122">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="95157-123">Confira também</span><span class="sxs-lookup"><span data-stu-id="95157-123">See also</span></span>

- [<span data-ttu-id="95157-124">Documentação de Referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="95157-124">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="95157-125">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="95157-125">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="95157-126">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="95157-126">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="95157-127">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="95157-127">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
