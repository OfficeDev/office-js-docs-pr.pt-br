---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: ''
ms.date: 07/17/2019
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: e1012b337b3713f57a5d3df7f7c7ccbcf509b5aa
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940841"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="ee4ac-102">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ee4ac-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="ee4ac-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ee4ac-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="ee4ac-106">A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="ee4ac-107">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="ee4ac-107">Requirement set</span></span>  |  <span data-ttu-id="ee4ac-108">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ee4ac-108">Office on the web</span></span> |
|:-----|:-----|
| <span data-ttu-id="ee4ac-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="ee4ac-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="ee4ac-110">Setembro de 2016</span><span class="sxs-lookup"><span data-stu-id="ee4ac-110">September 2016</span></span> |

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="ee4ac-111">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="ee4ac-111">OneNote JavaScript API 1.1</span></span>

<span data-ttu-id="ee4ac-112">A OneNote JavaScript API 1.1 é a primeira versão da API.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-112">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="ee4ac-113">Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="ee4ac-113">For details about the API, see the [OneNote JavaScript API programming overview](/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="ee4ac-114">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="ee4ac-114">Runtime requirement support check</span></span>

<span data-ttu-id="ee4ac-115">No tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, fazendo o seguinte.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-115">At runtime, add-ins can check if a particular host supports an API requirement set by doing the following.</span></span>

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', '1.1') === true) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="ee4ac-116">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="ee4ac-116">Manifest-based requirement support check</span></span>

<span data-ttu-id="ee4ac-117">Use o `Requirements` elemento no manifesto do suplemento para especificar conjuntos de requisitos críticos ou membros da API que seu suplemento deve usar.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-117">Use the `Requirements` element in the add-in manifest to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="ee4ac-118">Se o host ou a plataforma do Office não oferecer suporte aos conjuntos de requisitos ou membros `Requirements` de API especificados no elemento, o suplemento não será executado nesse host ou plataforma e não será exibido em meus suplementos.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-118">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="ee4ac-119">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="ee4ac-119">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="ee4ac-120">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="ee4ac-120">Office Common API requirement sets</span></span>

<span data-ttu-id="ee4ac-121">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="ee4ac-121">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ee4ac-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="ee4ac-122">See also</span></span>

- [<span data-ttu-id="ee4ac-123">Documentação de referência da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="ee4ac-123">OneNote JavaScript API reference documentation</span></span>](/javascript/api/onenote)
- [<span data-ttu-id="ee4ac-124">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="ee4ac-124">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="ee4ac-125">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="ee4ac-125">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="ee4ac-126">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ee4ac-126">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
