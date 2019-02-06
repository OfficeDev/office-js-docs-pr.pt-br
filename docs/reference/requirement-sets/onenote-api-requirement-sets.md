---
title: Conjuntos de requisitos da API JavaScript do OneNote
description: ''
ms.date: 10/09/2018
ms.prod: onenote
localization_priority: Normal
ms.openlocfilehash: 3d274700d002963937294c8bcb235331d2829542
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742104"
---
# <a name="onenote-javascript-api-requirement-sets"></a><span data-ttu-id="5cd91-102">Conjuntos de requisitos da API JavaScript do OneNote</span><span class="sxs-lookup"><span data-stu-id="5cd91-102">OneNote JavaScript API requirement sets</span></span>

<span data-ttu-id="5cd91-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="5cd91-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="5cd91-106">A tabela a seguir lista os conjuntos de requisitos do OneNote, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou datas de disponibilidade.</span><span class="sxs-lookup"><span data-stu-id="5cd91-106">The following table lists the OneNote requirement sets, the Office host applications that support those requirement sets, and the build versions or availability date.</span></span>

|  <span data-ttu-id="5cd91-107">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cd91-107">Requirement set</span></span>  |  <span data-ttu-id="5cd91-108">Office Online</span><span class="sxs-lookup"><span data-stu-id="5cd91-108">Office Online</span></span> | 
|:-----|:-----|
| <span data-ttu-id="5cd91-109">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="5cd91-109">OneNoteApi 1.1</span></span>  | <span data-ttu-id="5cd91-110">Setembro de 2016</span><span class="sxs-lookup"><span data-stu-id="5cd91-110">September 2016</span></span> |  

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="5cd91-111">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="5cd91-111">Office Common API requirement sets</span></span>

<span data-ttu-id="5cd91-112">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="5cd91-112">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="onenote-javascript-api-11"></a><span data-ttu-id="5cd91-113">OneNote JavaScript API 1.1</span><span class="sxs-lookup"><span data-stu-id="5cd91-113">OneNote JavaScript API 1.1</span></span> 

<span data-ttu-id="5cd91-114">A OneNote JavaScript API 1.1 é a primeira versão da API.</span><span class="sxs-lookup"><span data-stu-id="5cd91-114">OneNote JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="5cd91-115">Para obter mais detalhes sobre a API, confira o artigo [Visão geral da programação da API JavaScript do OneNote](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span><span class="sxs-lookup"><span data-stu-id="5cd91-115">For details about the API, see the [OneNote JavaScript API programming overview](https://docs.microsoft.com/office/dev/add-ins/onenote/onenote-add-ins-programming-overview).</span></span>

## <a name="runtime-requirement-support-check"></a><span data-ttu-id="5cd91-116">Verificação do suporte a requisitos de tempo de execução</span><span class="sxs-lookup"><span data-stu-id="5cd91-116">Runtime requirement support check</span></span>

<span data-ttu-id="5cd91-117">Durante o tempo de execução, os suplementos podem verificar se um determinado host oferece suporte a um conjunto de requisitos de API, realizando a seguinte verificação:</span><span class="sxs-lookup"><span data-stu-id="5cd91-117">During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check:</span></span> 

```js
if (Office.context.requirements.isSetSupported('OneNoteApi', 1.1) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## <a name="manifest-based-requirement-support-check"></a><span data-ttu-id="5cd91-118">Verificação de suporte a requisitos com base em manifesto</span><span class="sxs-lookup"><span data-stu-id="5cd91-118">Manifest-based requirement support check</span></span>

<span data-ttu-id="5cd91-p103">Use o elemento Requirements no manifesto do suplemento para especificar conjuntos de requisitos ou membros de API cruciais que o seu suplemento precisa usar. Se o host do Office ou a plataforma não der suporte ao conjunto de requisitos ou membros da API especificados no elemento Requirements, o suplemento não será executado no host ou na plataforma e não será exibido em Meus Suplementos.</span><span class="sxs-lookup"><span data-stu-id="5cd91-p103">Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins.</span></span>

<span data-ttu-id="5cd91-121">O exemplo de código a seguir mostra um suplemento que é carregado em todos os aplicativos host do Office que oferecem suporte ao conjunto de requisitos OneNoteApi, versão 1.1.</span><span class="sxs-lookup"><span data-stu-id="5cd91-121">The following code example shows an add-in that loads in all Office host applications that support the OneNoteApi requirement set, version 1.1.</span></span>

```xml
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="OneNoteApi" MinVersion="1.1"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a><span data-ttu-id="5cd91-122">Confira também</span><span class="sxs-lookup"><span data-stu-id="5cd91-122">See also</span></span>

- [<span data-ttu-id="5cd91-123">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="5cd91-123">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="5cd91-124">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="5cd91-124">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="5cd91-125">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5cd91-125">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
