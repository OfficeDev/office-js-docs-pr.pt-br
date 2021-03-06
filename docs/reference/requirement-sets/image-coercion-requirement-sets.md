---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de coerção de imagem com os complementos do Office no Excel, no PowerPoint e no Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 52ce46a46580500f5a292bf898674d4798378319
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505525"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="ce09d-103">Conjuntos de requisitos de Coerção de Imagens</span><span class="sxs-lookup"><span data-stu-id="ce09d-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="ce09d-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="ce09d-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="ce09d-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="ce09d-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="ce09d-108">ImageCoercion 1.1 permite a conversão em uma imagem ( ) ao escrever `Office.CoercionType.Image` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="ce09d-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="ce09d-109">Os seguintes aplicativos são suportados:</span><span class="sxs-lookup"><span data-stu-id="ce09d-109">The following applications are supported:</span></span>

- <span data-ttu-id="ce09d-110">Excel 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="ce09d-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="ce09d-111">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="ce09d-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="ce09d-112">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="ce09d-112">Excel on iPad</span></span>
- <span data-ttu-id="ce09d-113">OneNote Online</span><span class="sxs-lookup"><span data-stu-id="ce09d-113">OneNote on the web</span></span>
- <span data-ttu-id="ce09d-114">PowerPoint 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="ce09d-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="ce09d-115">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="ce09d-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="ce09d-116">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="ce09d-116">PowerPoint on the web</span></span>
- <span data-ttu-id="ce09d-117">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="ce09d-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="ce09d-118">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="ce09d-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="ce09d-119">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="ce09d-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="ce09d-120">Word Online</span><span class="sxs-lookup"><span data-stu-id="ce09d-120">Word on the web</span></span>
- <span data-ttu-id="ce09d-121">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="ce09d-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="ce09d-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="ce09d-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="ce09d-123">ImageCoercion 1.2 permite a conversão para o formato SVG ( ) ao escrever `Office.CoercionType.XmlSvg` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="ce09d-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="ce09d-124">Os seguintes aplicativos são suportados:</span><span class="sxs-lookup"><span data-stu-id="ce09d-124">The following applications are supported:</span></span>

- <span data-ttu-id="ce09d-125">Excel no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ce09d-126">Excel no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ce09d-127">PowerPoint no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ce09d-128">PowerPoint no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ce09d-129">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="ce09d-129">PowerPoint on the web</span></span>
- <span data-ttu-id="ce09d-130">Word no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="ce09d-131">Word no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ce09d-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="ce09d-132">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="ce09d-132">Office Common API requirement sets</span></span>

<span data-ttu-id="ce09d-133">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="ce09d-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ce09d-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="ce09d-134">See also</span></span>

- [<span data-ttu-id="ce09d-135">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="ce09d-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ce09d-136">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="ce09d-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="ce09d-137">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ce09d-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
