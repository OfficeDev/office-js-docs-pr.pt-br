---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 59f6891182f47bed1b7e3b6aa69a30e941bce7cb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094348"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="37f3c-103">Conjuntos de requisitos de Coerção de Imagens</span><span class="sxs-lookup"><span data-stu-id="37f3c-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="37f3c-104">Requirement sets are named groups of API members.</span><span class="sxs-lookup"><span data-stu-id="37f3c-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="37f3c-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span><span class="sxs-lookup"><span data-stu-id="37f3c-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="37f3c-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="37f3c-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="37f3c-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="37f3c-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="37f3c-108">ImageCoercion 1,1 permite a conversão para uma imagem ( `Office.CoercionType.Image` ) ao gravar dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="37f3c-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="37f3c-109">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="37f3c-109">The following hosts are supported:</span></span>

- <span data-ttu-id="37f3c-110">Excel 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="37f3c-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="37f3c-111">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="37f3c-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="37f3c-112">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="37f3c-112">Excel on iPad</span></span>
- <span data-ttu-id="37f3c-113">OneNote Online</span><span class="sxs-lookup"><span data-stu-id="37f3c-113">OneNote on the web</span></span>
- <span data-ttu-id="37f3c-114">PowerPoint 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="37f3c-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="37f3c-115">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="37f3c-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="37f3c-116">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="37f3c-116">PowerPoint on the web</span></span>
- <span data-ttu-id="37f3c-117">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="37f3c-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="37f3c-118">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="37f3c-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="37f3c-119">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="37f3c-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="37f3c-120">Word Online</span><span class="sxs-lookup"><span data-stu-id="37f3c-120">Word on the web</span></span>
- <span data-ttu-id="37f3c-121">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="37f3c-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="37f3c-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="37f3c-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="37f3c-123">ImageCoercion 1,2 permite conversão para o formato SVG ( `Office.CoercionType.XmlSvg` ) ao gravar dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="37f3c-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="37f3c-124">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="37f3c-124">The following hosts are supported:</span></span>

- <span data-ttu-id="37f3c-125">Excel no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-126">Excel no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-127">PowerPoint no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-128">PowerPoint no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-129">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="37f3c-129">PowerPoint on the web</span></span>
- <span data-ttu-id="37f3c-130">Word no Windows (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-131">Word no Mac (conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37f3c-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="37f3c-132">Word Online</span><span class="sxs-lookup"><span data-stu-id="37f3c-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="37f3c-133">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="37f3c-133">Office Common API requirement sets</span></span>

<span data-ttu-id="37f3c-134">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="37f3c-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="37f3c-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="37f3c-135">See also</span></span>

- [<span data-ttu-id="37f3c-136">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="37f3c-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="37f3c-137">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="37f3c-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="37f3c-138">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="37f3c-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
