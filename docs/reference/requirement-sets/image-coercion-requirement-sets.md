---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 83817bfc7cf8a193138a805b0e90b4357d605801
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596967"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="8aaf0-103">Conjuntos de requisitos de Coerção de Imagens</span><span class="sxs-lookup"><span data-stu-id="8aaf0-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="8aaf0-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="8aaf0-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="8aaf0-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="8aaf0-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="8aaf0-108">ImageCoercion 1,1 permite a conversão para uma imagem`Office.CoercionType.Image`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="8aaf0-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="8aaf0-109">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="8aaf0-109">The following hosts are supported:</span></span>

- <span data-ttu-id="8aaf0-110">Excel 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="8aaf0-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="8aaf0-111">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="8aaf0-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="8aaf0-112">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="8aaf0-112">Excel on iPad</span></span>
- <span data-ttu-id="8aaf0-113">OneNote Online</span><span class="sxs-lookup"><span data-stu-id="8aaf0-113">OneNote on the web</span></span>
- <span data-ttu-id="8aaf0-114">PowerPoint 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="8aaf0-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="8aaf0-115">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="8aaf0-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="8aaf0-116">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="8aaf0-116">PowerPoint on the web</span></span>
- <span data-ttu-id="8aaf0-117">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="8aaf0-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="8aaf0-118">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="8aaf0-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="8aaf0-119">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="8aaf0-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="8aaf0-120">Word Online</span><span class="sxs-lookup"><span data-stu-id="8aaf0-120">Word on the web</span></span>
- <span data-ttu-id="8aaf0-121">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="8aaf0-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="8aaf0-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="8aaf0-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="8aaf0-123">ImageCoercion 1,2 permite conversão para o formato SVG`Office.CoercionType.XmlSvg`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="8aaf0-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="8aaf0-124">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="8aaf0-124">The following hosts are supported:</span></span>

- <span data-ttu-id="8aaf0-125">Excel no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-125">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-126">Excel no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-126">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-127">PowerPoint no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-127">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-128">PowerPoint no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-128">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-129">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="8aaf0-129">PowerPoint on the web</span></span>
- <span data-ttu-id="8aaf0-130">Word no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-130">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-131">Word no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8aaf0-131">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="8aaf0-132">Word Online</span><span class="sxs-lookup"><span data-stu-id="8aaf0-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="8aaf0-133">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="8aaf0-133">Office Common API requirement sets</span></span>

<span data-ttu-id="8aaf0-134">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="8aaf0-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8aaf0-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="8aaf0-135">See also</span></span>

- [<span data-ttu-id="8aaf0-136">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="8aaf0-136">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="8aaf0-137">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="8aaf0-137">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="8aaf0-138">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8aaf0-138">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
