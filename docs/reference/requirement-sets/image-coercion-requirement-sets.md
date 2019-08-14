---
title: Conjuntos de requisitos de coerção de imagem
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 08/13/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 9d622c827315f6657cf0fddaace33968bd634d64
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395670"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="44c60-103">Conjuntos de requisitos de coerção de imagem</span><span class="sxs-lookup"><span data-stu-id="44c60-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="44c60-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="44c60-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="44c60-107">ImageCoercion 1,1</span><span class="sxs-lookup"><span data-stu-id="44c60-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="44c60-108">ImageCoercion 1,1 permite a conversão para uma imagem`Office.CoercionType.Image`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="44c60-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="44c60-109">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="44c60-109">The following hosts are supported:</span></span>

- <span data-ttu-id="44c60-110">Excel 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="44c60-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="44c60-111">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="44c60-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="44c60-112">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="44c60-112">Excel on iPad</span></span>
- <span data-ttu-id="44c60-113">OneNote na Web</span><span class="sxs-lookup"><span data-stu-id="44c60-113">OneNote on the web</span></span>
- <span data-ttu-id="44c60-114">PowerPoint 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="44c60-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="44c60-115">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="44c60-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="44c60-116">PowerPoint na Web</span><span class="sxs-lookup"><span data-stu-id="44c60-116">PowerPoint on the web</span></span>
- <span data-ttu-id="44c60-117">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="44c60-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="44c60-118">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="44c60-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="44c60-119">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="44c60-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="44c60-120">Word na Web</span><span class="sxs-lookup"><span data-stu-id="44c60-120">Word on the web</span></span>
- <span data-ttu-id="44c60-121">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="44c60-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="44c60-122">ImageCoercion 1,2</span><span class="sxs-lookup"><span data-stu-id="44c60-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="44c60-123">ImageCoercion 1,2 permite conversão para o formato SVG`Office.CoercionType.XmlSvg`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="44c60-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="44c60-124">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="44c60-124">The following hosts are supported:</span></span>

- <span data-ttu-id="44c60-125">Excel no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-125">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-126">Excel no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-126">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-127">PowerPoint no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-127">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-128">PowerPoint no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-128">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-129">PowerPoint na Web</span><span class="sxs-lookup"><span data-stu-id="44c60-129">PowerPoint on the web</span></span>
- <span data-ttu-id="44c60-130">Word no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-130">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-131">Word no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44c60-131">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="44c60-132">Word na Web</span><span class="sxs-lookup"><span data-stu-id="44c60-132">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="44c60-133">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="44c60-133">Office Common API requirement sets</span></span>

<span data-ttu-id="44c60-134">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="44c60-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="44c60-135">Confira também</span><span class="sxs-lookup"><span data-stu-id="44c60-135">See also</span></span>

- [<span data-ttu-id="44c60-136">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="44c60-136">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="44c60-137">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="44c60-137">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="44c60-138">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="44c60-138">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
