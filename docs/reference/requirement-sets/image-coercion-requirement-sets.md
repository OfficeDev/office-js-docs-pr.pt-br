---
title: Conjuntos de requisitos de coerção de imagem
description: Suporte para conjuntos de requisitos de coerção de imagens com suplementos do Office no Excel, PowerPoint e Word.
ms.date: 07/11/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 046a3f1f16d8b48cddbd64bddf80a31ed1e50583
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/11/2019
ms.locfileid: "35633988"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="a2139-103">Conjuntos de requisitos de coerção de imagem</span><span class="sxs-lookup"><span data-stu-id="a2139-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="a2139-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="a2139-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="a2139-107">Os suplementos do Office executam várias versões do Office.</span><span class="sxs-lookup"><span data-stu-id="a2139-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="a2139-108">A tabela a seguir lista os conjuntos de requisitos de coerção de imagem, os aplicativos host do Office que dão suporte a esse conjunto de requisitos e os números de compilação ou versão para o aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="a2139-108">The following table lists the Image Coercion requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="a2139-109">ImageCoercion 1,1</span><span class="sxs-lookup"><span data-stu-id="a2139-109">ImageCoercion 1.1</span></span>

<span data-ttu-id="a2139-110">ImageCoercion 1,1 permite a conversão para uma imagem`Office.CoercionType.Image`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="a2139-110">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="a2139-111">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="a2139-111">The following hosts are supported:</span></span>

- <span data-ttu-id="a2139-112">Excel 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="a2139-112">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="a2139-113">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="a2139-113">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="a2139-114">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-114">Excel on the web</span></span>
- <span data-ttu-id="a2139-115">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="a2139-115">Excel on iPad</span></span>
- <span data-ttu-id="a2139-116">OneNote na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-116">OneNote on the web</span></span>
- <span data-ttu-id="a2139-117">PowerPoint 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="a2139-117">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="a2139-118">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="a2139-118">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="a2139-119">PowerPoint na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-119">PowerPoint on the web</span></span>
- <span data-ttu-id="a2139-120">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="a2139-120">PowerPoint on iPad</span></span>
- <span data-ttu-id="a2139-121">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="a2139-121">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="a2139-122">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="a2139-122">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="a2139-123">Word na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-123">Word on the web</span></span>
- <span data-ttu-id="a2139-124">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="a2139-124">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="a2139-125">ImageCoercion 1,2</span><span class="sxs-lookup"><span data-stu-id="a2139-125">ImageCoercion 1.2</span></span>

<span data-ttu-id="a2139-126">ImageCoercion 1,2 permite conversão para o formato SVG`Office.CoercionType.XmlSvg`() ao gravar dados usando [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) o método.</span><span class="sxs-lookup"><span data-stu-id="a2139-126">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="a2139-127">Há suporte para os seguintes hosts:</span><span class="sxs-lookup"><span data-stu-id="a2139-127">The following hosts are supported:</span></span>

- <span data-ttu-id="a2139-128">Excel no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-128">Excel on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-129">Excel no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-129">Excel on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-130">Excel na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-130">Excel on the web</span></span>
- <span data-ttu-id="a2139-131">PowerPoint no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-131">PowerPoint on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-132">PowerPoint no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-132">PowerPoint on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-133">PowerPoint na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-133">PowerPoint on the web</span></span>
- <span data-ttu-id="a2139-134">Word no Windows (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-134">Word on Windows (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-135">Word no Mac (conectado a uma assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="a2139-135">Word on Mac (connected to an Office 365 subscription)</span></span>
- <span data-ttu-id="a2139-136">Word na Web</span><span class="sxs-lookup"><span data-stu-id="a2139-136">Word on the web</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a2139-137">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="a2139-137">Office Common API requirement sets</span></span>

<span data-ttu-id="a2139-138">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="a2139-138">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="a2139-139">Confira também</span><span class="sxs-lookup"><span data-stu-id="a2139-139">See also</span></span>

- [<span data-ttu-id="a2139-140">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="a2139-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a2139-141">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="a2139-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="a2139-142">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="a2139-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
