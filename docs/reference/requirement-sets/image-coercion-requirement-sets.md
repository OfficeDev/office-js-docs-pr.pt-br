---
title: Conjuntos de requisitos de Coerção de Imagens
description: Suporte para conjuntos de requisitos de Coerção de Imagem com Office de Excel, PowerPoint e Word.
ms.date: 02/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 29614718378fd51013360a2a922e11f89bca14b8
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350215"
---
# <a name="image-coercion-requirement-sets"></a><span data-ttu-id="25fce-103">Conjuntos de requisitos de Coerção de Imagens</span><span class="sxs-lookup"><span data-stu-id="25fce-103">Image Coercion requirement sets</span></span>

<span data-ttu-id="25fce-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os Suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um aplicativo do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="25fce-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="imagecoercion-11"></a><span data-ttu-id="25fce-107">ImageCoercion 1.1</span><span class="sxs-lookup"><span data-stu-id="25fce-107">ImageCoercion 1.1</span></span>

<span data-ttu-id="25fce-108">ImageCoercion 1.1 permite a conversão em uma imagem ( ) ao escrever `Office.CoercionType.Image` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="25fce-108">ImageCoercion 1.1 enables conversion to an image (`Office.CoercionType.Image`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="25fce-109">Os aplicativos a seguir são suportados.</span><span class="sxs-lookup"><span data-stu-id="25fce-109">The following applications are supported.</span></span>

- <span data-ttu-id="25fce-110">Excel 2013 e posterior em Windows</span><span class="sxs-lookup"><span data-stu-id="25fce-110">Excel 2013 and later on Windows</span></span>
- <span data-ttu-id="25fce-111">Excel 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="25fce-111">Excel 2016 and later on Mac</span></span>
- <span data-ttu-id="25fce-112">Excel no iPad</span><span class="sxs-lookup"><span data-stu-id="25fce-112">Excel on iPad</span></span>
- <span data-ttu-id="25fce-113">OneNote Online</span><span class="sxs-lookup"><span data-stu-id="25fce-113">OneNote on the web</span></span>
- <span data-ttu-id="25fce-114">PowerPoint 2013 e posterior em Windows</span><span class="sxs-lookup"><span data-stu-id="25fce-114">PowerPoint 2013 and later on Windows</span></span>
- <span data-ttu-id="25fce-115">PowerPoint 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="25fce-115">PowerPoint 2016 and later on Mac</span></span>
- <span data-ttu-id="25fce-116">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="25fce-116">PowerPoint on the web</span></span>
- <span data-ttu-id="25fce-117">PowerPoint no iPad</span><span class="sxs-lookup"><span data-stu-id="25fce-117">PowerPoint on iPad</span></span>
- <span data-ttu-id="25fce-118">Word 2013 e posterior no Windows</span><span class="sxs-lookup"><span data-stu-id="25fce-118">Word 2013 and later on Windows</span></span>
- <span data-ttu-id="25fce-119">Word 2016 e posterior no Mac</span><span class="sxs-lookup"><span data-stu-id="25fce-119">Word 2016 and later on Mac</span></span>
- <span data-ttu-id="25fce-120">Word Online</span><span class="sxs-lookup"><span data-stu-id="25fce-120">Word on the web</span></span>
- <span data-ttu-id="25fce-121">Word no iPad</span><span class="sxs-lookup"><span data-stu-id="25fce-121">Word on iPad</span></span>

## <a name="imagecoercion-12"></a><span data-ttu-id="25fce-122">ImageCoercion 1.2</span><span class="sxs-lookup"><span data-stu-id="25fce-122">ImageCoercion 1.2</span></span>

<span data-ttu-id="25fce-123">ImageCoercion 1.2 permite a conversão para o formato SVG ( ) ao escrever `Office.CoercionType.XmlSvg` dados usando o [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) método.</span><span class="sxs-lookup"><span data-stu-id="25fce-123">ImageCoercion 1.2 enables conversion to SVG format (`Office.CoercionType.XmlSvg`) when writing data using the [`Document.setSelectedDataAsync`](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method.</span></span> <span data-ttu-id="25fce-124">Os aplicativos a seguir são suportados.</span><span class="sxs-lookup"><span data-stu-id="25fce-124">The following applications are supported.</span></span>

- <span data-ttu-id="25fce-125">Excel no Windows (conectado a uma assinatura de Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="25fce-125">Excel on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="25fce-126">Excel no Mac (conectado a uma assinatura de Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="25fce-126">Excel on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="25fce-127">PowerPoint no Windows (conectado a uma assinatura Microsoft 365 assinatura)</span><span class="sxs-lookup"><span data-stu-id="25fce-127">PowerPoint on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="25fce-128">PowerPoint no Mac (conectado a uma assinatura de Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="25fce-128">PowerPoint on Mac (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="25fce-129">PowerPoint Online</span><span class="sxs-lookup"><span data-stu-id="25fce-129">PowerPoint on the web</span></span>
- <span data-ttu-id="25fce-130">Word no Windows (conectado a uma assinatura Microsoft 365 de assinatura)</span><span class="sxs-lookup"><span data-stu-id="25fce-130">Word on Windows (connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="25fce-131">Word no Mac (conectado a Microsoft 365 assinatura)</span><span class="sxs-lookup"><span data-stu-id="25fce-131">Word on Mac (connected to a Microsoft 365 subscription)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="25fce-132">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="25fce-132">Office Common API requirement sets</span></span>

<span data-ttu-id="25fce-133">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="25fce-133">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="25fce-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="25fce-134">See also</span></span>

- [<span data-ttu-id="25fce-135">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="25fce-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="25fce-136">Especificar requisitos da API e de aplicativos do Office</span><span class="sxs-lookup"><span data-stu-id="25fce-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="25fce-137">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="25fce-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
