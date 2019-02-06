---
title: Conjuntos de requisitos da Dialog API
description: ''
ms.date: 10/09/2018
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 95528b0973ef479dca109b159a3d623f945c14e6
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742272"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="b3927-102">Conjuntos de requisitos da Dialog API</span><span class="sxs-lookup"><span data-stu-id="b3927-102">Dialog API requirement sets</span></span>

<span data-ttu-id="b3927-p101">Os conjuntos de requisitos são grupos nomeados de membros da API. Os suplementos do Office usam conjuntos de requisitos especificados no manifesto ou usam uma verificação de tempo de execução para determinar se um host do Office dá suporte para as APIs necessárias para um suplemento. Para saber mais, confira [Versões do Office e conjuntos de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b3927-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="b3927-p102">Os suplementos do Office executam várias versões do Office. A tabela a seguir lista os conjuntos de requisitos da Dialog API, ou seja, os aplicativos de host do Office que oferecem suporte a esse conjunto de requisitos, e os números de versão ou de build dos aplicativos do Office.</span><span class="sxs-lookup"><span data-stu-id="b3927-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="b3927-108">Conjunto de requisitos</span><span class="sxs-lookup"><span data-stu-id="b3927-108">Requirement set</span></span>  | <span data-ttu-id="b3927-109">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="b3927-109">Office 2013 for Windows</span></span> | <span data-ttu-id="b3927-110">Office 2016 para Windows (Instalações MSI)</span><span class="sxs-lookup"><span data-stu-id="b3927-110">Office 2016 for Windows (MSI Installs)</span></span>   | <span data-ttu-id="b3927-111">Office 365 para Windows (Instalações C2R)</span><span class="sxs-lookup"><span data-stu-id="b3927-111">Office 365 for Windows (C2R Installs)</span></span>   |  <span data-ttu-id="b3927-112">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="b3927-112">Office 365 for iPad</span></span>  |  <span data-ttu-id="b3927-113">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="b3927-113">Office 365 for Mac</span></span>  | <span data-ttu-id="b3927-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="b3927-114">Office Online</span></span>  |  <span data-ttu-id="b3927-115">Servidor do Office Online</span><span class="sxs-lookup"><span data-stu-id="b3927-115">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="b3927-116">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="b3927-116">DialogApi 1.1</span></span>  | <span data-ttu-id="b3927-117">Build 15.0.4855.1000 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-117">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="b3927-118">Build 16.0.4390.1000 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-118">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="b3927-119">Versão 1602 (build 6741.0000) ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-119">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="b3927-120">1.22 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-120">1.22 or later</span></span> | <span data-ttu-id="b3927-121">15.20 ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-121">15.20 or later</span></span>| <span data-ttu-id="b3927-122">Janeiro de 2017</span><span class="sxs-lookup"><span data-stu-id="b3927-122">January 2017</span></span> | <span data-ttu-id="b3927-123">Versão 1608 (build 7601.6800) ou posterior</span><span class="sxs-lookup"><span data-stu-id="b3927-123">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="b3927-124">Para saber mais sobre versões, números de build e sobre o Servidor do Office Online, confira:</span><span class="sxs-lookup"><span data-stu-id="b3927-124">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- <span data-ttu-id="b3927-125">
  [Números de versão e de build de lançamentos de canais de atualização para clientes do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="b3927-125">[Version and build numbers of update channel releases for Office 365 clients](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- [<span data-ttu-id="b3927-126">Qual versão do Office estou usando?</span><span class="sxs-lookup"><span data-stu-id="b3927-126">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- <span data-ttu-id="b3927-127">
  [Onde você pode encontrar o número de versão e de build de um aplicativo cliente do Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span><span class="sxs-lookup"><span data-stu-id="b3927-127">[Where you can find the version and build number for an Office 365 client application](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)</span></span>
- <span data-ttu-id="b3927-128">
  [Visão geral sobre o Servidor do Office Online](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)</span><span class="sxs-lookup"><span data-stu-id="b3927-128">[Office Online Server overview](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)</span></span>

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="b3927-129">Conjuntos de requisitos da API comum do Office</span><span class="sxs-lookup"><span data-stu-id="b3927-129">Office Common API requirement sets</span></span>

<span data-ttu-id="b3927-130">Para saber mais sobre conjuntos de requisitos comuns da API, confira [Conjuntos de requisitos comuns da API do Office](office-add-in-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="b3927-130">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="b3927-131">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="b3927-131">Dialog API 1.1</span></span> 

<span data-ttu-id="b3927-132">O Dialog API 1.1 é a primeira versão da API.</span><span class="sxs-lookup"><span data-stu-id="b3927-132">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="b3927-133">Para saber mais sobre a API, confira o tópico de referência [Dialog API](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="b3927-133">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="b3927-134">Confira também</span><span class="sxs-lookup"><span data-stu-id="b3927-134">See also</span></span>

- [<span data-ttu-id="b3927-135">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="b3927-135">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="b3927-136">Especificar requisitos da API e de hosts do Office</span><span class="sxs-lookup"><span data-stu-id="b3927-136">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="b3927-137">Manifesto XML dos Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b3927-137">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
