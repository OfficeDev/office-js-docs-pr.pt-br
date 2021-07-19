---
title: Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 7b3bd770d74f29d1a0b27da5080284aa62146101
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455492"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="ac97f-103">Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ac97f-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="ac97f-p101">Para funcionar conforme o esperado, o Suplemento do Office pode depender de um aplicativo específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API. As tabelas a seguir contêm as plataformas disponíveis, pontos de extensão, conjuntos de requisitos de API e APIs comuns que são atualmente suportados para cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="ac97f-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span data-ttu-id="ac97f-106"><span>Excel</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-106"><span>Excel</span></a></span></span>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span data-ttu-id="ac97f-107"><span>OneNote</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-107"><span>OneNote</span></a></span></span>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span data-ttu-id="ac97f-108"><span>Outlook</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-108"><span>Outlook</span></a></span></span>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span data-ttu-id="ac97f-109"><span>PowerPoint</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-109"><span>PowerPoint</span></a></span></span>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span data-ttu-id="ac97f-110"><span>Project</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-110"><span>Project</span></a></span></span>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span data-ttu-id="ac97f-111"><span>Word</span></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-111"><span>Word</span></a></span></span>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> <span data-ttu-id="ac97f-112">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="ac97f-112">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ac97f-113">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="ac97f-113">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span> <span data-ttu-id="ac97f-114">Os Suplementos do Office podem não ter suporte em todos os serviços que são membros do [Programa de Parceiros de Armazenamento em Nuvem do Office](https://developer.microsoft.com/office/cloud-storage-partner-program), que permite a integração do Office na Web para trabalhar com documentos do Office como parte de sua oferta de serviço.</span><span class="sxs-lookup"><span data-stu-id="ac97f-114">Office Add-ins may not be supported on all services that are members of the [Office Cloud Storage Partner Program](https://developer.microsoft.com/office/cloud-storage-partner-program), which enables integrating Office on the web to work with Office documents as part of their service offering.</span></span> <span data-ttu-id="ac97f-115">Para obter mais informações, entre em contato com o serviço de membro.</span><span class="sxs-lookup"><span data-stu-id="ac97f-115">For more information, please contact the member service.</span></span>

## <a name="excel"></a><span data-ttu-id="ac97f-116">Excel</span><span class="sxs-lookup"><span data-stu-id="ac97f-116">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ac97f-117">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-117">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ac97f-118">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-118">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ac97f-119">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-119">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ac97f-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-120"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-121">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-121">Office on the web</span></span></td>
    <td><span data-ttu-id="ac97f-122">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-122">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-123">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-123">
      - Content</span></span><br><span data-ttu-id="ac97f-124">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-124">
      - CustomFunctions</span></span><br><span data-ttu-id="ac97f-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-125">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-126">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-127">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-128">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-129">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-130">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-131">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-132">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-133">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-134">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac97f-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-135">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac97f-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-136">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ac97f-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-137">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="ac97f-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-138">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="ac97f-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-139">
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a></span></span><br><span data-ttu-id="ac97f-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-140">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-141">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-142">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="ac97f-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-143">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-144">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-145">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-146">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-147">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-148">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-149">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-150">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-151">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-152">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-153">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-154">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-155">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-156">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-156">Office on Windows</span></span><br><span data-ttu-id="ac97f-157">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-157">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-158">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-158">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-159">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-159">
      - Content</span></span><br><span data-ttu-id="ac97f-160">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-160">
      - CustomFunctions</span></span><br><span data-ttu-id="ac97f-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-161">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-162">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-163">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-164">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-165">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-166">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-167">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-168">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-169">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-170">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac97f-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-171">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac97f-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-172">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ac97f-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-173">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="ac97f-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-174">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="ac97f-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-175">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-176">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-177">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-178">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-179">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="ac97f-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-180">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="ac97f-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-181">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-182">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-183">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-184">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-185">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-186">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-187">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-188">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-189">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-190">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-191">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-192">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-193">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-194">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-194">Office 2019 on Windows</span></span><br><span data-ttu-id="ac97f-195">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-195">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-196">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-196">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-197">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-197">
      - Content</span></span><br><span data-ttu-id="ac97f-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-198">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-199">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-200">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-201">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-202">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-203">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-204">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-205">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-206">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-207">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-208">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-209">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-210">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-211">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-212">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-213">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-214">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-215">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-216">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-217">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-218">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-219">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-220">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-221">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-221">Office 2016 on Windows</span></span><br><span data-ttu-id="ac97f-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-223">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-223">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-224">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-224">
      - Content</span></span> </td>
    <td><span data-ttu-id="ac97f-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-225">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-226">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-227">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-228">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-229">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-230">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-231">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-232">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-233">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-234">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-235">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-236">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-237">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-238">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-239">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-240">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-240">Office 2013 on Windows</span></span><br><span data-ttu-id="ac97f-241">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-241">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-242">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-242">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-243">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-243">
      - Content</span></span> </td>
    <td><span data-ttu-id="ac97f-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-244">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-245">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-246">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-247">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-248">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-249">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-250">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-251">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-252">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-253">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-254">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-255">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-256">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-257">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac97f-257">Office on iPad</span></span><br><span data-ttu-id="ac97f-258">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-258">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-259">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-259">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-260">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-260">
      - Content</span></span> </td>
    <td><span data-ttu-id="ac97f-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-261">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-262">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-263">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-264">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-265">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-266">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-267">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-268">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-269">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac97f-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-270">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac97f-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-271">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ac97f-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-272">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="ac97f-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-273">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="ac97f-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-274">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-275">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-276">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-277">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-278">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-279">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-280">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-281">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-282">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-283">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-284">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-285">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-286">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-287">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-288">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-288">Office on Mac</span></span><br><span data-ttu-id="ac97f-289">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-289">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-290">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-290">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-291">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-291">
      - Content</span></span><br><span data-ttu-id="ac97f-292">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-292">
      - CustomFunctions</span></span><br><span data-ttu-id="ac97f-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-293">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-294">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-295">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-296">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-297">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-298">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-299">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-300">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-301">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-302">
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac97f-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-303">
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac97f-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-304">
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ac97f-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-305">
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a></span></span><br><span data-ttu-id="ac97f-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-306">
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a></span></span><br><span data-ttu-id="ac97f-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-307">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-308">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-309">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-310">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-311">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a></span></span><br><span data-ttu-id="ac97f-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-312">
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a></span></span><br><span data-ttu-id="ac97f-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-313">
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-314">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-315">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-316">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-317">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-318">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-319">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-320">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-321">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-322">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-323">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-324">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-325">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-326">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-327">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-327">Office 2019 on Mac</span></span><br><span data-ttu-id="ac97f-328">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-329">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-329">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-330">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-330">
      - Content</span></span><br><span data-ttu-id="ac97f-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-331">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-332">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-333">
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac97f-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-334">
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac97f-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-335">
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac97f-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-336">
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac97f-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-337">
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac97f-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-338">
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac97f-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-339">
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac97f-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-340">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-341">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-342">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-343">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-344">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-345">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-346">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-347">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-348">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-349">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-350">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-351">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-352">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-353">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-354">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-355">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-355">Office 2016 on Mac</span></span><br><span data-ttu-id="ac97f-356">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-356">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-357">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-357">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-358">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-358">
      - Content</span></span> </td>
    <td><span data-ttu-id="ac97f-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-359">
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac97f-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-360">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-361">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-362">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-363">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-364">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-365">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-366">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-367">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-368">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-369">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-370">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-371">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-372">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-373">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-374">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="ac97f-375">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac97f-375">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ac97f-376">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="ac97f-376">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-377">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-377">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-378">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-378">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-379">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-379">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-380"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-381">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-381">Office on the web</span></span></td>
    <td><span data-ttu-id="ac97f-382">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-382">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="ac97f-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-383">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="ac97f-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-384">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="ac97f-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-385">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-386">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-386">Office on Windows</span></span><br><span data-ttu-id="ac97f-387">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-387">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-388">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-388">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="ac97f-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-389">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="ac97f-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-390">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="ac97f-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-391">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-392">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-392">Office on Mac</span></span><br><span data-ttu-id="ac97f-393">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-393">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-394">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="ac97f-394">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="ac97f-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-395">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="ac97f-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-396">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="ac97f-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-397">
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ac97f-398">Outlook</span><span class="sxs-lookup"><span data-stu-id="ac97f-398">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-399">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-399">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-400">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-400">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-401">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-401">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-402"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-403">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-403">Office on the web</span></span><br><span data-ttu-id="ac97f-404">(moderno)</span><span class="sxs-lookup"><span data-stu-id="ac97f-404">(modern)</span></span></td>
    <td><span data-ttu-id="ac97f-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-405">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-406">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-407">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-408">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-409">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-410">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-411">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-412">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-413">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-414">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-415">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac97f-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-416">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac97f-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-417">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="ac97f-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Caixa de correio 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-418">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="ac97f-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Caixa de correio 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-419">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="ac97f-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-420">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="ac97f-421">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-421">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-422">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-422">Office on the web</span></span><br><span data-ttu-id="ac97f-423">(clássico)</span><span class="sxs-lookup"><span data-stu-id="ac97f-423">(classic)</span></span></td>
    <td><span data-ttu-id="ac97f-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-424">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-425">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-426">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-427">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-428">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-429">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-430">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-431">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-432">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-433">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-434">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-435">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-436">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-436">Office on Windows</span></span><br><span data-ttu-id="ac97f-437">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-437">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-438">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-439">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-440">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-441">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-442">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="ac97f-443">
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-443">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-444">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-445">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-446">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-447">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-448">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-449">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac97f-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-450">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac97f-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-451">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="ac97f-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Caixa de correio 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-452">
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Mailbox 1.9</a></span></span><br><span data-ttu-id="ac97f-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Caixa de correio 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-453">
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Mailbox 1.10</a></span></span><br><span data-ttu-id="ac97f-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="ac97f-454">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="ac97f-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-455">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-457">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-457">Office 2019 on Windows</span></span><br><span data-ttu-id="ac97f-458">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-458">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-459">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-460">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-461">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-462">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-463">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="ac97f-464">
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-464">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-465">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-466">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-467">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-468">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-469">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-470">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac97f-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-471">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-472">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-472">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-473">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-473">Office 2016 on Windows</span></span><br><span data-ttu-id="ac97f-474">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-474">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-475">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-476">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-477">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-478">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-479">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a></span></span><br><span data-ttu-id="ac97f-480">
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-480">
      - <a href="../reference/manifest/extensionpoint.md#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-481">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-482">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-483">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Caixa de correio 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-484">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="ac97f-485">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-485">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-486">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-486">Office 2013 on Windows</span></span><br><span data-ttu-id="ac97f-487">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-487">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-488">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-489">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-490">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-491">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="ac97f-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-492">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-493">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Caixa de correio 1.3</a><sup>2</sup></span><span class="sxs-lookup"><span data-stu-id="ac97f-494">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a><sup>2</sup></span></span><br><span data-ttu-id="ac97f-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Caixa de correio 1.4</a><sup>2</sup>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-495">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a><sup>2</sup>
    </span></span></td>
    <td><span data-ttu-id="ac97f-496">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-497">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="ac97f-497">Office on iOS</span></span><br><span data-ttu-id="ac97f-498">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-498">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-499">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organizador de compromissos (Redigir): reunião online</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-500">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="ac97f-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-501">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-502">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-503">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-504">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-505">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-506">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-507">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-507">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-508">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-508">Office on Mac</span></span><br><span data-ttu-id="ac97f-509">(Interface do Usuário atual,</span><span class="sxs-lookup"><span data-stu-id="ac97f-509">(current UI,</span></span><br><span data-ttu-id="ac97f-510">conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-510">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-511">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-512">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-513">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-514">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-515">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-516">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-517">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-518">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-519">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-520">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-521">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac97f-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-522">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac97f-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-523">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="ac97f-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span><span class="sxs-lookup"><span data-stu-id="ac97f-524">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup></span></span><br><span data-ttu-id="ac97f-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-525">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-526">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-526">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-527">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-527">Office on Mac</span></span><br><span data-ttu-id="ac97f-528">(nova Interface do Usuário (visualização)<sup>3</sup>,</span><span class="sxs-lookup"><span data-stu-id="ac97f-528">(new UI (preview)<sup>3</sup>,</span></span><br><span data-ttu-id="ac97f-529">conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-529">connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-530">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-531">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-532">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-533">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-534">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-535">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-536">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-537">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-538">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-539">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-540">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac97f-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-541">
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac97f-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-542">
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Mailbox 1.8</a></span></span><br><span data-ttu-id="ac97f-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-543">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </span></span></td>
    <td><span data-ttu-id="ac97f-544">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-544">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-545">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-545">Office 2019 on Mac</span></span><br><span data-ttu-id="ac97f-546">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-546">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-547">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-548">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-549">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-550">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-551">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-552">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-553">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-554">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-555">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-556">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-557">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-558">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-558">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-559">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-559">Office 2016 on Mac</span></span><br><span data-ttu-id="ac97f-560">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-560">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-561">
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-562">
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ac97f-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-563">
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ac97f-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-564">
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ac97f-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-565">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-566">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-567">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-568">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-569">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-570">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac97f-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-571">
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-572">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-572">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-573">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="ac97f-573">Office on Android</span></span><br><span data-ttu-id="ac97f-574">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-574">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-575">
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ac97f-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organizador de compromissos (Redigir): reunião online</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-576">
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Appointment Organizer (Compose): online meeting</a></span></span><br><span data-ttu-id="ac97f-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-577">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-578">
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac97f-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-579">
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac97f-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-580">
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac97f-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-581">
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac97f-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Caixa de correio 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-582">
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-583">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac97f-583">Not available</span></span></td>
  </tr>
</table>

> [!NOTE]
> <span data-ttu-id="ac97f-584"><sup>1</sup> Para exigir o conjunto 1.3 da API de Identidade no código do suplemento, verifique se ele tem suporte ligando para `isSetSupported('IdentityAPI', '1.3')`.</span><span class="sxs-lookup"><span data-stu-id="ac97f-584"><sup>1</sup> To require Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="ac97f-585">Não há suporte para declará-lo no manifesto do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="ac97f-585">Declaring it in your add-in's manifest isn't supported.</span></span> <span data-ttu-id="ac97f-586">Você também pode determinar se a API tem suporte, verificando se ela não é `undefined`.</span><span class="sxs-lookup"><span data-stu-id="ac97f-586">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="ac97f-587">Para mais detalhes, confira [Usar APIs de conjuntos de requisitos posteriores](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="ac97f-587">For further details, see [Using APIs from later requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>
>
> <span data-ttu-id="ac97f-588"><sup>2</sup> Adicionado com atualizações pós-lançamento.</span><span class="sxs-lookup"><span data-stu-id="ac97f-588"><sup>2</sup> Added with post-release updates.</span></span>
>
> <span data-ttu-id="ac97f-589"><sup>3</sup> O suporte para a nova Interface do Usuário do Mac (visualização) está disponível no Outlook versão 16.38.506.</span><span class="sxs-lookup"><span data-stu-id="ac97f-589"><sup>3</sup> Support for the new Mac UI (preview) is available from Outlook version 16.38.506.</span></span> <span data-ttu-id="ac97f-590">Para mais informações, consulte a seção [Suporte de Suplemento no Outlook na nova Interface do Usuário do Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).</span><span class="sxs-lookup"><span data-stu-id="ac97f-590">For more information, see the [Add-in support in Outlook on new Mac UI](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview) section.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac97f-591">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="ac97f-591">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ac97f-592">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="ac97f-592">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ac97f-593">Word</span><span class="sxs-lookup"><span data-stu-id="ac97f-593">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-594">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-594">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-595">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-595">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-596">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-596">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-597"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-598">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-598">Office on the web</span></span></td>
    <td><span data-ttu-id="ac97f-599">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-599">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-600">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-601">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-602">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-603">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-604">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-605">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-606">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-607">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-608">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-609">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-610">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-611">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-612">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-613">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-614">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-615">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-616">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-617">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-618">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-619">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-620">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-621">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-622">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-623">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-623">Office on Windows</span></span><br><span data-ttu-id="ac97f-624">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-624">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-625">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-625">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-626">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-627">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-628">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-629">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-630">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-631">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-632">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-633">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-634">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-635">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-636">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-637">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-638">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-639">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-640">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-641">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-642">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-643">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-644">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-645">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-646">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-647">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-648">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-649">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-650">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-651">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-652">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-652">Office 2019 on Windows</span></span><br><span data-ttu-id="ac97f-653">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-653">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-654">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-654">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-655">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-656">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-657">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-658">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-659">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-660">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-661">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-662">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-663">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-664">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-665">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-666">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-667">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-668">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-669">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-670">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-671">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-672">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-673">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-674">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-675">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-676">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-677">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-678">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-678">Office 2016 on Windows</span></span><br><span data-ttu-id="ac97f-679">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-679">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-680">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-680">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-681">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-682">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-683">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-684">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-685">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-686">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-687">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-688">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-689">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-690">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-691">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-692">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-693">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-694">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-695">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-696">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-697">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-698">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-699">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-700">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-701">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-701">Office 2013 on Windows</span></span><br><span data-ttu-id="ac97f-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-702">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-703">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-704">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-705">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-706">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-707">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-708">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-709">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-710">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-711">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-712">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-713">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-714">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-715">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-716">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-717">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-718">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-719">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-720">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-721">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-722">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-723">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac97f-723">Office on iPad</span></span><br><span data-ttu-id="ac97f-724">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-724">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-725">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-725">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-726">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-727">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-728">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-729">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-730">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-731">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-732">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-733">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-734">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-735">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-736">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-737">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-738">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-739">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-740">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-741">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-742">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-743">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-744">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-745">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-746">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-747">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-748">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-748">Office on Mac</span></span><br><span data-ttu-id="ac97f-749">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-749">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-750">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-750">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-751">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-752">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-753">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-754">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-755">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-756">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-757">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-758">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-759">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-760">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-761">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-762">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-763">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-764">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-765">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-766">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-767">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-768">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-769">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-770">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-771">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-772">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-773">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-774">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-775">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-776">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-777">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-777">Office 2019 on Mac</span></span><br><span data-ttu-id="ac97f-778">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-778">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-779">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-779">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-780">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-781">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-782">
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a></span></span><br><span data-ttu-id="ac97f-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-783">
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a></span></span><br><span data-ttu-id="ac97f-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-784">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-785">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-786">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-787">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-788">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-789">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-790">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-791">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-792">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-793">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-794">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-795">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-796">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-797">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-798">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-799">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-800">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-801">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-802">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-803">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-803">Office 2016 on Mac</span></span><br><span data-ttu-id="ac97f-804">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-804">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-805">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-805">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-806">
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a></span></span><br><span data-ttu-id="ac97f-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-807">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-808">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-809">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="ac97f-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-810">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-811">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="ac97f-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-812">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-813">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-814">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-815">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="ac97f-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-816">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="ac97f-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-817">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="ac97f-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-818">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-819">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-820">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-821">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="ac97f-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-822">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="ac97f-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-823">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a></span></span><br><span data-ttu-id="ac97f-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-824">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="ac97f-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-825">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="ac97f-826">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac97f-826">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ac97f-827">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ac97f-827">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-828">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-828">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-829">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-829">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-830">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-830">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-831"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-832">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-832">Office on the web</span></span></td>
    <td><span data-ttu-id="ac97f-833">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-833">
      - Content</span></span><br><span data-ttu-id="ac97f-834">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-834">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-835">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-836">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac97f-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-837">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="ac97f-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-838">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-839">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-840">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-841">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-842">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-843">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-844">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-845">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-846">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-847">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-848">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-849">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-850">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-850">Office on Windows</span></span><br><span data-ttu-id="ac97f-851">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-851">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-852">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-852">
      - Content</span></span><br><span data-ttu-id="ac97f-853">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-853">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-854">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-855">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac97f-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-856">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="ac97f-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-857">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-858">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-859">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-860">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-861">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-862">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-863">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-864">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-865">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-866">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-867">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-868">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-869">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-870">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-870">Office 2019 on Windows</span></span><br><span data-ttu-id="ac97f-871">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-871">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-872">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-872">
      - Content</span></span><br><span data-ttu-id="ac97f-873">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-873">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-874">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-875">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-876">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-877">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-878">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-879">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-880">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-881">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-882">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-883">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-884">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-885">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-885">Office 2016 on Windows</span></span><br><span data-ttu-id="ac97f-886">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-886">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-887">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-887">
      - Content</span></span><br><span data-ttu-id="ac97f-888">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-888">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="ac97f-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-889">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-890">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-891">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-892">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-893">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-894">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-895">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-896">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-897">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-898">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-899">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-899">Office 2013 on Windows</span></span><br><span data-ttu-id="ac97f-900">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-900">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-901">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-901">
      - Content</span></span><br><span data-ttu-id="ac97f-902">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-902">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="ac97f-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-903">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-904">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-905">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-906">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-907">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-908">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-909">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-910">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-911">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-912">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-913">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac97f-913">Office on iPad</span></span><br><span data-ttu-id="ac97f-914">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-914">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-915">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-915">
      - Content</span></span><br><span data-ttu-id="ac97f-916">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-916">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="ac97f-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-917">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac97f-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-918">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="ac97f-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-919">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-920">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-921">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-922">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-923">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-924">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-925">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-926">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-927">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-928">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-929">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-930">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-930">Office on Mac</span></span><br><span data-ttu-id="ac97f-931">(conectado a uma assinatura do Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="ac97f-931">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac97f-932">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-932">
      - Content</span></span><br><span data-ttu-id="ac97f-933">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-933">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-934">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-935">
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac97f-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-936">
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a></span></span><br><span data-ttu-id="ac97f-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-937">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-938">
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a></span></span><br><span data-ttu-id="ac97f-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-939">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac97f-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-940">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a></span></span><br><span data-ttu-id="ac97f-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-941">
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-942">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-943">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-944">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-945">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-946">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-947">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-948">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-949">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-950">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-950">Office 2019 on Mac</span></span><br><span data-ttu-id="ac97f-951">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-951">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-952">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-952">
      - Content</span></span><br><span data-ttu-id="ac97f-953">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-953">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-954">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-955">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-956">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-957">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-958">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-959">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-960">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-961">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-962">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-963">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-964">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-965">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-965">Office 2016 on Mac</span></span><br><span data-ttu-id="ac97f-966">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-966">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-967">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-967">
      - Content</span></span><br><span data-ttu-id="ac97f-968">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-968">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="ac97f-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac97f-969">
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac97f-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-970">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-971">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a></span></span><br><span data-ttu-id="ac97f-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-972">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="ac97f-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-973">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-974">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">File</a></span></span><br><span data-ttu-id="ac97f-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-975">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a></span></span><br><span data-ttu-id="ac97f-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-976">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-977">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-978">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="ac97f-979">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac97f-979">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ac97f-980">OneNote</span><span class="sxs-lookup"><span data-stu-id="ac97f-980">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-981">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-981">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-982">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-982">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-983">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-983">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-984"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-985">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac97f-985">Office on the web</span></span></td>
    <td><span data-ttu-id="ac97f-986">
      - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac97f-986">
      - Content</span></span><br><span data-ttu-id="ac97f-987">
      - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-987">
      - TaskPane</span></span><br><span data-ttu-id="ac97f-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-988">
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-989">
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ac97f-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-990">
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac97f-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-991">
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="ac97f-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-992">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="ac97f-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-993">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="ac97f-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-994">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Settings</a></span></span><br><span data-ttu-id="ac97f-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-995">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ac97f-996">Project</span><span class="sxs-lookup"><span data-stu-id="ac97f-996">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac97f-997">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac97f-997">Platform</span></span></th>
    <th><span data-ttu-id="ac97f-998">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac97f-998">Extension points</span></span></th>
    <th><span data-ttu-id="ac97f-999">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-999">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac97f-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1000"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-1001">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-1001">Office 2019 on Windows</span></span><br><span data-ttu-id="ac97f-1002">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1002">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-1003">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-1003">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1004">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ac97f-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1005">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-1006">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-1007">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-1007">Office 2016 on Windows</span></span><br><span data-ttu-id="ac97f-1008">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1008">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-1009">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-1009">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1010">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ac97f-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1011">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-1012">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac97f-1013">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac97f-1013">Office 2013 on Windows</span></span><br><span data-ttu-id="ac97f-1014">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1014">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac97f-1015">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac97f-1015">- TaskPane</span></span></td>
    <td><span data-ttu-id="ac97f-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1016">- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ac97f-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="ac97f-1017">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a></span></span><br><span data-ttu-id="ac97f-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac97f-1018">
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ac97f-1019">Confira também</span><span class="sxs-lookup"><span data-stu-id="ac97f-1019">See also</span></span>

- [<span data-ttu-id="ac97f-1020">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ac97f-1020">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ac97f-1021">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="ac97f-1021">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ac97f-1022">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-1022">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ac97f-1023">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="ac97f-1023">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ac97f-1024">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="ac97f-1024">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ac97f-1025">Histórico de atualizações para Microsoft 365 Apps</span><span class="sxs-lookup"><span data-stu-id="ac97f-1025">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ac97f-1026">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1026">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ac97f-1027">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1027">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ac97f-1028">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1028">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ac97f-1029">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac97f-1029">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ac97f-1030">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="ac97f-1030">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ac97f-1031">Desenvolver Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="ac97f-1031">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
