---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611979"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c2fd6-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c2fd6-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c2fd6-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="c2fd6-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="c2fd6-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="c2fd6-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c2fd6-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="c2fd6-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c2fd6-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="c2fd6-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c2fd6-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c2fd6-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c2fd6-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c2fd6-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c2fd6-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c2fd6-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c2fd6-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-114">- TaskPane</span></span><br><span data-ttu-id="c2fd6-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-115">
        - Content</span></span><br><span data-ttu-id="c2fd6-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-116">
        - Custom Functions</span></span><br><span data-ttu-id="c2fd6-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="c2fd6-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c2fd6-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c2fd6-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c2fd6-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c2fd6-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c2fd6-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-131">
        - BindingEvents</span></span><br><span data-ttu-id="c2fd6-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-132">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-133">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-134">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-134">
        - File</span></span><br><span data-ttu-id="c2fd6-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-135">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-137">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-137">
        - Selection</span></span><br><span data-ttu-id="c2fd6-138">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-138">
        - Settings</span></span><br><span data-ttu-id="c2fd6-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-139">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-140">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-141">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-143">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-143">Office on Windows</span></span><br><span data-ttu-id="c2fd6-144">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-145">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-145">- TaskPane</span></span><br><span data-ttu-id="c2fd6-146">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-146">
        - Content</span></span><br><span data-ttu-id="c2fd6-147">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-147">
        - Custom Functions</span></span><br><span data-ttu-id="c2fd6-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="c2fd6-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c2fd6-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c2fd6-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c2fd6-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c2fd6-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c2fd6-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-163">
        - BindingEvents</span></span><br><span data-ttu-id="c2fd6-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-164">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-165">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-166">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-166">
        - File</span></span><br><span data-ttu-id="c2fd6-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-167">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-169">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-169">
        - Selection</span></span><br><span data-ttu-id="c2fd6-170">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-170">
        - Settings</span></span><br><span data-ttu-id="c2fd6-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-171">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-172">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-173">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-175">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-175">Office 2019 on Windows</span></span><br><span data-ttu-id="c2fd6-176">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c2fd6-177">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-177">- TaskPane</span></span><br><span data-ttu-id="c2fd6-178">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-178">
        - Content</span></span><br><span data-ttu-id="c2fd6-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c2fd6-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-190">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-191">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-192">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-193">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-193">
        - File</span></span><br><span data-ttu-id="c2fd6-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-194">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-196">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-196">
        - Selection</span></span><br><span data-ttu-id="c2fd6-197">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-197">
        - Settings</span></span><br><span data-ttu-id="c2fd6-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-198">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-199">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-200">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-202">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-202">Office 2016 on Windows</span></span><br><span data-ttu-id="c2fd6-203">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c2fd6-204">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-204">- TaskPane</span></span><br><span data-ttu-id="c2fd6-205">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-205">
        - Content</span></span></td>
    <td><span data-ttu-id="c2fd6-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c2fd6-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-209">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-210">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-211">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-212">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-212">
        - File</span></span><br><span data-ttu-id="c2fd6-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-213">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-215">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-215">
        - Selection</span></span><br><span data-ttu-id="c2fd6-216">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-216">
        - Settings</span></span><br><span data-ttu-id="c2fd6-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-217">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-218">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-219">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-221">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-221">Office 2013 on Windows</span></span><br><span data-ttu-id="c2fd6-222">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c2fd6-223">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-223">
        - TaskPane</span></span><br><span data-ttu-id="c2fd6-224">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c2fd6-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c2fd6-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-227">
        - BindingEvents</span></span><br><span data-ttu-id="c2fd6-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-228">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-229">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-230">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-230">
        - File</span></span><br><span data-ttu-id="c2fd6-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-231">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-233">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-233">
        - Selection</span></span><br><span data-ttu-id="c2fd6-234">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-234">
        - Settings</span></span><br><span data-ttu-id="c2fd6-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-235">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-236">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-237">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-239">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c2fd6-239">Office on iPad</span></span><br><span data-ttu-id="c2fd6-240">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c2fd6-241">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-241">- TaskPane</span></span><br><span data-ttu-id="c2fd6-242">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-242">
        - Content</span></span></td>
    <td><span data-ttu-id="c2fd6-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c2fd6-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c2fd6-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c2fd6-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-256">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-257">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-258">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-258">
        - File</span></span><br><span data-ttu-id="c2fd6-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-259">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-261">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-261">
        - Selection</span></span><br><span data-ttu-id="c2fd6-262">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-262">
        - Settings</span></span><br><span data-ttu-id="c2fd6-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-263">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-264">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-265">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-267">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-267">Office on Mac</span></span><br><span data-ttu-id="c2fd6-268">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c2fd6-269">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-269">- TaskPane</span></span><br><span data-ttu-id="c2fd6-270">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-270">
        - Content</span></span><br><span data-ttu-id="c2fd6-271">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-271">
        - Custom Functions</span></span><br><span data-ttu-id="c2fd6-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c2fd6-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c2fd6-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c2fd6-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c2fd6-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c2fd6-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-287">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-288">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-289">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-290">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-290">
        - File</span></span><br><span data-ttu-id="c2fd6-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-291">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-293">
        - PdfFile</span></span><br><span data-ttu-id="c2fd6-294">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-294">
        - Selection</span></span><br><span data-ttu-id="c2fd6-295">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-295">
        - Settings</span></span><br><span data-ttu-id="c2fd6-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-296">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-297">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-298">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-300">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-300">Office 2019 on Mac</span></span><br><span data-ttu-id="c2fd6-301">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c2fd6-302">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-302">- TaskPane</span></span><br><span data-ttu-id="c2fd6-303">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-303">
        - Content</span></span><br><span data-ttu-id="c2fd6-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c2fd6-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c2fd6-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c2fd6-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c2fd6-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c2fd6-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c2fd6-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-315">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-316">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-317">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-318">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-318">
        - File</span></span><br><span data-ttu-id="c2fd6-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-319">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-321">
        - PdfFile</span></span><br><span data-ttu-id="c2fd6-322">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-322">
        - Selection</span></span><br><span data-ttu-id="c2fd6-323">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-323">
        - Settings</span></span><br><span data-ttu-id="c2fd6-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-324">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-325">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-326">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-328">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-328">Office 2016 on Mac</span></span><br><span data-ttu-id="c2fd6-329">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c2fd6-330">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-330">- TaskPane</span></span><br><span data-ttu-id="c2fd6-331">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-331">
        - Content</span></span></td>
    <td><span data-ttu-id="c2fd6-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c2fd6-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c2fd6-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-335">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-336">
        - CompressedFile</span></span><br><span data-ttu-id="c2fd6-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-337">
        - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-338">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-338">
        - File</span></span><br><span data-ttu-id="c2fd6-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-339">
        - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-341">
        - PdfFile</span></span><br><span data-ttu-id="c2fd6-342">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-342">
        - Selection</span></span><br><span data-ttu-id="c2fd6-343">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-343">
        - Settings</span></span><br><span data-ttu-id="c2fd6-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-344">
        - TableBindings</span></span><br><span data-ttu-id="c2fd6-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-345">
        - TableCoercion</span></span><br><span data-ttu-id="c2fd6-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-346">
        - TextBindings</span></span><br><span data-ttu-id="c2fd6-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c2fd6-348">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c2fd6-349">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c2fd6-350">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c2fd6-351">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c2fd6-352">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c2fd6-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-354">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-354">Office on the web</span></span></td>
    <td><span data-ttu-id="c2fd6-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c2fd6-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-357">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-357">Office on Windows</span></span><br><span data-ttu-id="c2fd6-358">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c2fd6-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c2fd6-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-361">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-361">Office for Mac</span></span><br><span data-ttu-id="c2fd6-362">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c2fd6-363">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="c2fd6-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c2fd6-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c2fd6-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="c2fd6-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c2fd6-366">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-366">Platform</span></span></th>
    <th><span data-ttu-id="c2fd6-367">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-367">Extension points</span></span></th>
    <th><span data-ttu-id="c2fd6-368">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="c2fd6-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-370">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-370">Office on the web</span></span><br><span data-ttu-id="c2fd6-371">(moderno)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-371">(modern)</span></span></td>
    <td> <span data-ttu-id="c2fd6-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c2fd6-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c2fd6-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c2fd6-385">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-386">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-386">Office on the web</span></span><br><span data-ttu-id="c2fd6-387">(clássico)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-387">(classic)</span></span></td>
    <td> <span data-ttu-id="c2fd6-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c2fd6-399">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-400">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-400">Office on Windows</span></span><br><span data-ttu-id="c2fd6-401">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c2fd6-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c2fd6-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c2fd6-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c2fd6-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-417">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-417">Office 2019 on Windows</span></span><br><span data-ttu-id="c2fd6-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c2fd6-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c2fd6-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c2fd6-432">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-433">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-433">Office 2016 on Windows</span></span><br><span data-ttu-id="c2fd6-434">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c2fd6-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c2fd6-445">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-446">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-446">Office 2013 on Windows</span></span><br><span data-ttu-id="c2fd6-447">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="c2fd6-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c2fd6-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c2fd6-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-457">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="c2fd6-457">Office on iOS</span></span><br><span data-ttu-id="c2fd6-458">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c2fd6-466">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-467">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-467">Office on Mac</span></span><br><span data-ttu-id="c2fd6-468">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c2fd6-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c2fd6-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c2fd6-482">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-483">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-483">Office 2019 on Mac</span></span><br><span data-ttu-id="c2fd6-484">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c2fd6-496">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-497">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-497">Office 2016 on Mac</span></span><br><span data-ttu-id="c2fd6-498">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c2fd6-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c2fd6-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c2fd6-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c2fd6-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c2fd6-510">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-511">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="c2fd6-511">Office on Android</span></span><br><span data-ttu-id="c2fd6-512">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c2fd6-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="c2fd6-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c2fd6-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c2fd6-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c2fd6-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c2fd6-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c2fd6-521">Não disponível</span><span class="sxs-lookup"><span data-stu-id="c2fd6-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c2fd6-522">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c2fd6-523">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="c2fd6-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c2fd6-524">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="c2fd6-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c2fd6-525">Word</span><span class="sxs-lookup"><span data-stu-id="c2fd6-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c2fd6-526">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-526">Platform</span></span></th>
    <th><span data-ttu-id="c2fd6-527">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-527">Extension points</span></span></th>
    <th><span data-ttu-id="c2fd6-528">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="c2fd6-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-530">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="c2fd6-531">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-531">- TaskPane</span></span><br><span data-ttu-id="c2fd6-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-539">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-541">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-542">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-542">
         - File</span></span><br><span data-ttu-id="c2fd6-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-544">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-547">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-548">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-548">
         - Selection</span></span><br><span data-ttu-id="c2fd6-549">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-549">
         - Settings</span></span><br><span data-ttu-id="c2fd6-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-550">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-551">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-552">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-553">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-555">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-555">Office on Windows</span></span><br><span data-ttu-id="c2fd6-556">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-557">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-557">- TaskPane</span></span><br><span data-ttu-id="c2fd6-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-565">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-566">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-568">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-569">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-569">
         - File</span></span><br><span data-ttu-id="c2fd6-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-571">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-574">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-575">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-575">
         - Selection</span></span><br><span data-ttu-id="c2fd6-576">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-576">
         - Settings</span></span><br><span data-ttu-id="c2fd6-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-577">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-578">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-579">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-580">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-582">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-582">Office 2019 on Windows</span></span><br><span data-ttu-id="c2fd6-583">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-584">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-584">- TaskPane</span></span><br><span data-ttu-id="c2fd6-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-591">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-592">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-595">
         - File</span></span><br><span data-ttu-id="c2fd6-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-600">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-601">
         - Selection</span></span><br><span data-ttu-id="c2fd6-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-602">
         - Settings</span></span><br><span data-ttu-id="c2fd6-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-603">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-604">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-605">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-606">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-608">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-608">Office 2016 on Windows</span></span><br><span data-ttu-id="c2fd6-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c2fd6-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-614">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-615">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-617">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-618">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-618">
         - File</span></span><br><span data-ttu-id="c2fd6-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-620">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-623">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-624">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-624">
         - Selection</span></span><br><span data-ttu-id="c2fd6-625">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-625">
         - Settings</span></span><br><span data-ttu-id="c2fd6-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-626">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-627">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-628">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-629">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-631">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-631">Office 2013 on Windows</span></span><br><span data-ttu-id="c2fd6-632">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-633">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c2fd6-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-636">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-637">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-639">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-640">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-640">
         - File</span></span><br><span data-ttu-id="c2fd6-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-642">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-645">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-646">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-646">
         - Selection</span></span><br><span data-ttu-id="c2fd6-647">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-647">
         - Settings</span></span><br><span data-ttu-id="c2fd6-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-648">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-649">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-650">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-651">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-653">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c2fd6-653">Office on iPad</span></span><br><span data-ttu-id="c2fd6-654">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-655">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c2fd6-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-661">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-662">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-664">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-665">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-665">
         - File</span></span><br><span data-ttu-id="c2fd6-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-667">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-670">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-671">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-671">
         - Selection</span></span><br><span data-ttu-id="c2fd6-672">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-672">
         - Settings</span></span><br><span data-ttu-id="c2fd6-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-673">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-674">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-675">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-676">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-678">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-678">Office on Mac</span></span><br><span data-ttu-id="c2fd6-679">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-680">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-680">- TaskPane</span></span><br><span data-ttu-id="c2fd6-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c2fd6-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-688">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-689">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-691">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-692">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-692">
         - File</span></span><br><span data-ttu-id="c2fd6-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-694">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-697">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-698">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-698">
         - Selection</span></span><br><span data-ttu-id="c2fd6-699">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-699">
         - Settings</span></span><br><span data-ttu-id="c2fd6-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-700">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-701">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-702">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-703">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-705">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-705">Office 2019 on Mac</span></span><br><span data-ttu-id="c2fd6-706">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-707">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-707">- TaskPane</span></span><br><span data-ttu-id="c2fd6-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c2fd6-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c2fd6-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c2fd6-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-714">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-715">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-718">
         - File</span></span><br><span data-ttu-id="c2fd6-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-723">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-724">
         - Selection</span></span><br><span data-ttu-id="c2fd6-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-725">
         - Settings</span></span><br><span data-ttu-id="c2fd6-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-726">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-727">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-728">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-729">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-731">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-731">Office 2016 on Mac</span></span><br><span data-ttu-id="c2fd6-732">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-733">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c2fd6-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-737">- BindingEvents</span></span><br><span data-ttu-id="c2fd6-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-738">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c2fd6-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="c2fd6-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-740">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-741">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-741">
         - File</span></span><br><span data-ttu-id="c2fd6-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-743">
         - MatrixBindings</span></span><br><span data-ttu-id="c2fd6-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="c2fd6-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c2fd6-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-746">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-747">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-747">
         - Selection</span></span><br><span data-ttu-id="c2fd6-748">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-748">
         - Settings</span></span><br><span data-ttu-id="c2fd6-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-749">
         - TableBindings</span></span><br><span data-ttu-id="c2fd6-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-750">
         - TableCoercion</span></span><br><span data-ttu-id="c2fd6-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c2fd6-751">
         - TextBindings</span></span><br><span data-ttu-id="c2fd6-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-752">
         - TextCoercion</span></span><br><span data-ttu-id="c2fd6-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c2fd6-754">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c2fd6-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c2fd6-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c2fd6-756">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-756">Platform</span></span></th>
    <th><span data-ttu-id="c2fd6-757">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-757">Extension points</span></span></th>
    <th><span data-ttu-id="c2fd6-758">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="c2fd6-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-760">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="c2fd6-761">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-761">- Content</span></span><br><span data-ttu-id="c2fd6-762">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-762">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-768">- ActiveView</span></span><br><span data-ttu-id="c2fd6-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-769">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-770">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-771">
         - File</span></span><br><span data-ttu-id="c2fd6-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-772">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-773">
         - Selection</span></span><br><span data-ttu-id="c2fd6-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-774">
         - Settings</span></span><br><span data-ttu-id="c2fd6-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-776">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-776">Office on Windows</span></span><br><span data-ttu-id="c2fd6-777">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-778">- Content</span></span><br><span data-ttu-id="c2fd6-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-779">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-785">- ActiveView</span></span><br><span data-ttu-id="c2fd6-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-786">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-787">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-788">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-788">
         - File</span></span><br><span data-ttu-id="c2fd6-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-789">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-790">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-790">
         - Selection</span></span><br><span data-ttu-id="c2fd6-791">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-791">
         - Settings</span></span><br><span data-ttu-id="c2fd6-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-793">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-793">Office 2019 on Windows</span></span><br><span data-ttu-id="c2fd6-794">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-795">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-795">- Content</span></span><br><span data-ttu-id="c2fd6-796">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-796">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-800">- ActiveView</span></span><br><span data-ttu-id="c2fd6-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-801">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-802">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-803">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-803">
         - File</span></span><br><span data-ttu-id="c2fd6-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-804">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-805">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-805">
         - Selection</span></span><br><span data-ttu-id="c2fd6-806">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-806">
         - Settings</span></span><br><span data-ttu-id="c2fd6-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-808">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-808">Office 2016 on Windows</span></span><br><span data-ttu-id="c2fd6-809">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-810">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-810">- Content</span></span><br><span data-ttu-id="c2fd6-811">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c2fd6-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-814">- ActiveView</span></span><br><span data-ttu-id="c2fd6-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-815">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-816">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-817">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-817">
         - File</span></span><br><span data-ttu-id="c2fd6-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-818">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-819">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-819">
         - Selection</span></span><br><span data-ttu-id="c2fd6-820">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-820">
         - Settings</span></span><br><span data-ttu-id="c2fd6-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-822">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-822">Office 2013 on Windows</span></span><br><span data-ttu-id="c2fd6-823">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-824">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-824">- Content</span></span><br><span data-ttu-id="c2fd6-825">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c2fd6-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c2fd6-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-828">- ActiveView</span></span><br><span data-ttu-id="c2fd6-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-829">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-830">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-831">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-831">
         - File</span></span><br><span data-ttu-id="c2fd6-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-832">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-833">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-833">
         - Selection</span></span><br><span data-ttu-id="c2fd6-834">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-834">
         - Settings</span></span><br><span data-ttu-id="c2fd6-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-836">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="c2fd6-836">Office on iPad</span></span><br><span data-ttu-id="c2fd6-837">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-838">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-838">- Content</span></span><br><span data-ttu-id="c2fd6-839">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-843">- ActiveView</span></span><br><span data-ttu-id="c2fd6-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-844">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-845">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-846">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-846">
         - File</span></span><br><span data-ttu-id="c2fd6-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-847">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-848">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-848">
         - Selection</span></span><br><span data-ttu-id="c2fd6-849">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-849">
         - Settings</span></span><br><span data-ttu-id="c2fd6-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-851">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-851">Office on Mac</span></span><br><span data-ttu-id="c2fd6-852">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c2fd6-853">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-853">- Content</span></span><br><span data-ttu-id="c2fd6-854">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-854">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c2fd6-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-860">- ActiveView</span></span><br><span data-ttu-id="c2fd6-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-861">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-862">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-863">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-863">
         - File</span></span><br><span data-ttu-id="c2fd6-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-864">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-865">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-865">
         - Selection</span></span><br><span data-ttu-id="c2fd6-866">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-866">
         - Settings</span></span><br><span data-ttu-id="c2fd6-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-868">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-868">Office 2019 on Mac</span></span><br><span data-ttu-id="c2fd6-869">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-870">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-870">- Content</span></span><br><span data-ttu-id="c2fd6-871">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-871">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-875">- ActiveView</span></span><br><span data-ttu-id="c2fd6-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-876">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-877">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-878">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-878">
         - File</span></span><br><span data-ttu-id="c2fd6-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-879">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-880">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-880">
         - Selection</span></span><br><span data-ttu-id="c2fd6-881">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-881">
         - Settings</span></span><br><span data-ttu-id="c2fd6-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-883">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-883">Office 2016 on Mac</span></span><br><span data-ttu-id="c2fd6-884">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-885">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-885">- Content</span></span><br><span data-ttu-id="c2fd6-886">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c2fd6-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c2fd6-889">- ActiveView</span></span><br><span data-ttu-id="c2fd6-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-890">
         - CompressedFile</span></span><br><span data-ttu-id="c2fd6-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-891">
         - DocumentEvents</span></span><br><span data-ttu-id="c2fd6-892">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-892">
         - File</span></span><br><span data-ttu-id="c2fd6-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c2fd6-893">
         - PdfFile</span></span><br><span data-ttu-id="c2fd6-894">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-894">
         - Selection</span></span><br><span data-ttu-id="c2fd6-895">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-895">
         - Settings</span></span><br><span data-ttu-id="c2fd6-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c2fd6-897">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="c2fd6-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c2fd6-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="c2fd6-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c2fd6-899">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-899">Platform</span></span></th>
    <th><span data-ttu-id="c2fd6-900">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-900">Extension points</span></span></th>
    <th><span data-ttu-id="c2fd6-901">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="c2fd6-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-903">Office na Web</span><span class="sxs-lookup"><span data-stu-id="c2fd6-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="c2fd6-904">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="c2fd6-904">- Content</span></span><br><span data-ttu-id="c2fd6-905">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-905">
         - TaskPane</span></span><br><span data-ttu-id="c2fd6-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c2fd6-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c2fd6-910">- DocumentEvents</span></span><br><span data-ttu-id="c2fd6-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="c2fd6-912">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="c2fd6-912">
         - Settings</span></span><br><span data-ttu-id="c2fd6-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c2fd6-914">Project</span><span class="sxs-lookup"><span data-stu-id="c2fd6-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c2fd6-915">Plataforma</span><span class="sxs-lookup"><span data-stu-id="c2fd6-915">Platform</span></span></th>
    <th><span data-ttu-id="c2fd6-916">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="c2fd6-916">Extension points</span></span></th>
    <th><span data-ttu-id="c2fd6-917">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="c2fd6-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-919">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-919">Office 2019 on Windows</span></span><br><span data-ttu-id="c2fd6-920">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-921">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-923">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-923">- Selection</span></span><br><span data-ttu-id="c2fd6-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-925">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-925">Office 2016 on Windows</span></span><br><span data-ttu-id="c2fd6-926">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-927">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-929">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-929">- Selection</span></span><br><span data-ttu-id="c2fd6-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c2fd6-931">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="c2fd6-931">Office 2013 on Windows</span></span><br><span data-ttu-id="c2fd6-932">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c2fd6-933">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="c2fd6-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c2fd6-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c2fd6-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c2fd6-935">- Seleção</span><span class="sxs-lookup"><span data-stu-id="c2fd6-935">- Selection</span></span><br><span data-ttu-id="c2fd6-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c2fd6-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c2fd6-937">Confira também</span><span class="sxs-lookup"><span data-stu-id="c2fd6-937">See also</span></span>

- [<span data-ttu-id="c2fd6-938">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="c2fd6-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c2fd6-939">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="c2fd6-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c2fd6-940">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c2fd6-941">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="c2fd6-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c2fd6-942">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="c2fd6-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c2fd6-943">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="c2fd6-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c2fd6-944">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c2fd6-945">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c2fd6-946">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c2fd6-947">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="c2fd6-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c2fd6-948">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="c2fd6-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c2fd6-949">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="c2fd6-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)