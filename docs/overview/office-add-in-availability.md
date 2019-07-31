---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/26/2019
localization_priority: Priority
ms.openlocfilehash: 7039ca59af22f1101bdff7b6bcd4506497d6c9cd
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940833"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="3bef7-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3bef7-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="3bef7-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="3bef7-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="3bef7-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="3bef7-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="3bef7-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="3bef7-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="3bef7-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="3bef7-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="3bef7-108">Excel</span><span class="sxs-lookup"><span data-stu-id="3bef7-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3bef7-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3bef7-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3bef7-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3bef7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="3bef7-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-114">- TaskPane</span></span><br><span data-ttu-id="3bef7-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-115">
        - Content</span></span><br><span data-ttu-id="3bef7-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-116">
        - Custom Functions</span></span><br><span data-ttu-id="3bef7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3bef7-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3bef7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3bef7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3bef7-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-130">
        - BindingEvents</span></span><br><span data-ttu-id="3bef7-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-131">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-132">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-133">
        - File</span></span><br><span data-ttu-id="3bef7-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-134">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-136">
        - Selection</span></span><br><span data-ttu-id="3bef7-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-137">
        - Settings</span></span><br><span data-ttu-id="3bef7-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-138">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-139">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-140">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-142">Office on Windows</span></span><br><span data-ttu-id="3bef7-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-144">- TaskPane</span></span><br><span data-ttu-id="3bef7-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-145">
        - Content</span></span><br><span data-ttu-id="3bef7-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-146">
        - Custom Functions</span></span><br><span data-ttu-id="3bef7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="3bef7-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="3bef7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3bef7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3bef7-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-160">
        - BindingEvents</span></span><br><span data-ttu-id="3bef7-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-161">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-162">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-163">
        - File</span></span><br><span data-ttu-id="3bef7-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-164">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-166">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-166">
        - Selection</span></span><br><span data-ttu-id="3bef7-167">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-167">
        - Settings</span></span><br><span data-ttu-id="3bef7-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-168">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-169">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-170">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-172">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-172">Office 2019 on Windows</span></span><br><span data-ttu-id="3bef7-173">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3bef7-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-174">- TaskPane</span></span><br><span data-ttu-id="3bef7-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-175">
        - Content</span></span><br><span data-ttu-id="3bef7-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3bef7-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-187">- BindingEvents</span></span><br><span data-ttu-id="3bef7-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-188">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-189">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-190">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-190">
        - File</span></span><br><span data-ttu-id="3bef7-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-191">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-193">
        - Selection</span></span><br><span data-ttu-id="3bef7-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-194">
        - Settings</span></span><br><span data-ttu-id="3bef7-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-195">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-196">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-197">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-199">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-199">Office 2016 on Windows</span></span><br><span data-ttu-id="3bef7-200">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3bef7-201">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-201">- TaskPane</span></span><br><span data-ttu-id="3bef7-202">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-202">
        - Content</span></span></td>
    <td><span data-ttu-id="3bef7-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3bef7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-206">- BindingEvents</span></span><br><span data-ttu-id="3bef7-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-207">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-208">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-209">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-209">
        - File</span></span><br><span data-ttu-id="3bef7-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-210">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-212">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-212">
        - Selection</span></span><br><span data-ttu-id="3bef7-213">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-213">
        - Settings</span></span><br><span data-ttu-id="3bef7-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-214">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-215">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-216">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-218">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-218">Office 2013 on Windows</span></span><br><span data-ttu-id="3bef7-219">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3bef7-220">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-220">
        - TaskPane</span></span><br><span data-ttu-id="3bef7-221">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="3bef7-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3bef7-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3bef7-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-224">
        - BindingEvents</span></span><br><span data-ttu-id="3bef7-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-225">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-226">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-227">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-227">
        - File</span></span><br><span data-ttu-id="3bef7-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-228">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-230">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-230">
        - Selection</span></span><br><span data-ttu-id="3bef7-231">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-231">
        - Settings</span></span><br><span data-ttu-id="3bef7-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-232">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-233">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-234">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-236">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3bef7-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3bef7-237">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3bef7-238">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-238">- TaskPane</span></span><br><span data-ttu-id="3bef7-239">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-239">
        - Content</span></span><br><span data-ttu-id="3bef7-240">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3bef7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3bef7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-252">- BindingEvents</span></span><br><span data-ttu-id="3bef7-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-253">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-254">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-254">
        - File</span></span><br><span data-ttu-id="3bef7-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-255">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-257">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-257">
        - Selection</span></span><br><span data-ttu-id="3bef7-258">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-258">
        - Settings</span></span><br><span data-ttu-id="3bef7-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-259">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-260">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-261">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-263">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-263">Office apps on Mac</span></span><br><span data-ttu-id="3bef7-264">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3bef7-265">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-265">- TaskPane</span></span><br><span data-ttu-id="3bef7-266">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-266">
        - Content</span></span><br><span data-ttu-id="3bef7-267">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-267">
        - Custom Functions</span></span><br><span data-ttu-id="3bef7-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3bef7-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="3bef7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="3bef7-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-281">- BindingEvents</span></span><br><span data-ttu-id="3bef7-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-282">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-283">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-284">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-284">
        - File</span></span><br><span data-ttu-id="3bef7-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-285">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-287">
        - PdfFile</span></span><br><span data-ttu-id="3bef7-288">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-288">
        - Selection</span></span><br><span data-ttu-id="3bef7-289">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-289">
        - Settings</span></span><br><span data-ttu-id="3bef7-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-290">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-291">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-292">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-294">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-294">Office 2019 for Mac</span></span><br><span data-ttu-id="3bef7-295">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3bef7-296">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-296">- TaskPane</span></span><br><span data-ttu-id="3bef7-297">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-297">
        - Content</span></span><br><span data-ttu-id="3bef7-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="3bef7-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="3bef7-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="3bef7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="3bef7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="3bef7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="3bef7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="3bef7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="3bef7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-309">- BindingEvents</span></span><br><span data-ttu-id="3bef7-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-310">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-311">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-312">
        - File</span></span><br><span data-ttu-id="3bef7-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-313">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-315">
        - PdfFile</span></span><br><span data-ttu-id="3bef7-316">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-316">
        - Selection</span></span><br><span data-ttu-id="3bef7-317">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-317">
        - Settings</span></span><br><span data-ttu-id="3bef7-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-318">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-319">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-320">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-322">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3bef7-323">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="3bef7-324">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-324">- TaskPane</span></span><br><span data-ttu-id="3bef7-325">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-325">
        - Content</span></span></td>
    <td><span data-ttu-id="3bef7-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="3bef7-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3bef7-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="3bef7-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-329">- BindingEvents</span></span><br><span data-ttu-id="3bef7-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-330">
        - CompressedFile</span></span><br><span data-ttu-id="3bef7-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-331">
        - DocumentEvents</span></span><br><span data-ttu-id="3bef7-332">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-332">
        - File</span></span><br><span data-ttu-id="3bef7-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-333">
        - MatrixBindings</span></span><br><span data-ttu-id="3bef7-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-335">
        - PdfFile</span></span><br><span data-ttu-id="3bef7-336">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-336">
        - Selection</span></span><br><span data-ttu-id="3bef7-337">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-337">
        - Settings</span></span><br><span data-ttu-id="3bef7-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-338">
        - TableBindings</span></span><br><span data-ttu-id="3bef7-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-339">
        - TableCoercion</span></span><br><span data-ttu-id="3bef7-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-340">
        - TextBindings</span></span><br><span data-ttu-id="3bef7-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3bef7-342">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3bef7-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="3bef7-343">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="3bef7-344">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="3bef7-345">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="3bef7-346">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="3bef7-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-348">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-348">Office on the web</span></span></td>
    <td><span data-ttu-id="3bef7-349">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3bef7-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-351">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-351">Office on Windows</span></span><br><span data-ttu-id="3bef7-352">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="3bef7-353">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3bef7-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-355">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-355">Office for Mac</span></span><br><span data-ttu-id="3bef7-356">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="3bef7-357">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3bef7-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="3bef7-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="3bef7-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="3bef7-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3bef7-360">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-360">Platform</span></span></th>
    <th><span data-ttu-id="3bef7-361">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-361">Extension points</span></span></th>
    <th><span data-ttu-id="3bef7-362">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="3bef7-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-364">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-364">Office on the web</span></span><br><span data-ttu-id="3bef7-365">(moderno)</span><span class="sxs-lookup"><span data-stu-id="3bef7-365">Modern</span></span></td>
    <td> <span data-ttu-id="3bef7-366">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-366">- Mail Read</span></span><br><span data-ttu-id="3bef7-367">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-367">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3bef7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3bef7-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-377">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-377">Office on the web</span></span><br><span data-ttu-id="3bef7-378">(clássico)</span><span class="sxs-lookup"><span data-stu-id="3bef7-378">Classic.</span></span></td>
    <td> <span data-ttu-id="3bef7-379">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-379">- Mail Read</span></span><br><span data-ttu-id="3bef7-380">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-380">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3bef7-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-389">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-389">Office on Windows</span></span><br><span data-ttu-id="3bef7-390">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-391">- Mail Read</span></span><br><span data-ttu-id="3bef7-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-392">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3bef7-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="3bef7-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3bef7-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3bef7-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3bef7-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-403">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-403">Office 2019 on Windows</span></span><br><span data-ttu-id="3bef7-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-405">- Mail Read</span></span><br><span data-ttu-id="3bef7-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-406">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3bef7-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="3bef7-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3bef7-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3bef7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3bef7-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-417">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-417">Office 2016 on Windows</span></span><br><span data-ttu-id="3bef7-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-419">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-419">- Mail Read</span></span><br><span data-ttu-id="3bef7-420">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-420">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="3bef7-422">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="3bef7-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="3bef7-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3bef7-427">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-428">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-428">Office 2013 on Windows</span></span><br><span data-ttu-id="3bef7-429">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-430">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-430">- Mail Read</span></span><br><span data-ttu-id="3bef7-431">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="3bef7-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="3bef7-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="3bef7-436">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-437">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="3bef7-437">Office apps on iOS</span></span><br><span data-ttu-id="3bef7-438">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-439">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-439">- Mail Read</span></span><br><span data-ttu-id="3bef7-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3bef7-446">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-447">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-447">Office apps on Mac</span></span><br><span data-ttu-id="3bef7-448">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-449">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-449">- Mail Read</span></span><br><span data-ttu-id="3bef7-450">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-450">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="3bef7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="3bef7-459">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-460">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-460">Office 2019 for Mac</span></span><br><span data-ttu-id="3bef7-461">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-462">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-462">- Mail Read</span></span><br><span data-ttu-id="3bef7-463">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-463">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3bef7-471">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-472">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3bef7-473">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-474">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-474">- Mail Read</span></span><br><span data-ttu-id="3bef7-475">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-475">
      - Mail Compose</span></span><br><span data-ttu-id="3bef7-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="3bef7-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="3bef7-483">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-484">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="3bef7-484">Office apps on Android</span></span><br><span data-ttu-id="3bef7-485">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-486">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="3bef7-486">- Mail Read</span></span><br><span data-ttu-id="3bef7-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="3bef7-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="3bef7-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="3bef7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="3bef7-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="3bef7-493">Não disponível</span><span class="sxs-lookup"><span data-stu-id="3bef7-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="3bef7-494">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3bef7-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="3bef7-495">Word</span><span class="sxs-lookup"><span data-stu-id="3bef7-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3bef7-496">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-496">Platform</span></span></th>
    <th><span data-ttu-id="3bef7-497">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-497">Extension points</span></span></th>
    <th><span data-ttu-id="3bef7-498">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="3bef7-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-500">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="3bef7-501">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-501">- TaskPane</span></span><br><span data-ttu-id="3bef7-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3bef7-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-509">- BindingEvents</span></span><br><span data-ttu-id="3bef7-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-511">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-512">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-512">
         - File</span></span><br><span data-ttu-id="3bef7-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-514">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-517">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-518">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-518">
         - Selection</span></span><br><span data-ttu-id="3bef7-519">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-519">
         - Settings</span></span><br><span data-ttu-id="3bef7-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-520">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-521">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-522">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-523">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-525">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-525">Office on Windows</span></span><br><span data-ttu-id="3bef7-526">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-527">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-527">- TaskPane</span></span><br><span data-ttu-id="3bef7-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3bef7-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-535">- BindingEvents</span></span><br><span data-ttu-id="3bef7-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-536">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-538">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-539">
         - File</span></span><br><span data-ttu-id="3bef7-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-541">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-544">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-545">
         - Selection</span></span><br><span data-ttu-id="3bef7-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-546">
         - Settings</span></span><br><span data-ttu-id="3bef7-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-547">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-548">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-549">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-550">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-552">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-552">Office 2019 on Windows</span></span><br><span data-ttu-id="3bef7-553">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-554">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-554">- TaskPane</span></span><br><span data-ttu-id="3bef7-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-561">- BindingEvents</span></span><br><span data-ttu-id="3bef7-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-562">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-564">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-565">
         - File</span></span><br><span data-ttu-id="3bef7-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-567">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-570">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-571">
         - Selection</span></span><br><span data-ttu-id="3bef7-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-572">
         - Settings</span></span><br><span data-ttu-id="3bef7-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-573">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-574">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-575">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-576">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-578">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-578">Office 2016 on Windows</span></span><br><span data-ttu-id="3bef7-579">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-580">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3bef7-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-584">- BindingEvents</span></span><br><span data-ttu-id="3bef7-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-585">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-587">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-588">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-588">
         - File</span></span><br><span data-ttu-id="3bef7-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-590">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-593">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-594">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-594">
         - Selection</span></span><br><span data-ttu-id="3bef7-595">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-595">
         - Settings</span></span><br><span data-ttu-id="3bef7-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-596">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-597">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-598">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-599">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-601">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-601">Office 2013 on Windows</span></span><br><span data-ttu-id="3bef7-602">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-603">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3bef7-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3bef7-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-606">- BindingEvents</span></span><br><span data-ttu-id="3bef7-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-607">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-609">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-610">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-610">
         - File</span></span><br><span data-ttu-id="3bef7-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-612">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-615">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-616">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-616">
         - Selection</span></span><br><span data-ttu-id="3bef7-617">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-617">
         - Settings</span></span><br><span data-ttu-id="3bef7-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-618">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-619">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-620">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-621">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-623">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3bef7-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3bef7-624">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-625">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3bef7-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-631">- BindingEvents</span></span><br><span data-ttu-id="3bef7-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-632">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-634">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-635">
         - File</span></span><br><span data-ttu-id="3bef7-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-637">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-640">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-641">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-641">
         - Selection</span></span><br><span data-ttu-id="3bef7-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-642">
         - Settings</span></span><br><span data-ttu-id="3bef7-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-643">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-644">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-645">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-646">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-648">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-648">Office apps on Mac</span></span><br><span data-ttu-id="3bef7-649">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-650">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-650">- TaskPane</span></span><br><span data-ttu-id="3bef7-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="3bef7-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-658">- BindingEvents</span></span><br><span data-ttu-id="3bef7-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-659">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-661">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-662">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-662">
         - File</span></span><br><span data-ttu-id="3bef7-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-664">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-667">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-668">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-668">
         - Selection</span></span><br><span data-ttu-id="3bef7-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-669">
         - Settings</span></span><br><span data-ttu-id="3bef7-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-670">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-671">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-672">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-673">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-675">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-675">Office 2019 for Mac</span></span><br><span data-ttu-id="3bef7-676">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-677">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-677">- TaskPane</span></span><br><span data-ttu-id="3bef7-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="3bef7-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="3bef7-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="3bef7-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-684">- BindingEvents</span></span><br><span data-ttu-id="3bef7-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-685">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-687">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-688">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-688">
         - File</span></span><br><span data-ttu-id="3bef7-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-690">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-693">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-694">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-694">
         - Selection</span></span><br><span data-ttu-id="3bef7-695">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-695">
         - Settings</span></span><br><span data-ttu-id="3bef7-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-696">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-697">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-698">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-699">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-701">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3bef7-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="3bef7-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="3bef7-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="3bef7-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-707">- BindingEvents</span></span><br><span data-ttu-id="3bef7-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-708">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="3bef7-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="3bef7-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-710">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-711">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-711">
         - File</span></span><br><span data-ttu-id="3bef7-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-713">
         - MatrixBindings</span></span><br><span data-ttu-id="3bef7-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="3bef7-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="3bef7-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-716">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-717">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-717">
         - Selection</span></span><br><span data-ttu-id="3bef7-718">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-718">
         - Settings</span></span><br><span data-ttu-id="3bef7-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-719">
         - TableBindings</span></span><br><span data-ttu-id="3bef7-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-720">
         - TableCoercion</span></span><br><span data-ttu-id="3bef7-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="3bef7-721">
         - TextBindings</span></span><br><span data-ttu-id="3bef7-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-722">
         - TextCoercion</span></span><br><span data-ttu-id="3bef7-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="3bef7-724">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3bef7-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="3bef7-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="3bef7-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3bef7-726">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-726">Platform</span></span></th>
    <th><span data-ttu-id="3bef7-727">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-727">Extension points</span></span></th>
    <th><span data-ttu-id="3bef7-728">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="3bef7-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-730">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="3bef7-731">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-731">- Content</span></span><br><span data-ttu-id="3bef7-732">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-732">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3bef7-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3bef7-738">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-738">- ActiveView</span></span><br><span data-ttu-id="3bef7-739">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-739">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-740">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-741">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-741">
         - File</span></span><br><span data-ttu-id="3bef7-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-742">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-743">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-743">
         - Selection</span></span><br><span data-ttu-id="3bef7-744">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-744">
         - Settings</span></span><br><span data-ttu-id="3bef7-745">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-745">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-746">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-746">Office on Windows</span></span><br><span data-ttu-id="3bef7-747">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-747">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-748">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-748">- Content</span></span><br><span data-ttu-id="3bef7-749">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-749">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3bef7-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3bef7-755">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-755">- ActiveView</span></span><br><span data-ttu-id="3bef7-756">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-756">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-757">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-757">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-758">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-758">
         - File</span></span><br><span data-ttu-id="3bef7-759">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-759">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-760">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-760">
         - Selection</span></span><br><span data-ttu-id="3bef7-761">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-761">
         - Settings</span></span><br><span data-ttu-id="3bef7-762">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-762">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-763">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-763">Office 2019 on Windows</span></span><br><span data-ttu-id="3bef7-764">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-764">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-765">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-765">- Content</span></span><br><span data-ttu-id="3bef7-766">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-766">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-770">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-770">- ActiveView</span></span><br><span data-ttu-id="3bef7-771">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-771">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-772">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-772">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-773">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-773">
         - File</span></span><br><span data-ttu-id="3bef7-774">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-774">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-775">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-775">
         - Selection</span></span><br><span data-ttu-id="3bef7-776">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-776">
         - Settings</span></span><br><span data-ttu-id="3bef7-777">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-777">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-778">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-778">Office 2016 on Windows</span></span><br><span data-ttu-id="3bef7-779">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-779">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-780">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-780">- Content</span></span><br><span data-ttu-id="3bef7-781">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-781">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3bef7-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3bef7-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-784">- ActiveView</span></span><br><span data-ttu-id="3bef7-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-785">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-786">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-787">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-787">
         - File</span></span><br><span data-ttu-id="3bef7-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-788">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-789">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-789">
         - Selection</span></span><br><span data-ttu-id="3bef7-790">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-790">
         - Settings</span></span><br><span data-ttu-id="3bef7-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-792">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-792">Office 2013 on Windows</span></span><br><span data-ttu-id="3bef7-793">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-794">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-794">- Content</span></span><br><span data-ttu-id="3bef7-795">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-795">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="3bef7-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3bef7-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3bef7-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-798">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-798">- ActiveView</span></span><br><span data-ttu-id="3bef7-799">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-799">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-800">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-800">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-801">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-801">
         - File</span></span><br><span data-ttu-id="3bef7-802">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-802">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-803">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-803">
         - Selection</span></span><br><span data-ttu-id="3bef7-804">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-804">
         - Settings</span></span><br><span data-ttu-id="3bef7-805">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-805">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-806">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="3bef7-806">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3bef7-807">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-807">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-808">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-808">- Content</span></span><br><span data-ttu-id="3bef7-809">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-809">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3bef7-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-813">- ActiveView</span></span><br><span data-ttu-id="3bef7-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-814">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-815">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-816">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-816">
         - File</span></span><br><span data-ttu-id="3bef7-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-817">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-818">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-818">
         - Selection</span></span><br><span data-ttu-id="3bef7-819">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-819">
         - Settings</span></span><br><span data-ttu-id="3bef7-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-821">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-821">Office apps on Mac</span></span><br><span data-ttu-id="3bef7-822">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="3bef7-822">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="3bef7-823">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-823">- Content</span></span><br><span data-ttu-id="3bef7-824">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-824">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="3bef7-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="3bef7-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="3bef7-830">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-830">- ActiveView</span></span><br><span data-ttu-id="3bef7-831">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-831">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-832">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-832">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-833">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-833">
         - File</span></span><br><span data-ttu-id="3bef7-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-834">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-835">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-835">
         - Selection</span></span><br><span data-ttu-id="3bef7-836">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-836">
         - Settings</span></span><br><span data-ttu-id="3bef7-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-838">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-838">Office 2019 for Mac</span></span><br><span data-ttu-id="3bef7-839">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-840">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-840">- Content</span></span><br><span data-ttu-id="3bef7-841">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-841">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-845">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-845">- ActiveView</span></span><br><span data-ttu-id="3bef7-846">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-846">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-847">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-847">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-848">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-848">
         - File</span></span><br><span data-ttu-id="3bef7-849">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-849">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-850">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-850">
         - Selection</span></span><br><span data-ttu-id="3bef7-851">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-851">
         - Settings</span></span><br><span data-ttu-id="3bef7-852">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-852">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-853">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-853">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="3bef7-854">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-854">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-855">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-855">- Content</span></span><br><span data-ttu-id="3bef7-856">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-856">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="3bef7-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="3bef7-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="3bef7-859">- ActiveView</span></span><br><span data-ttu-id="3bef7-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-860">
         - CompressedFile</span></span><br><span data-ttu-id="3bef7-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-861">
         - DocumentEvents</span></span><br><span data-ttu-id="3bef7-862">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="3bef7-862">
         - File</span></span><br><span data-ttu-id="3bef7-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="3bef7-863">
         - PdfFile</span></span><br><span data-ttu-id="3bef7-864">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-864">
         - Selection</span></span><br><span data-ttu-id="3bef7-865">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-865">
         - Settings</span></span><br><span data-ttu-id="3bef7-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-866">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="3bef7-867">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="3bef7-867">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="3bef7-868">OneNote</span><span class="sxs-lookup"><span data-stu-id="3bef7-868">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3bef7-869">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-869">Platform</span></span></th>
    <th><span data-ttu-id="3bef7-870">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-870">Extension points</span></span></th>
    <th><span data-ttu-id="3bef7-871">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-871">API requirement sets</span></span></th>
    <th><span data-ttu-id="3bef7-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-873">Office na Web</span><span class="sxs-lookup"><span data-stu-id="3bef7-873">Office on the web</span></span></td>
    <td> <span data-ttu-id="3bef7-874">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="3bef7-874">- Content</span></span><br><span data-ttu-id="3bef7-875">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-875">
         - TaskPane</span></span><br><span data-ttu-id="3bef7-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="3bef7-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="3bef7-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="3bef7-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-880">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="3bef7-880">- DocumentEvents</span></span><br><span data-ttu-id="3bef7-881">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-881">
         - HtmlCoercion</span></span><br><span data-ttu-id="3bef7-882">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="3bef7-882">
         - Settings</span></span><br><span data-ttu-id="3bef7-883">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-883">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="3bef7-884">Project</span><span class="sxs-lookup"><span data-stu-id="3bef7-884">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="3bef7-885">Plataforma</span><span class="sxs-lookup"><span data-stu-id="3bef7-885">Platform</span></span></th>
    <th><span data-ttu-id="3bef7-886">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="3bef7-886">Extension points</span></span></th>
    <th><span data-ttu-id="3bef7-887">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-887">API requirement sets</span></span></th>
    <th><span data-ttu-id="3bef7-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="3bef7-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-889">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-889">Office 2019 on Windows</span></span><br><span data-ttu-id="3bef7-890">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-890">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-891">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-891">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-893">- Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-893">- Selection</span></span><br><span data-ttu-id="3bef7-894">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-894">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-895">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-895">Office 2016 on Windows</span></span><br><span data-ttu-id="3bef7-896">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-896">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-897">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-897">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-899">- Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-899">- Selection</span></span><br><span data-ttu-id="3bef7-900">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-900">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="3bef7-901">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="3bef7-901">Office 2013 on Windows</span></span><br><span data-ttu-id="3bef7-902">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="3bef7-902">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="3bef7-903">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="3bef7-903">- TaskPane</span></span></td>
    <td> <span data-ttu-id="3bef7-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="3bef7-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="3bef7-905">- Seleção</span><span class="sxs-lookup"><span data-stu-id="3bef7-905">- Selection</span></span><br><span data-ttu-id="3bef7-906">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="3bef7-906">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="3bef7-907">Confira também</span><span class="sxs-lookup"><span data-stu-id="3bef7-907">See also</span></span>

- [<span data-ttu-id="3bef7-908">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="3bef7-908">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="3bef7-909">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="3bef7-909">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="3bef7-910">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="3bef7-910">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="3bef7-911">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="3bef7-911">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="3bef7-912">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="3bef7-912">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="3bef7-913">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="3bef7-913">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="3bef7-914">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="3bef7-914">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="3bef7-915">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="3bef7-915">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="3bef7-916">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3bef7-916">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="3bef7-917">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="3bef7-917">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="3bef7-918">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="3bef7-918">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
