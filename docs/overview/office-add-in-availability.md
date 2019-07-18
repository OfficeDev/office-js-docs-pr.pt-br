---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: 2bfeb7cc5c6e8846f1d882abf3a0149302e53914
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771832"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f0d07-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f0d07-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f0d07-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="f0d07-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="f0d07-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="f0d07-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="f0d07-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="f0d07-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="f0d07-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="f0d07-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="f0d07-108">Excel</span><span class="sxs-lookup"><span data-stu-id="f0d07-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f0d07-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f0d07-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f0d07-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f0d07-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="f0d07-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-114">- TaskPane</span></span><br><span data-ttu-id="f0d07-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-115">
        - Content</span></span><br><span data-ttu-id="f0d07-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-116">
        - Custom Functions</span></span><br><span data-ttu-id="f0d07-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="f0d07-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f0d07-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f0d07-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f0d07-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-130">
        - BindingEvents</span></span><br><span data-ttu-id="f0d07-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-131">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-132">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-133">
        - File</span></span><br><span data-ttu-id="f0d07-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-134">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-136">
        - Selection</span></span><br><span data-ttu-id="f0d07-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-137">
        - Settings</span></span><br><span data-ttu-id="f0d07-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-138">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-139">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-140">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-142">Office on Windows</span></span><br><span data-ttu-id="f0d07-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-144">- TaskPane</span></span><br><span data-ttu-id="f0d07-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-145">
        - Content</span></span><br><span data-ttu-id="f0d07-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-146">
        - Custom Functions</span></span><br><span data-ttu-id="f0d07-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="f0d07-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f0d07-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f0d07-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f0d07-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-160">
        - BindingEvents</span></span><br><span data-ttu-id="f0d07-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-161">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-162">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-163">
        - File</span></span><br><span data-ttu-id="f0d07-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-164">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-166">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-166">
        - Selection</span></span><br><span data-ttu-id="f0d07-167">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-167">
        - Settings</span></span><br><span data-ttu-id="f0d07-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-168">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-169">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-170">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-172">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-172">Office 2019 on Windows</span></span><br><span data-ttu-id="f0d07-173">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f0d07-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-174">- TaskPane</span></span><br><span data-ttu-id="f0d07-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-175">
        - Content</span></span><br><span data-ttu-id="f0d07-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f0d07-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-187">- BindingEvents</span></span><br><span data-ttu-id="f0d07-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-188">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-189">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-190">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-190">
        - File</span></span><br><span data-ttu-id="f0d07-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-191">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-193">
        - Selection</span></span><br><span data-ttu-id="f0d07-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-194">
        - Settings</span></span><br><span data-ttu-id="f0d07-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-195">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-196">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-197">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-199">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-199">Office 2016 on Windows</span></span><br><span data-ttu-id="f0d07-200">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f0d07-201">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-201">- TaskPane</span></span><br><span data-ttu-id="f0d07-202">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-202">
        - Content</span></span></td>
    <td><span data-ttu-id="f0d07-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f0d07-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-206">- BindingEvents</span></span><br><span data-ttu-id="f0d07-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-207">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-208">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-209">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-209">
        - File</span></span><br><span data-ttu-id="f0d07-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-210">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-212">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-212">
        - Selection</span></span><br><span data-ttu-id="f0d07-213">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-213">
        - Settings</span></span><br><span data-ttu-id="f0d07-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-214">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-215">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-216">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-218">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-218">Office 2013 on Windows</span></span><br><span data-ttu-id="f0d07-219">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f0d07-220">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-220">
        - TaskPane</span></span><br><span data-ttu-id="f0d07-221">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="f0d07-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f0d07-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f0d07-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-224">
        - BindingEvents</span></span><br><span data-ttu-id="f0d07-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-225">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-226">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-227">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-227">
        - File</span></span><br><span data-ttu-id="f0d07-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-228">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-230">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-230">
        - Selection</span></span><br><span data-ttu-id="f0d07-231">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-231">
        - Settings</span></span><br><span data-ttu-id="f0d07-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-232">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-233">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-234">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-236">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="f0d07-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f0d07-237">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f0d07-238">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-238">- TaskPane</span></span><br><span data-ttu-id="f0d07-239">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-239">
        - Content</span></span><br><span data-ttu-id="f0d07-240">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f0d07-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f0d07-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-252">- BindingEvents</span></span><br><span data-ttu-id="f0d07-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-253">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-254">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-254">
        - File</span></span><br><span data-ttu-id="f0d07-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-255">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-257">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-257">
        - Selection</span></span><br><span data-ttu-id="f0d07-258">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-258">
        - Settings</span></span><br><span data-ttu-id="f0d07-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-259">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-260">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-261">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-263">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-263">Office apps on Mac</span></span><br><span data-ttu-id="f0d07-264">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f0d07-265">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-265">- TaskPane</span></span><br><span data-ttu-id="f0d07-266">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-266">
        - Content</span></span><br><span data-ttu-id="f0d07-267">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-267">
        - Custom Functions</span></span><br><span data-ttu-id="f0d07-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f0d07-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f0d07-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f0d07-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-281">- BindingEvents</span></span><br><span data-ttu-id="f0d07-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-282">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-283">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-284">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-284">
        - File</span></span><br><span data-ttu-id="f0d07-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-285">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-287">
        - PdfFile</span></span><br><span data-ttu-id="f0d07-288">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-288">
        - Selection</span></span><br><span data-ttu-id="f0d07-289">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-289">
        - Settings</span></span><br><span data-ttu-id="f0d07-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-290">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-291">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-292">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-294">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-294">Office 2019 for Mac</span></span><br><span data-ttu-id="f0d07-295">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f0d07-296">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-296">- TaskPane</span></span><br><span data-ttu-id="f0d07-297">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-297">
        - Content</span></span><br><span data-ttu-id="f0d07-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f0d07-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f0d07-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f0d07-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f0d07-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f0d07-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f0d07-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f0d07-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f0d07-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-309">- BindingEvents</span></span><br><span data-ttu-id="f0d07-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-310">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-311">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-312">
        - File</span></span><br><span data-ttu-id="f0d07-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-313">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-315">
        - PdfFile</span></span><br><span data-ttu-id="f0d07-316">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-316">
        - Selection</span></span><br><span data-ttu-id="f0d07-317">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-317">
        - Settings</span></span><br><span data-ttu-id="f0d07-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-318">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-319">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-320">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-322">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="f0d07-323">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f0d07-324">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-324">- TaskPane</span></span><br><span data-ttu-id="f0d07-325">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-325">
        - Content</span></span></td>
    <td><span data-ttu-id="f0d07-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f0d07-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f0d07-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f0d07-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-329">- BindingEvents</span></span><br><span data-ttu-id="f0d07-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-330">
        - CompressedFile</span></span><br><span data-ttu-id="f0d07-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-331">
        - DocumentEvents</span></span><br><span data-ttu-id="f0d07-332">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-332">
        - File</span></span><br><span data-ttu-id="f0d07-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-333">
        - MatrixBindings</span></span><br><span data-ttu-id="f0d07-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-335">
        - PdfFile</span></span><br><span data-ttu-id="f0d07-336">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-336">
        - Selection</span></span><br><span data-ttu-id="f0d07-337">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-337">
        - Settings</span></span><br><span data-ttu-id="f0d07-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-338">
        - TableBindings</span></span><br><span data-ttu-id="f0d07-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-339">
        - TableCoercion</span></span><br><span data-ttu-id="f0d07-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-340">
        - TextBindings</span></span><br><span data-ttu-id="f0d07-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f0d07-342">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="f0d07-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="f0d07-343">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f0d07-344">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f0d07-345">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f0d07-346">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f0d07-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-348">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-348">Office on the web</span></span></td>
    <td><span data-ttu-id="f0d07-349">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f0d07-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-351">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-351">Office on Windows</span></span><br><span data-ttu-id="f0d07-352">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f0d07-353">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f0d07-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-355">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-355">Office for Mac</span></span><br><span data-ttu-id="f0d07-356">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="f0d07-357">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="f0d07-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f0d07-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="f0d07-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="f0d07-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f0d07-360">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-360">Platform</span></span></th>
    <th><span data-ttu-id="f0d07-361">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-361">Extension points</span></span></th>
    <th><span data-ttu-id="f0d07-362">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="f0d07-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-364">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-364">Office on the web</span></span><br><span data-ttu-id="f0d07-365">(novo)</span><span class="sxs-lookup"><span data-stu-id="f0d07-365">New</span></span></td>
    <td> <span data-ttu-id="f0d07-366">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-366">- Mail Read</span></span><br><span data-ttu-id="f0d07-367">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-367">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f0d07-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f0d07-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-377">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-377">Office on the web</span></span><br><span data-ttu-id="f0d07-378">(clássico)</span><span class="sxs-lookup"><span data-stu-id="f0d07-378">Classic.</span></span></td>
    <td> <span data-ttu-id="f0d07-379">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-379">- Mail Read</span></span><br><span data-ttu-id="f0d07-380">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-380">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f0d07-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-389">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-389">Office on Windows</span></span><br><span data-ttu-id="f0d07-390">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-391">- Mail Read</span></span><br><span data-ttu-id="f0d07-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-392">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f0d07-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="f0d07-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f0d07-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f0d07-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f0d07-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-403">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-403">Office 2019 on Windows</span></span><br><span data-ttu-id="f0d07-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-405">- Mail Read</span></span><br><span data-ttu-id="f0d07-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-406">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f0d07-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="f0d07-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f0d07-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f0d07-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f0d07-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-417">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-417">Office 2016 on Windows</span></span><br><span data-ttu-id="f0d07-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-419">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-419">- Mail Read</span></span><br><span data-ttu-id="f0d07-420">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-420">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f0d07-422">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="f0d07-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="f0d07-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f0d07-427">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-428">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-428">Office 2013 on Windows</span></span><br><span data-ttu-id="f0d07-429">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-430">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-430">- Mail Read</span></span><br><span data-ttu-id="f0d07-431">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="f0d07-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="f0d07-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f0d07-436">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-437">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="f0d07-437">Office apps on iOS</span></span><br><span data-ttu-id="f0d07-438">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-439">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-439">- Mail Read</span></span><br><span data-ttu-id="f0d07-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f0d07-446">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-447">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-447">Office apps on Mac</span></span><br><span data-ttu-id="f0d07-448">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-449">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-449">- Mail Read</span></span><br><span data-ttu-id="f0d07-450">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-450">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f0d07-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f0d07-459">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-460">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-460">Office 2019 for Mac</span></span><br><span data-ttu-id="f0d07-461">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-462">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-462">- Mail Read</span></span><br><span data-ttu-id="f0d07-463">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-463">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f0d07-471">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-472">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="f0d07-473">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-474">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-474">- Mail Read</span></span><br><span data-ttu-id="f0d07-475">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-475">
      - Mail Compose</span></span><br><span data-ttu-id="f0d07-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f0d07-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f0d07-483">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-484">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="f0d07-484">Office apps on Android</span></span><br><span data-ttu-id="f0d07-485">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-486">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="f0d07-486">- Mail Read</span></span><br><span data-ttu-id="f0d07-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f0d07-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f0d07-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f0d07-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f0d07-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f0d07-493">Não disponível</span><span class="sxs-lookup"><span data-stu-id="f0d07-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="f0d07-494">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="f0d07-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="f0d07-495">Word</span><span class="sxs-lookup"><span data-stu-id="f0d07-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f0d07-496">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-496">Platform</span></span></th>
    <th><span data-ttu-id="f0d07-497">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-497">Extension points</span></span></th>
    <th><span data-ttu-id="f0d07-498">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="f0d07-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-500">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="f0d07-501">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-501">- TaskPane</span></span><br><span data-ttu-id="f0d07-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f0d07-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-509">- BindingEvents</span></span><br><span data-ttu-id="f0d07-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-511">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-512">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-512">
         - File</span></span><br><span data-ttu-id="f0d07-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-514">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-517">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-518">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-518">
         - Selection</span></span><br><span data-ttu-id="f0d07-519">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-519">
         - Settings</span></span><br><span data-ttu-id="f0d07-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-520">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-521">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-522">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-523">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-525">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-525">Office on Windows</span></span><br><span data-ttu-id="f0d07-526">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-527">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-527">- TaskPane</span></span><br><span data-ttu-id="f0d07-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f0d07-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-535">- BindingEvents</span></span><br><span data-ttu-id="f0d07-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-536">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-538">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-539">
         - File</span></span><br><span data-ttu-id="f0d07-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-541">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-544">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-545">
         - Selection</span></span><br><span data-ttu-id="f0d07-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-546">
         - Settings</span></span><br><span data-ttu-id="f0d07-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-547">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-548">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-549">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-550">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-552">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-552">Office 2019 on Windows</span></span><br><span data-ttu-id="f0d07-553">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-554">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-554">- TaskPane</span></span><br><span data-ttu-id="f0d07-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-561">- BindingEvents</span></span><br><span data-ttu-id="f0d07-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-562">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-564">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-565">
         - File</span></span><br><span data-ttu-id="f0d07-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-567">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-570">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-571">
         - Selection</span></span><br><span data-ttu-id="f0d07-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-572">
         - Settings</span></span><br><span data-ttu-id="f0d07-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-573">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-574">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-575">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-576">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-578">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-578">Office 2016 on Windows</span></span><br><span data-ttu-id="f0d07-579">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-580">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f0d07-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-584">- BindingEvents</span></span><br><span data-ttu-id="f0d07-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-585">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-587">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-588">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-588">
         - File</span></span><br><span data-ttu-id="f0d07-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-590">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-593">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-594">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-594">
         - Selection</span></span><br><span data-ttu-id="f0d07-595">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-595">
         - Settings</span></span><br><span data-ttu-id="f0d07-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-596">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-597">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-598">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-599">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-601">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-601">Office 2013 on Windows</span></span><br><span data-ttu-id="f0d07-602">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-603">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f0d07-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f0d07-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-606">- BindingEvents</span></span><br><span data-ttu-id="f0d07-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-607">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-609">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-610">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-610">
         - File</span></span><br><span data-ttu-id="f0d07-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-612">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-615">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-616">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-616">
         - Selection</span></span><br><span data-ttu-id="f0d07-617">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-617">
         - Settings</span></span><br><span data-ttu-id="f0d07-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-618">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-619">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-620">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-621">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-623">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="f0d07-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f0d07-624">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-625">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f0d07-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-631">- BindingEvents</span></span><br><span data-ttu-id="f0d07-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-632">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-634">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-635">
         - File</span></span><br><span data-ttu-id="f0d07-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-637">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-640">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-641">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-641">
         - Selection</span></span><br><span data-ttu-id="f0d07-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-642">
         - Settings</span></span><br><span data-ttu-id="f0d07-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-643">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-644">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-645">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-646">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-648">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-648">Office apps on Mac</span></span><br><span data-ttu-id="f0d07-649">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-650">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-650">- TaskPane</span></span><br><span data-ttu-id="f0d07-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="f0d07-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-658">- BindingEvents</span></span><br><span data-ttu-id="f0d07-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-659">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-661">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-662">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-662">
         - File</span></span><br><span data-ttu-id="f0d07-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-664">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-667">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-668">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-668">
         - Selection</span></span><br><span data-ttu-id="f0d07-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-669">
         - Settings</span></span><br><span data-ttu-id="f0d07-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-670">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-671">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-672">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-673">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-675">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-675">Office 2019 for Mac</span></span><br><span data-ttu-id="f0d07-676">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-677">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-677">- TaskPane</span></span><br><span data-ttu-id="f0d07-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="f0d07-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="f0d07-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f0d07-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-684">- BindingEvents</span></span><br><span data-ttu-id="f0d07-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-685">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-687">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-688">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-688">
         - File</span></span><br><span data-ttu-id="f0d07-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-690">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-693">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-694">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-694">
         - Selection</span></span><br><span data-ttu-id="f0d07-695">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-695">
         - Settings</span></span><br><span data-ttu-id="f0d07-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-696">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-697">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-698">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-699">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-701">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="f0d07-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="f0d07-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f0d07-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f0d07-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-707">- BindingEvents</span></span><br><span data-ttu-id="f0d07-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-708">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f0d07-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="f0d07-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-710">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-711">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-711">
         - File</span></span><br><span data-ttu-id="f0d07-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-713">
         - MatrixBindings</span></span><br><span data-ttu-id="f0d07-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="f0d07-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f0d07-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-716">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-717">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-717">
         - Selection</span></span><br><span data-ttu-id="f0d07-718">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-718">
         - Settings</span></span><br><span data-ttu-id="f0d07-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-719">
         - TableBindings</span></span><br><span data-ttu-id="f0d07-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-720">
         - TableCoercion</span></span><br><span data-ttu-id="f0d07-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f0d07-721">
         - TextBindings</span></span><br><span data-ttu-id="f0d07-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-722">
         - TextCoercion</span></span><br><span data-ttu-id="f0d07-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="f0d07-724">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="f0d07-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f0d07-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f0d07-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f0d07-726">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-726">Platform</span></span></th>
    <th><span data-ttu-id="f0d07-727">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-727">Extension points</span></span></th>
    <th><span data-ttu-id="f0d07-728">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="f0d07-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-730">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="f0d07-731">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-731">- Content</span></span><br><span data-ttu-id="f0d07-732">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-732">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f0d07-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-737">- ActiveView</span></span><br><span data-ttu-id="f0d07-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-738">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-739">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-740">
         - File</span></span><br><span data-ttu-id="f0d07-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-741">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-742">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-742">
         - Selection</span></span><br><span data-ttu-id="f0d07-743">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-743">
         - Settings</span></span><br><span data-ttu-id="f0d07-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-745">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-745">Office on Windows</span></span><br><span data-ttu-id="f0d07-746">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-747">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-747">- Content</span></span><br><span data-ttu-id="f0d07-748">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-748">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f0d07-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-753">- ActiveView</span></span><br><span data-ttu-id="f0d07-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-754">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-755">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-756">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-756">
         - File</span></span><br><span data-ttu-id="f0d07-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-757">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-758">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-758">
         - Selection</span></span><br><span data-ttu-id="f0d07-759">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-759">
         - Settings</span></span><br><span data-ttu-id="f0d07-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-761">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-761">Office 2019 on Windows</span></span><br><span data-ttu-id="f0d07-762">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-763">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-763">- Content</span></span><br><span data-ttu-id="f0d07-764">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-764">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-768">- ActiveView</span></span><br><span data-ttu-id="f0d07-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-769">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-770">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-771">
         - File</span></span><br><span data-ttu-id="f0d07-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-772">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-773">
         - Selection</span></span><br><span data-ttu-id="f0d07-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-774">
         - Settings</span></span><br><span data-ttu-id="f0d07-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-776">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-776">Office 2016 on Windows</span></span><br><span data-ttu-id="f0d07-777">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-778">- Content</span></span><br><span data-ttu-id="f0d07-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f0d07-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f0d07-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-782">- ActiveView</span></span><br><span data-ttu-id="f0d07-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-783">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-784">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-785">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-785">
         - File</span></span><br><span data-ttu-id="f0d07-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-786">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-787">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-787">
         - Selection</span></span><br><span data-ttu-id="f0d07-788">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-788">
         - Settings</span></span><br><span data-ttu-id="f0d07-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-790">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-790">Office 2013 on Windows</span></span><br><span data-ttu-id="f0d07-791">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-792">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-792">- Content</span></span><br><span data-ttu-id="f0d07-793">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="f0d07-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f0d07-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f0d07-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-796">- ActiveView</span></span><br><span data-ttu-id="f0d07-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-797">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-798">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-799">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-799">
         - File</span></span><br><span data-ttu-id="f0d07-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-800">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-801">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-801">
         - Selection</span></span><br><span data-ttu-id="f0d07-802">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-802">
         - Settings</span></span><br><span data-ttu-id="f0d07-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-804">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="f0d07-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="f0d07-805">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-806">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-806">- Content</span></span><br><span data-ttu-id="f0d07-807">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-810">- ActiveView</span></span><br><span data-ttu-id="f0d07-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-811">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-812">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-813">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-813">
         - File</span></span><br><span data-ttu-id="f0d07-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-814">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-815">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-815">
         - Selection</span></span><br><span data-ttu-id="f0d07-816">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-816">
         - Settings</span></span><br><span data-ttu-id="f0d07-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-818">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-818">Office apps on Mac</span></span><br><span data-ttu-id="f0d07-819">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="f0d07-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f0d07-820">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-820">- Content</span></span><br><span data-ttu-id="f0d07-821">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-821">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f0d07-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f0d07-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-826">- ActiveView</span></span><br><span data-ttu-id="f0d07-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-827">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-828">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-829">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-829">
         - File</span></span><br><span data-ttu-id="f0d07-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-830">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-831">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-831">
         - Selection</span></span><br><span data-ttu-id="f0d07-832">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-832">
         - Settings</span></span><br><span data-ttu-id="f0d07-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-834">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-834">Office 2019 for Mac</span></span><br><span data-ttu-id="f0d07-835">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-836">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-836">- Content</span></span><br><span data-ttu-id="f0d07-837">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-837">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-841">- ActiveView</span></span><br><span data-ttu-id="f0d07-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-842">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-843">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-844">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-844">
         - File</span></span><br><span data-ttu-id="f0d07-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-845">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-846">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-846">
         - Selection</span></span><br><span data-ttu-id="f0d07-847">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-847">
         - Settings</span></span><br><span data-ttu-id="f0d07-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-849">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="f0d07-850">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-851">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-851">- Content</span></span><br><span data-ttu-id="f0d07-852">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f0d07-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f0d07-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f0d07-855">- ActiveView</span></span><br><span data-ttu-id="f0d07-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-856">
         - CompressedFile</span></span><br><span data-ttu-id="f0d07-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-857">
         - DocumentEvents</span></span><br><span data-ttu-id="f0d07-858">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="f0d07-858">
         - File</span></span><br><span data-ttu-id="f0d07-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f0d07-859">
         - PdfFile</span></span><br><span data-ttu-id="f0d07-860">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-860">
         - Selection</span></span><br><span data-ttu-id="f0d07-861">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-861">
         - Settings</span></span><br><span data-ttu-id="f0d07-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f0d07-863">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="f0d07-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="f0d07-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="f0d07-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f0d07-865">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-865">Platform</span></span></th>
    <th><span data-ttu-id="f0d07-866">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-866">Extension points</span></span></th>
    <th><span data-ttu-id="f0d07-867">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="f0d07-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-869">Office na Web</span><span class="sxs-lookup"><span data-stu-id="f0d07-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="f0d07-870">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="f0d07-870">- Content</span></span><br><span data-ttu-id="f0d07-871">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-871">
         - TaskPane</span></span><br><span data-ttu-id="f0d07-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f0d07-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="f0d07-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f0d07-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f0d07-876">- DocumentEvents</span></span><br><span data-ttu-id="f0d07-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="f0d07-878">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="f0d07-878">
         - Settings</span></span><br><span data-ttu-id="f0d07-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="f0d07-880">Project</span><span class="sxs-lookup"><span data-stu-id="f0d07-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f0d07-881">Plataforma</span><span class="sxs-lookup"><span data-stu-id="f0d07-881">Platform</span></span></th>
    <th><span data-ttu-id="f0d07-882">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="f0d07-882">Extension points</span></span></th>
    <th><span data-ttu-id="f0d07-883">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="f0d07-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="f0d07-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-885">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-885">Office 2019 on Windows</span></span><br><span data-ttu-id="f0d07-886">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-887">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-889">- Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-889">- Selection</span></span><br><span data-ttu-id="f0d07-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-891">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-891">Office 2016 on Windows</span></span><br><span data-ttu-id="f0d07-892">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-893">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-895">- Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-895">- Selection</span></span><br><span data-ttu-id="f0d07-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f0d07-897">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="f0d07-897">Office 2013 on Windows</span></span><br><span data-ttu-id="f0d07-898">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="f0d07-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f0d07-899">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="f0d07-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f0d07-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f0d07-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f0d07-901">- Seleção</span><span class="sxs-lookup"><span data-stu-id="f0d07-901">- Selection</span></span><br><span data-ttu-id="f0d07-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f0d07-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f0d07-903">Confira também</span><span class="sxs-lookup"><span data-stu-id="f0d07-903">See also</span></span>

- [<span data-ttu-id="f0d07-904">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="f0d07-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f0d07-905">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="f0d07-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="f0d07-906">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="f0d07-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="f0d07-907">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="f0d07-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="f0d07-908">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="f0d07-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="f0d07-909">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="f0d07-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="f0d07-910">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="f0d07-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="f0d07-911">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="f0d07-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="f0d07-912">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="f0d07-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="f0d07-913">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="f0d07-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="f0d07-914">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="f0d07-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
