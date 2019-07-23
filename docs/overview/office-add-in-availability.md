---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/18/2019
localization_priority: Priority
ms.openlocfilehash: 510f2419d5d364a536f8c96f2057505161f03993
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804643"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="bc142-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bc142-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="bc142-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="bc142-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="bc142-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="bc142-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="bc142-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="bc142-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="bc142-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="bc142-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="bc142-108">Excel</span><span class="sxs-lookup"><span data-stu-id="bc142-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="bc142-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="bc142-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="bc142-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="bc142-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="bc142-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-114">- TaskPane</span></span><br><span data-ttu-id="bc142-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-115">
        - Content</span></span><br><span data-ttu-id="bc142-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-116">
        - Custom Functions</span></span><br><span data-ttu-id="bc142-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="bc142-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="bc142-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="bc142-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="bc142-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="bc142-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-130">
        - BindingEvents</span></span><br><span data-ttu-id="bc142-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-131">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-132">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-133">
        - File</span></span><br><span data-ttu-id="bc142-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-134">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-136">
        - Selection</span></span><br><span data-ttu-id="bc142-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-137">
        - Settings</span></span><br><span data-ttu-id="bc142-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-138">
        - TableBindings</span></span><br><span data-ttu-id="bc142-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-139">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-140">
        - TextBindings</span></span><br><span data-ttu-id="bc142-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-142">Office on Windows</span></span><br><span data-ttu-id="bc142-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-144">- TaskPane</span></span><br><span data-ttu-id="bc142-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-145">
        - Content</span></span><br><span data-ttu-id="bc142-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-146">
        - Custom Functions</span></span><br><span data-ttu-id="bc142-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="bc142-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="bc142-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="bc142-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="bc142-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="bc142-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-160">
        - BindingEvents</span></span><br><span data-ttu-id="bc142-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-161">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-162">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-163">
        - File</span></span><br><span data-ttu-id="bc142-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-164">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-166">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-166">
        - Selection</span></span><br><span data-ttu-id="bc142-167">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-167">
        - Settings</span></span><br><span data-ttu-id="bc142-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-168">
        - TableBindings</span></span><br><span data-ttu-id="bc142-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-169">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-170">
        - TextBindings</span></span><br><span data-ttu-id="bc142-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-172">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-172">Office 2019 on Windows</span></span><br><span data-ttu-id="bc142-173">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="bc142-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-174">- TaskPane</span></span><br><span data-ttu-id="bc142-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-175">
        - Content</span></span><br><span data-ttu-id="bc142-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc142-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-187">- BindingEvents</span></span><br><span data-ttu-id="bc142-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-188">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-189">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-190">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-190">
        - File</span></span><br><span data-ttu-id="bc142-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-191">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-193">
        - Selection</span></span><br><span data-ttu-id="bc142-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-194">
        - Settings</span></span><br><span data-ttu-id="bc142-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-195">
        - TableBindings</span></span><br><span data-ttu-id="bc142-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-196">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-197">
        - TextBindings</span></span><br><span data-ttu-id="bc142-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-199">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-199">Office 2016 on Windows</span></span><br><span data-ttu-id="bc142-200">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="bc142-201">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-201">- TaskPane</span></span><br><span data-ttu-id="bc142-202">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-202">
        - Content</span></span></td>
    <td><span data-ttu-id="bc142-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="bc142-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-206">- BindingEvents</span></span><br><span data-ttu-id="bc142-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-207">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-208">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-209">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-209">
        - File</span></span><br><span data-ttu-id="bc142-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-210">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-212">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-212">
        - Selection</span></span><br><span data-ttu-id="bc142-213">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-213">
        - Settings</span></span><br><span data-ttu-id="bc142-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-214">
        - TableBindings</span></span><br><span data-ttu-id="bc142-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-215">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-216">
        - TextBindings</span></span><br><span data-ttu-id="bc142-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-218">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-218">Office 2013 on Windows</span></span><br><span data-ttu-id="bc142-219">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="bc142-220">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-220">
        - TaskPane</span></span><br><span data-ttu-id="bc142-221">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="bc142-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="bc142-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="bc142-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-224">
        - BindingEvents</span></span><br><span data-ttu-id="bc142-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-225">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-226">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-227">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-227">
        - File</span></span><br><span data-ttu-id="bc142-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-228">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-230">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-230">
        - Selection</span></span><br><span data-ttu-id="bc142-231">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-231">
        - Settings</span></span><br><span data-ttu-id="bc142-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-232">
        - TableBindings</span></span><br><span data-ttu-id="bc142-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-233">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-234">
        - TextBindings</span></span><br><span data-ttu-id="bc142-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-236">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="bc142-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="bc142-237">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="bc142-238">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-238">- TaskPane</span></span><br><span data-ttu-id="bc142-239">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-239">
        - Content</span></span><br><span data-ttu-id="bc142-240">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="bc142-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="bc142-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="bc142-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-252">- BindingEvents</span></span><br><span data-ttu-id="bc142-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-253">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-254">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-254">
        - File</span></span><br><span data-ttu-id="bc142-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-255">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-257">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-257">
        - Selection</span></span><br><span data-ttu-id="bc142-258">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-258">
        - Settings</span></span><br><span data-ttu-id="bc142-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-259">
        - TableBindings</span></span><br><span data-ttu-id="bc142-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-260">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-261">
        - TextBindings</span></span><br><span data-ttu-id="bc142-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-263">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-263">Office apps on Mac</span></span><br><span data-ttu-id="bc142-264">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="bc142-265">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-265">- TaskPane</span></span><br><span data-ttu-id="bc142-266">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-266">
        - Content</span></span><br><span data-ttu-id="bc142-267">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-267">
        - Custom Functions</span></span><br><span data-ttu-id="bc142-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc142-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="bc142-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="bc142-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="bc142-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-281">- BindingEvents</span></span><br><span data-ttu-id="bc142-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-282">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-283">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-284">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-284">
        - File</span></span><br><span data-ttu-id="bc142-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-285">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-287">
        - PdfFile</span></span><br><span data-ttu-id="bc142-288">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-288">
        - Selection</span></span><br><span data-ttu-id="bc142-289">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-289">
        - Settings</span></span><br><span data-ttu-id="bc142-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-290">
        - TableBindings</span></span><br><span data-ttu-id="bc142-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-291">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-292">
        - TextBindings</span></span><br><span data-ttu-id="bc142-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-294">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-294">Office 2019 for Mac</span></span><br><span data-ttu-id="bc142-295">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="bc142-296">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-296">- TaskPane</span></span><br><span data-ttu-id="bc142-297">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-297">
        - Content</span></span><br><span data-ttu-id="bc142-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bc142-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bc142-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bc142-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bc142-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bc142-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bc142-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bc142-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bc142-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bc142-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-309">- BindingEvents</span></span><br><span data-ttu-id="bc142-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-310">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-311">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-312">
        - File</span></span><br><span data-ttu-id="bc142-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-313">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-315">
        - PdfFile</span></span><br><span data-ttu-id="bc142-316">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-316">
        - Selection</span></span><br><span data-ttu-id="bc142-317">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-317">
        - Settings</span></span><br><span data-ttu-id="bc142-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-318">
        - TableBindings</span></span><br><span data-ttu-id="bc142-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-319">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-320">
        - TextBindings</span></span><br><span data-ttu-id="bc142-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-322">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-322">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="bc142-323">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="bc142-324">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-324">- TaskPane</span></span><br><span data-ttu-id="bc142-325">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-325">
        - Content</span></span></td>
    <td><span data-ttu-id="bc142-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bc142-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="bc142-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="bc142-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-329">- BindingEvents</span></span><br><span data-ttu-id="bc142-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-330">
        - CompressedFile</span></span><br><span data-ttu-id="bc142-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-331">
        - DocumentEvents</span></span><br><span data-ttu-id="bc142-332">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-332">
        - File</span></span><br><span data-ttu-id="bc142-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-333">
        - MatrixBindings</span></span><br><span data-ttu-id="bc142-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="bc142-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-335">
        - PdfFile</span></span><br><span data-ttu-id="bc142-336">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-336">
        - Selection</span></span><br><span data-ttu-id="bc142-337">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-337">
        - Settings</span></span><br><span data-ttu-id="bc142-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-338">
        - TableBindings</span></span><br><span data-ttu-id="bc142-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-339">
        - TableCoercion</span></span><br><span data-ttu-id="bc142-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-340">
        - TextBindings</span></span><br><span data-ttu-id="bc142-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="bc142-342">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="bc142-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="bc142-343">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="bc142-344">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="bc142-345">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="bc142-346">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="bc142-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-348">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-348">Office on the web</span></span></td>
    <td><span data-ttu-id="bc142-349">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="bc142-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-351">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-351">Office on Windows</span></span><br><span data-ttu-id="bc142-352">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="bc142-353">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="bc142-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-355">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-355">Office for Mac</span></span><br><span data-ttu-id="bc142-356">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="bc142-357">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="bc142-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="bc142-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="bc142-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="bc142-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc142-360">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-360">Platform</span></span></th>
    <th><span data-ttu-id="bc142-361">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-361">Extension points</span></span></th>
    <th><span data-ttu-id="bc142-362">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc142-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-364">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-364">Office on the web</span></span><br><span data-ttu-id="bc142-365">(moderno)</span><span class="sxs-lookup"><span data-stu-id="bc142-365">Modern</span></span></td>
    <td> <span data-ttu-id="bc142-366">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-366">- Mail Read</span></span><br><span data-ttu-id="bc142-367">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-367">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc142-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc142-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-377">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-377">Office on the web</span></span><br><span data-ttu-id="bc142-378">(clássico)</span><span class="sxs-lookup"><span data-stu-id="bc142-378">Classic.</span></span></td>
    <td> <span data-ttu-id="bc142-379">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-379">- Mail Read</span></span><br><span data-ttu-id="bc142-380">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-380">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bc142-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-389">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-389">Office on Windows</span></span><br><span data-ttu-id="bc142-390">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-391">- Mail Read</span></span><br><span data-ttu-id="bc142-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-392">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bc142-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="bc142-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bc142-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc142-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc142-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-403">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-403">Office 2019 on Windows</span></span><br><span data-ttu-id="bc142-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-405">- Mail Read</span></span><br><span data-ttu-id="bc142-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-406">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bc142-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="bc142-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bc142-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc142-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc142-416">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-417">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-417">Office 2016 on Windows</span></span><br><span data-ttu-id="bc142-418">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-419">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-419">- Mail Read</span></span><br><span data-ttu-id="bc142-420">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-420">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bc142-422">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="bc142-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bc142-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="bc142-427">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-428">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-428">Office 2013 on Windows</span></span><br><span data-ttu-id="bc142-429">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-430">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-430">- Mail Read</span></span><br><span data-ttu-id="bc142-431">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="bc142-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="bc142-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="bc142-436">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-437">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="bc142-437">Office apps on iOS</span></span><br><span data-ttu-id="bc142-438">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-439">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-439">- Mail Read</span></span><br><span data-ttu-id="bc142-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bc142-446">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-447">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-447">Office apps on Mac</span></span><br><span data-ttu-id="bc142-448">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-449">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-449">- Mail Read</span></span><br><span data-ttu-id="bc142-450">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-450">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bc142-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bc142-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bc142-459">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-460">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-460">Office 2019 for Mac</span></span><br><span data-ttu-id="bc142-461">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-462">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-462">- Mail Read</span></span><br><span data-ttu-id="bc142-463">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-463">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bc142-471">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-472">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-472">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="bc142-473">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-474">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-474">- Mail Read</span></span><br><span data-ttu-id="bc142-475">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="bc142-475">
      - Mail Compose</span></span><br><span data-ttu-id="bc142-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bc142-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bc142-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bc142-483">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-484">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="bc142-484">Office apps on Android</span></span><br><span data-ttu-id="bc142-485">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-486">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="bc142-486">- Mail Read</span></span><br><span data-ttu-id="bc142-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bc142-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bc142-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bc142-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bc142-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bc142-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bc142-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bc142-493">Não disponível</span><span class="sxs-lookup"><span data-stu-id="bc142-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="bc142-494">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="bc142-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="bc142-495">Word</span><span class="sxs-lookup"><span data-stu-id="bc142-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc142-496">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-496">Platform</span></span></th>
    <th><span data-ttu-id="bc142-497">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-497">Extension points</span></span></th>
    <th><span data-ttu-id="bc142-498">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc142-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-500">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="bc142-501">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-501">- TaskPane</span></span><br><span data-ttu-id="bc142-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="bc142-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-509">- BindingEvents</span></span><br><span data-ttu-id="bc142-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-511">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-512">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-512">
         - File</span></span><br><span data-ttu-id="bc142-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-514">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-517">
         - PdfFile</span></span><br><span data-ttu-id="bc142-518">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-518">
         - Selection</span></span><br><span data-ttu-id="bc142-519">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-519">
         - Settings</span></span><br><span data-ttu-id="bc142-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-520">
         - TableBindings</span></span><br><span data-ttu-id="bc142-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-521">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-522">
         - TextBindings</span></span><br><span data-ttu-id="bc142-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-523">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-525">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-525">Office on Windows</span></span><br><span data-ttu-id="bc142-526">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-527">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-527">- TaskPane</span></span><br><span data-ttu-id="bc142-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="bc142-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-535">- BindingEvents</span></span><br><span data-ttu-id="bc142-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-536">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-538">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-539">
         - File</span></span><br><span data-ttu-id="bc142-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-541">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-544">
         - PdfFile</span></span><br><span data-ttu-id="bc142-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-545">
         - Selection</span></span><br><span data-ttu-id="bc142-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-546">
         - Settings</span></span><br><span data-ttu-id="bc142-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-547">
         - TableBindings</span></span><br><span data-ttu-id="bc142-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-548">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-549">
         - TextBindings</span></span><br><span data-ttu-id="bc142-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-550">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-552">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-552">Office 2019 on Windows</span></span><br><span data-ttu-id="bc142-553">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-554">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-554">- TaskPane</span></span><br><span data-ttu-id="bc142-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-561">- BindingEvents</span></span><br><span data-ttu-id="bc142-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-562">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-564">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-565">
         - File</span></span><br><span data-ttu-id="bc142-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-567">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-570">
         - PdfFile</span></span><br><span data-ttu-id="bc142-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-571">
         - Selection</span></span><br><span data-ttu-id="bc142-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-572">
         - Settings</span></span><br><span data-ttu-id="bc142-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-573">
         - TableBindings</span></span><br><span data-ttu-id="bc142-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-574">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-575">
         - TextBindings</span></span><br><span data-ttu-id="bc142-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-576">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-578">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-578">Office 2016 on Windows</span></span><br><span data-ttu-id="bc142-579">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-580">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="bc142-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-584">- BindingEvents</span></span><br><span data-ttu-id="bc142-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-585">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-587">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-588">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-588">
         - File</span></span><br><span data-ttu-id="bc142-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-590">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-593">
         - PdfFile</span></span><br><span data-ttu-id="bc142-594">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-594">
         - Selection</span></span><br><span data-ttu-id="bc142-595">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-595">
         - Settings</span></span><br><span data-ttu-id="bc142-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-596">
         - TableBindings</span></span><br><span data-ttu-id="bc142-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-597">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-598">
         - TextBindings</span></span><br><span data-ttu-id="bc142-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-599">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-601">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-601">Office 2013 on Windows</span></span><br><span data-ttu-id="bc142-602">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-603">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="bc142-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="bc142-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-606">- BindingEvents</span></span><br><span data-ttu-id="bc142-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-607">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-609">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-610">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-610">
         - File</span></span><br><span data-ttu-id="bc142-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-612">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-615">
         - PdfFile</span></span><br><span data-ttu-id="bc142-616">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-616">
         - Selection</span></span><br><span data-ttu-id="bc142-617">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-617">
         - Settings</span></span><br><span data-ttu-id="bc142-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-618">
         - TableBindings</span></span><br><span data-ttu-id="bc142-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-619">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-620">
         - TextBindings</span></span><br><span data-ttu-id="bc142-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-621">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-623">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="bc142-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="bc142-624">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-625">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="bc142-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-631">- BindingEvents</span></span><br><span data-ttu-id="bc142-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-632">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-634">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-635">
         - File</span></span><br><span data-ttu-id="bc142-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-637">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-640">
         - PdfFile</span></span><br><span data-ttu-id="bc142-641">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-641">
         - Selection</span></span><br><span data-ttu-id="bc142-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-642">
         - Settings</span></span><br><span data-ttu-id="bc142-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-643">
         - TableBindings</span></span><br><span data-ttu-id="bc142-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-644">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-645">
         - TextBindings</span></span><br><span data-ttu-id="bc142-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-646">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-648">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-648">Office apps on Mac</span></span><br><span data-ttu-id="bc142-649">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-650">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-650">- TaskPane</span></span><br><span data-ttu-id="bc142-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="bc142-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-658">- BindingEvents</span></span><br><span data-ttu-id="bc142-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-659">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-661">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-662">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-662">
         - File</span></span><br><span data-ttu-id="bc142-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-664">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-667">
         - PdfFile</span></span><br><span data-ttu-id="bc142-668">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-668">
         - Selection</span></span><br><span data-ttu-id="bc142-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-669">
         - Settings</span></span><br><span data-ttu-id="bc142-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-670">
         - TableBindings</span></span><br><span data-ttu-id="bc142-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-671">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-672">
         - TextBindings</span></span><br><span data-ttu-id="bc142-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-673">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-675">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-675">Office 2019 for Mac</span></span><br><span data-ttu-id="bc142-676">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-677">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-677">- TaskPane</span></span><br><span data-ttu-id="bc142-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="bc142-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bc142-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="bc142-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="bc142-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-684">- BindingEvents</span></span><br><span data-ttu-id="bc142-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-685">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-687">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-688">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-688">
         - File</span></span><br><span data-ttu-id="bc142-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-690">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-693">
         - PdfFile</span></span><br><span data-ttu-id="bc142-694">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-694">
         - Selection</span></span><br><span data-ttu-id="bc142-695">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-695">
         - Settings</span></span><br><span data-ttu-id="bc142-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-696">
         - TableBindings</span></span><br><span data-ttu-id="bc142-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-697">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-698">
         - TextBindings</span></span><br><span data-ttu-id="bc142-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-699">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-701">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-701">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="bc142-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="bc142-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="bc142-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="bc142-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-707">- BindingEvents</span></span><br><span data-ttu-id="bc142-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-708">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bc142-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="bc142-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-710">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-711">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-711">
         - File</span></span><br><span data-ttu-id="bc142-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-713">
         - MatrixBindings</span></span><br><span data-ttu-id="bc142-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="bc142-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bc142-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-716">
         - PdfFile</span></span><br><span data-ttu-id="bc142-717">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-717">
         - Selection</span></span><br><span data-ttu-id="bc142-718">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-718">
         - Settings</span></span><br><span data-ttu-id="bc142-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-719">
         - TableBindings</span></span><br><span data-ttu-id="bc142-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-720">
         - TableCoercion</span></span><br><span data-ttu-id="bc142-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bc142-721">
         - TextBindings</span></span><br><span data-ttu-id="bc142-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-722">
         - TextCoercion</span></span><br><span data-ttu-id="bc142-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bc142-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="bc142-724">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="bc142-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="bc142-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bc142-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc142-726">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-726">Platform</span></span></th>
    <th><span data-ttu-id="bc142-727">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-727">Extension points</span></span></th>
    <th><span data-ttu-id="bc142-728">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc142-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-730">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="bc142-731">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-731">- Content</span></span><br><span data-ttu-id="bc142-732">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-732">
         - TaskPane</span></span><br><span data-ttu-id="bc142-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="bc142-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-737">- ActiveView</span></span><br><span data-ttu-id="bc142-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-738">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-739">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-740">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-740">
         - File</span></span><br><span data-ttu-id="bc142-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-741">
         - PdfFile</span></span><br><span data-ttu-id="bc142-742">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-742">
         - Selection</span></span><br><span data-ttu-id="bc142-743">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-743">
         - Settings</span></span><br><span data-ttu-id="bc142-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-745">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-745">Office on Windows</span></span><br><span data-ttu-id="bc142-746">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-747">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-747">- Content</span></span><br><span data-ttu-id="bc142-748">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-748">
         - TaskPane</span></span><br><span data-ttu-id="bc142-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="bc142-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-753">- ActiveView</span></span><br><span data-ttu-id="bc142-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-754">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-755">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-756">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-756">
         - File</span></span><br><span data-ttu-id="bc142-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-757">
         - PdfFile</span></span><br><span data-ttu-id="bc142-758">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-758">
         - Selection</span></span><br><span data-ttu-id="bc142-759">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-759">
         - Settings</span></span><br><span data-ttu-id="bc142-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-761">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-761">Office 2019 on Windows</span></span><br><span data-ttu-id="bc142-762">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-763">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-763">- Content</span></span><br><span data-ttu-id="bc142-764">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-764">
         - TaskPane</span></span><br><span data-ttu-id="bc142-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-768">- ActiveView</span></span><br><span data-ttu-id="bc142-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-769">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-770">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-771">
         - File</span></span><br><span data-ttu-id="bc142-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-772">
         - PdfFile</span></span><br><span data-ttu-id="bc142-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-773">
         - Selection</span></span><br><span data-ttu-id="bc142-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-774">
         - Settings</span></span><br><span data-ttu-id="bc142-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-776">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-776">Office 2016 on Windows</span></span><br><span data-ttu-id="bc142-777">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-778">- Content</span></span><br><span data-ttu-id="bc142-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="bc142-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="bc142-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-782">- ActiveView</span></span><br><span data-ttu-id="bc142-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-783">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-784">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-785">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-785">
         - File</span></span><br><span data-ttu-id="bc142-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-786">
         - PdfFile</span></span><br><span data-ttu-id="bc142-787">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-787">
         - Selection</span></span><br><span data-ttu-id="bc142-788">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-788">
         - Settings</span></span><br><span data-ttu-id="bc142-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-790">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-790">Office 2013 on Windows</span></span><br><span data-ttu-id="bc142-791">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-792">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-792">- Content</span></span><br><span data-ttu-id="bc142-793">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="bc142-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="bc142-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="bc142-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-796">- ActiveView</span></span><br><span data-ttu-id="bc142-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-797">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-798">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-799">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-799">
         - File</span></span><br><span data-ttu-id="bc142-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-800">
         - PdfFile</span></span><br><span data-ttu-id="bc142-801">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-801">
         - Selection</span></span><br><span data-ttu-id="bc142-802">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-802">
         - Settings</span></span><br><span data-ttu-id="bc142-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-804">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="bc142-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="bc142-805">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-806">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-806">- Content</span></span><br><span data-ttu-id="bc142-807">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-810">- ActiveView</span></span><br><span data-ttu-id="bc142-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-811">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-812">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-813">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-813">
         - File</span></span><br><span data-ttu-id="bc142-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-814">
         - PdfFile</span></span><br><span data-ttu-id="bc142-815">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-815">
         - Selection</span></span><br><span data-ttu-id="bc142-816">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-816">
         - Settings</span></span><br><span data-ttu-id="bc142-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-818">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-818">Office apps on Mac</span></span><br><span data-ttu-id="bc142-819">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="bc142-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="bc142-820">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-820">- Content</span></span><br><span data-ttu-id="bc142-821">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-821">
         - TaskPane</span></span><br><span data-ttu-id="bc142-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="bc142-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bc142-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="bc142-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-826">- ActiveView</span></span><br><span data-ttu-id="bc142-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-827">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-828">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-829">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-829">
         - File</span></span><br><span data-ttu-id="bc142-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-830">
         - PdfFile</span></span><br><span data-ttu-id="bc142-831">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-831">
         - Selection</span></span><br><span data-ttu-id="bc142-832">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-832">
         - Settings</span></span><br><span data-ttu-id="bc142-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-834">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-834">Office 2019 for Mac</span></span><br><span data-ttu-id="bc142-835">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-836">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-836">- Content</span></span><br><span data-ttu-id="bc142-837">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-837">
         - TaskPane</span></span><br><span data-ttu-id="bc142-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-841">- ActiveView</span></span><br><span data-ttu-id="bc142-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-842">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-843">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-844">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-844">
         - File</span></span><br><span data-ttu-id="bc142-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-845">
         - PdfFile</span></span><br><span data-ttu-id="bc142-846">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-846">
         - Selection</span></span><br><span data-ttu-id="bc142-847">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-847">
         - Settings</span></span><br><span data-ttu-id="bc142-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-849">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-849">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="bc142-850">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-851">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-851">- Content</span></span><br><span data-ttu-id="bc142-852">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="bc142-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="bc142-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bc142-855">- ActiveView</span></span><br><span data-ttu-id="bc142-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bc142-856">
         - CompressedFile</span></span><br><span data-ttu-id="bc142-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-857">
         - DocumentEvents</span></span><br><span data-ttu-id="bc142-858">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="bc142-858">
         - File</span></span><br><span data-ttu-id="bc142-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bc142-859">
         - PdfFile</span></span><br><span data-ttu-id="bc142-860">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-860">
         - Selection</span></span><br><span data-ttu-id="bc142-861">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-861">
         - Settings</span></span><br><span data-ttu-id="bc142-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="bc142-863">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="bc142-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="bc142-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="bc142-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc142-865">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-865">Platform</span></span></th>
    <th><span data-ttu-id="bc142-866">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-866">Extension points</span></span></th>
    <th><span data-ttu-id="bc142-867">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc142-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-869">Office na Web</span><span class="sxs-lookup"><span data-stu-id="bc142-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="bc142-870">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="bc142-870">- Content</span></span><br><span data-ttu-id="bc142-871">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-871">
         - TaskPane</span></span><br><span data-ttu-id="bc142-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="bc142-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bc142-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="bc142-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="bc142-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bc142-876">- DocumentEvents</span></span><br><span data-ttu-id="bc142-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="bc142-878">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="bc142-878">
         - Settings</span></span><br><span data-ttu-id="bc142-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="bc142-880">Project</span><span class="sxs-lookup"><span data-stu-id="bc142-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bc142-881">Plataforma</span><span class="sxs-lookup"><span data-stu-id="bc142-881">Platform</span></span></th>
    <th><span data-ttu-id="bc142-882">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="bc142-882">Extension points</span></span></th>
    <th><span data-ttu-id="bc142-883">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="bc142-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="bc142-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="bc142-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-885">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-885">Office 2019 on Windows</span></span><br><span data-ttu-id="bc142-886">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-887">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-889">- Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-889">- Selection</span></span><br><span data-ttu-id="bc142-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-891">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-891">Office 2016 on Windows</span></span><br><span data-ttu-id="bc142-892">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-893">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-895">- Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-895">- Selection</span></span><br><span data-ttu-id="bc142-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bc142-897">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="bc142-897">Office 2013 on Windows</span></span><br><span data-ttu-id="bc142-898">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="bc142-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="bc142-899">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="bc142-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bc142-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bc142-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bc142-901">- Seleção</span><span class="sxs-lookup"><span data-stu-id="bc142-901">- Selection</span></span><br><span data-ttu-id="bc142-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bc142-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="bc142-903">Confira também</span><span class="sxs-lookup"><span data-stu-id="bc142-903">See also</span></span>

- [<span data-ttu-id="bc142-904">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="bc142-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="bc142-905">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="bc142-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="bc142-906">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="bc142-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="bc142-907">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="bc142-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="bc142-908">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="bc142-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="bc142-909">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="bc142-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="bc142-910">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="bc142-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="bc142-911">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="bc142-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="bc142-912">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="bc142-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="bc142-913">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="bc142-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="bc142-914">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="bc142-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
