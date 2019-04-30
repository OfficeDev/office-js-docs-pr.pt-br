---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448144"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="6a194-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6a194-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="6a194-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="6a194-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="6a194-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="6a194-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="6a194-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="6a194-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="6a194-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="6a194-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="6a194-108">Excel</span><span class="sxs-lookup"><span data-stu-id="6a194-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="6a194-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="6a194-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="6a194-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="6a194-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="6a194-113">Office Online</span></span></td>
    <td> <span data-ttu-id="6a194-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-114">- TaskPane</span></span><br><span data-ttu-id="6a194-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-115">
        - Content</span></span><br><span data-ttu-id="6a194-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6a194-116">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6a194-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-126">
        - BindingEvents</span></span><br><span data-ttu-id="6a194-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-127">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-128">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-129">
        - File</span></span><br><span data-ttu-id="6a194-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-130">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-132">
        - Selection</span></span><br><span data-ttu-id="6a194-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-133">
        - Settings</span></span><br><span data-ttu-id="6a194-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-134">
        - TableBindings</span></span><br><span data-ttu-id="6a194-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-135">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-136">
        - TextBindings</span></span><br><span data-ttu-id="6a194-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-138">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-138">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-139">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-139">- TaskPane</span></span><br><span data-ttu-id="6a194-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-140">
        - Content</span></span><br><span data-ttu-id="6a194-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="6a194-141">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="6a194-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-151">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-151">
        - BindingEvents</span></span><br><span data-ttu-id="6a194-152">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-152">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-153">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-153">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-154">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-154">
        - File</span></span><br><span data-ttu-id="6a194-155">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-155">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-156">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-156">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-157">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-157">
        - Selection</span></span><br><span data-ttu-id="6a194-158">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-158">
        - Settings</span></span><br><span data-ttu-id="6a194-159">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-159">
        - TableBindings</span></span><br><span data-ttu-id="6a194-160">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-160">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-161">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-161">
        - TextBindings</span></span><br><span data-ttu-id="6a194-162">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-162">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-163">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-163">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="6a194-164">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-164">- TaskPane</span></span><br><span data-ttu-id="6a194-165">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-165">
        - Content</span></span><br><span data-ttu-id="6a194-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-166">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6a194-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-167">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-168">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-176">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-176">- BindingEvents</span></span><br><span data-ttu-id="6a194-177">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-177">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-178">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-178">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-179">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-179">
        - File</span></span><br><span data-ttu-id="6a194-180">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-180">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-181">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-181">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-182">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-182">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-183">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-183">
        - Selection</span></span><br><span data-ttu-id="6a194-184">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-184">
        - Settings</span></span><br><span data-ttu-id="6a194-185">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-185">
        - TableBindings</span></span><br><span data-ttu-id="6a194-186">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-186">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-187">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-187">
        - TextBindings</span></span><br><span data-ttu-id="6a194-188">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-188">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-189">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-189">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="6a194-190">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-190">- TaskPane</span></span><br><span data-ttu-id="6a194-191">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-191">
        - Content</span></span></td>
    <td><span data-ttu-id="6a194-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-192">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-193">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="6a194-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-194">- BindingEvents</span></span><br><span data-ttu-id="6a194-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-195">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-196">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-197">
        - File</span></span><br><span data-ttu-id="6a194-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-198">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-199">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-201">
        - Selection</span></span><br><span data-ttu-id="6a194-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-202">
        - Settings</span></span><br><span data-ttu-id="6a194-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-203">
        - TableBindings</span></span><br><span data-ttu-id="6a194-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-204">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-205">
        - TextBindings</span></span><br><span data-ttu-id="6a194-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-207">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-207">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="6a194-208">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-208">
        - TaskPane</span></span><br><span data-ttu-id="6a194-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-209">
        - Content</span></span></td>
    <td>  <span data-ttu-id="6a194-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6a194-210">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="6a194-211">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-211">
        - BindingEvents</span></span><br><span data-ttu-id="6a194-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-212">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-213">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-214">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-214">
        - File</span></span><br><span data-ttu-id="6a194-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-215">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-216">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-218">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-218">
        - Selection</span></span><br><span data-ttu-id="6a194-219">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-219">
        - Settings</span></span><br><span data-ttu-id="6a194-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-220">
        - TableBindings</span></span><br><span data-ttu-id="6a194-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-221">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-222">
        - TextBindings</span></span><br><span data-ttu-id="6a194-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-224">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="6a194-224">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="6a194-225">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-225">- TaskPane</span></span><br><span data-ttu-id="6a194-226">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-226">
        - Content</span></span></td>
    <td><span data-ttu-id="6a194-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-227">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-228">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-234">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-235">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-236">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-236">- BindingEvents</span></span><br><span data-ttu-id="6a194-237">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-237">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-238">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-238">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-239">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-239">
        - File</span></span><br><span data-ttu-id="6a194-240">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-240">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-241">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-241">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-242">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-242">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-243">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-243">
        - Selection</span></span><br><span data-ttu-id="6a194-244">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-244">
        - Settings</span></span><br><span data-ttu-id="6a194-245">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-245">
        - TableBindings</span></span><br><span data-ttu-id="6a194-246">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-246">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-247">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-247">
        - TextBindings</span></span><br><span data-ttu-id="6a194-248">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-248">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-249">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-249">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="6a194-250">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-250">- TaskPane</span></span><br><span data-ttu-id="6a194-251">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-251">
        - Content</span></span><br><span data-ttu-id="6a194-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6a194-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-253">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-262">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-262">- BindingEvents</span></span><br><span data-ttu-id="6a194-263">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-263">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-264">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-264">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-265">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-265">
        - File</span></span><br><span data-ttu-id="6a194-266">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-266">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-267">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-267">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-268">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-268">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-269">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-269">
        - PdfFile</span></span><br><span data-ttu-id="6a194-270">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-270">
        - Selection</span></span><br><span data-ttu-id="6a194-271">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-271">
        - Settings</span></span><br><span data-ttu-id="6a194-272">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-272">
        - TableBindings</span></span><br><span data-ttu-id="6a194-273">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-273">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-274">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-274">
        - TextBindings</span></span><br><span data-ttu-id="6a194-275">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-275">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-276">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-276">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="6a194-277">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-277">- TaskPane</span></span><br><span data-ttu-id="6a194-278">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-278">
        - Content</span></span><br><span data-ttu-id="6a194-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="6a194-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-280">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="6a194-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="6a194-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="6a194-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="6a194-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="6a194-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="6a194-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="6a194-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="6a194-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="6a194-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-289">- BindingEvents</span></span><br><span data-ttu-id="6a194-290">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-290">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-291">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-291">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-292">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-292">
        - File</span></span><br><span data-ttu-id="6a194-293">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-293">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-294">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-294">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-295">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-295">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-296">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-296">
        - PdfFile</span></span><br><span data-ttu-id="6a194-297">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-297">
        - Selection</span></span><br><span data-ttu-id="6a194-298">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-298">
        - Settings</span></span><br><span data-ttu-id="6a194-299">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-299">
        - TableBindings</span></span><br><span data-ttu-id="6a194-300">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-300">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-301">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-301">
        - TextBindings</span></span><br><span data-ttu-id="6a194-302">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-302">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-303">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-303">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="6a194-304">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-304">- TaskPane</span></span><br><span data-ttu-id="6a194-305">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-305">
        - Content</span></span></td>
    <td><span data-ttu-id="6a194-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-306">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="6a194-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="6a194-308">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-308">- BindingEvents</span></span><br><span data-ttu-id="6a194-309">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-309">
        - CompressedFile</span></span><br><span data-ttu-id="6a194-310">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-310">
        - DocumentEvents</span></span><br><span data-ttu-id="6a194-311">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-311">
        - File</span></span><br><span data-ttu-id="6a194-312">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-312">
        - ImageCoercion</span></span><br><span data-ttu-id="6a194-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-313">
        - MatrixBindings</span></span><br><span data-ttu-id="6a194-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="6a194-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-315">
        - PdfFile</span></span><br><span data-ttu-id="6a194-316">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-316">
        - Selection</span></span><br><span data-ttu-id="6a194-317">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-317">
        - Settings</span></span><br><span data-ttu-id="6a194-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-318">
        - TableBindings</span></span><br><span data-ttu-id="6a194-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-319">
        - TableCoercion</span></span><br><span data-ttu-id="6a194-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-320">
        - TextBindings</span></span><br><span data-ttu-id="6a194-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-321">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6a194-322">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6a194-322">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="6a194-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="6a194-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6a194-324">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-324">Platform</span></span></th>
    <th><span data-ttu-id="6a194-325">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-325">Extension points</span></span></th>
    <th><span data-ttu-id="6a194-326">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-326">API requirement sets</span></span></th>
    <th><span data-ttu-id="6a194-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-327"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="6a194-328">Office Online</span></span></td>
    <td> <span data-ttu-id="6a194-329">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-329">- Mail Read</span></span><br><span data-ttu-id="6a194-330">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-330">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-332">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-333">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6a194-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6a194-339">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-340">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-340">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-341">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-341">- Mail Read</span></span><br><span data-ttu-id="6a194-342">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-342">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6a194-344">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="6a194-344">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6a194-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-345">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-346">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6a194-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6a194-352">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-353">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-353">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-354">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-354">- Mail Read</span></span><br><span data-ttu-id="6a194-355">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-355">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-356">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6a194-357">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="6a194-357">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6a194-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-358">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-359">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="6a194-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="6a194-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="6a194-365">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-366">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-366">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-367">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-367">- Mail Read</span></span><br><span data-ttu-id="6a194-368">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-368">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-369">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="6a194-370">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="6a194-370">
      - Modules</span></span></td>
    <td> <span data-ttu-id="6a194-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6a194-375">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-376">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-376">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-377">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-377">- Mail Read</span></span><br><span data-ttu-id="6a194-378">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-378">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="6a194-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="6a194-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="6a194-383">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-384">Office 365 para iOS</span><span class="sxs-lookup"><span data-stu-id="6a194-384">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="6a194-385">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-385">- Mail Read</span></span><br><span data-ttu-id="6a194-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-386">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-387">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6a194-392">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-393">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-393">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-394">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-394">- Mail Read</span></span><br><span data-ttu-id="6a194-395">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-395">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-397">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6a194-403">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-404">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-404">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-405">- Mail Read</span></span><br><span data-ttu-id="6a194-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-406">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6a194-414">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-415">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-415">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-416">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-416">- Mail Read</span></span><br><span data-ttu-id="6a194-417">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="6a194-417">
      - Mail Compose</span></span><br><span data-ttu-id="6a194-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-419">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="6a194-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="6a194-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="6a194-425">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-426">Office 365 para Android</span><span class="sxs-lookup"><span data-stu-id="6a194-426">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="6a194-427">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="6a194-427">- Mail Read</span></span><br><span data-ttu-id="6a194-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-428">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="6a194-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="6a194-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="6a194-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="6a194-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="6a194-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="6a194-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="6a194-434">Não disponível</span><span class="sxs-lookup"><span data-stu-id="6a194-434">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="6a194-435">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6a194-435">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="6a194-436">Word</span><span class="sxs-lookup"><span data-stu-id="6a194-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6a194-437">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-437">Platform</span></span></th>
    <th><span data-ttu-id="6a194-438">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-438">Extension points</span></span></th>
    <th><span data-ttu-id="6a194-439">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="6a194-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-440"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="6a194-441">Office Online</span></span></td>
    <td> <span data-ttu-id="6a194-442">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-442">- TaskPane</span></span><br><span data-ttu-id="6a194-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-443">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-444">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-445">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-448">- BindingEvents</span></span><br><span data-ttu-id="6a194-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-450">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-451">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-451">
         - File</span></span><br><span data-ttu-id="6a194-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-453">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-454">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-457">
         - PdfFile</span></span><br><span data-ttu-id="6a194-458">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-458">
         - Selection</span></span><br><span data-ttu-id="6a194-459">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-459">
         - Settings</span></span><br><span data-ttu-id="6a194-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-460">
         - TableBindings</span></span><br><span data-ttu-id="6a194-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-461">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-462">
         - TextBindings</span></span><br><span data-ttu-id="6a194-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-463">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-465">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-466">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-466">- TaskPane</span></span><br><span data-ttu-id="6a194-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-467">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-468">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-469">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-472">- BindingEvents</span></span><br><span data-ttu-id="6a194-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-473">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-475">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-476">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-476">
         - File</span></span><br><span data-ttu-id="6a194-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-478">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-479">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-482">
         - PdfFile</span></span><br><span data-ttu-id="6a194-483">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-483">
         - Selection</span></span><br><span data-ttu-id="6a194-484">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-484">
         - Settings</span></span><br><span data-ttu-id="6a194-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-485">
         - TableBindings</span></span><br><span data-ttu-id="6a194-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-486">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-487">
         - TextBindings</span></span><br><span data-ttu-id="6a194-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-488">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-490">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-491">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-491">- TaskPane</span></span><br><span data-ttu-id="6a194-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-493">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-494">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-497">- BindingEvents</span></span><br><span data-ttu-id="6a194-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-498">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-500">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-501">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-501">
         - File</span></span><br><span data-ttu-id="6a194-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-503">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-504">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-507">
         - PdfFile</span></span><br><span data-ttu-id="6a194-508">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-508">
         - Selection</span></span><br><span data-ttu-id="6a194-509">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-509">
         - Settings</span></span><br><span data-ttu-id="6a194-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-510">
         - TableBindings</span></span><br><span data-ttu-id="6a194-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-511">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-512">
         - TextBindings</span></span><br><span data-ttu-id="6a194-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-513">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-515">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-516">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-517">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-518">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="6a194-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-519">- BindingEvents</span></span><br><span data-ttu-id="6a194-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-520">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-522">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-523">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-523">
         - File</span></span><br><span data-ttu-id="6a194-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-525">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-526">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-529">
         - PdfFile</span></span><br><span data-ttu-id="6a194-530">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-530">
         - Selection</span></span><br><span data-ttu-id="6a194-531">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-531">
         - Settings</span></span><br><span data-ttu-id="6a194-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-532">
         - TableBindings</span></span><br><span data-ttu-id="6a194-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-533">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-534">
         - TextBindings</span></span><br><span data-ttu-id="6a194-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-535">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-537">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-538">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6a194-539">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="6a194-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-540">- BindingEvents</span></span><br><span data-ttu-id="6a194-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-541">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-543">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-544">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-544">
         - File</span></span><br><span data-ttu-id="6a194-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-546">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-547">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-550">
         - PdfFile</span></span><br><span data-ttu-id="6a194-551">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-551">
         - Selection</span></span><br><span data-ttu-id="6a194-552">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-552">
         - Settings</span></span><br><span data-ttu-id="6a194-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-553">
         - TableBindings</span></span><br><span data-ttu-id="6a194-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-554">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-555">
         - TextBindings</span></span><br><span data-ttu-id="6a194-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-556">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-558">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="6a194-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="6a194-559">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-560">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-561">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6a194-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6a194-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-564">- BindingEvents</span></span><br><span data-ttu-id="6a194-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-565">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-567">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-568">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-568">
         - File</span></span><br><span data-ttu-id="6a194-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-570">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-571">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-574">
         - PdfFile</span></span><br><span data-ttu-id="6a194-575">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-575">
         - Selection</span></span><br><span data-ttu-id="6a194-576">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-576">
         - Settings</span></span><br><span data-ttu-id="6a194-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-577">
         - TableBindings</span></span><br><span data-ttu-id="6a194-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-578">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-579">
         - TextBindings</span></span><br><span data-ttu-id="6a194-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-580">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-582">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-583">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-583">- TaskPane</span></span><br><span data-ttu-id="6a194-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6a194-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6a194-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-589">- BindingEvents</span></span><br><span data-ttu-id="6a194-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-590">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-592">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-593">
         - File</span></span><br><span data-ttu-id="6a194-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-595">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-596">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-599">
         - PdfFile</span></span><br><span data-ttu-id="6a194-600">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-600">
         - Selection</span></span><br><span data-ttu-id="6a194-601">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-601">
         - Settings</span></span><br><span data-ttu-id="6a194-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-602">
         - TableBindings</span></span><br><span data-ttu-id="6a194-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-603">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-604">
         - TextBindings</span></span><br><span data-ttu-id="6a194-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-605">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-607">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-608">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-608">- TaskPane</span></span><br><span data-ttu-id="6a194-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-609">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="6a194-611">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="6a194-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="6a194-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="6a194-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="6a194-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="6a194-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-614">- BindingEvents</span></span><br><span data-ttu-id="6a194-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-615">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-617">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-618">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-618">
         - File</span></span><br><span data-ttu-id="6a194-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-620">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-621">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-624">
         - PdfFile</span></span><br><span data-ttu-id="6a194-625">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-625">
         - Selection</span></span><br><span data-ttu-id="6a194-626">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-626">
         - Settings</span></span><br><span data-ttu-id="6a194-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-627">
         - TableBindings</span></span><br><span data-ttu-id="6a194-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-628">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-629">
         - TextBindings</span></span><br><span data-ttu-id="6a194-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-630">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-632">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-633">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-634">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="6a194-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="6a194-635">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="6a194-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-636">- BindingEvents</span></span><br><span data-ttu-id="6a194-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-637">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="6a194-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="6a194-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-639">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-640">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-640">
         - File</span></span><br><span data-ttu-id="6a194-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-642">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-643">
         - MatrixBindings</span></span><br><span data-ttu-id="6a194-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="6a194-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="6a194-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-646">
         - PdfFile</span></span><br><span data-ttu-id="6a194-647">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-647">
         - Selection</span></span><br><span data-ttu-id="6a194-648">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-648">
         - Settings</span></span><br><span data-ttu-id="6a194-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-649">
         - TableBindings</span></span><br><span data-ttu-id="6a194-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-650">
         - TableCoercion</span></span><br><span data-ttu-id="6a194-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="6a194-651">
         - TextBindings</span></span><br><span data-ttu-id="6a194-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-652">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="6a194-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="6a194-654">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6a194-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="6a194-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="6a194-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6a194-656">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-656">Platform</span></span></th>
    <th><span data-ttu-id="6a194-657">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-657">Extension points</span></span></th>
    <th><span data-ttu-id="6a194-658">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="6a194-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-659"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="6a194-660">Office Online</span></span></td>
    <td> <span data-ttu-id="6a194-661">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-661">- Content</span></span><br><span data-ttu-id="6a194-662">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-662">
         - TaskPane</span></span><br><span data-ttu-id="6a194-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-663">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-664">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-665">- ActiveView</span></span><br><span data-ttu-id="6a194-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-666">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-667">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-668">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-668">
         - File</span></span><br><span data-ttu-id="6a194-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-669">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-670">
         - PdfFile</span></span><br><span data-ttu-id="6a194-671">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-671">
         - Selection</span></span><br><span data-ttu-id="6a194-672">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-672">
         - Settings</span></span><br><span data-ttu-id="6a194-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-674">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-675">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-675">- Content</span></span><br><span data-ttu-id="6a194-676">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-676">
         - TaskPane</span></span><br><span data-ttu-id="6a194-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-678">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-679">- ActiveView</span></span><br><span data-ttu-id="6a194-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-680">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-681">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-682">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-682">
         - File</span></span><br><span data-ttu-id="6a194-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-683">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-684">
         - PdfFile</span></span><br><span data-ttu-id="6a194-685">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-685">
         - Selection</span></span><br><span data-ttu-id="6a194-686">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-686">
         - Settings</span></span><br><span data-ttu-id="6a194-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-688">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-689">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-689">- Content</span></span><br><span data-ttu-id="6a194-690">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-690">
         - TaskPane</span></span><br><span data-ttu-id="6a194-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-691">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-692">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-693">- ActiveView</span></span><br><span data-ttu-id="6a194-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-694">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-695">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-696">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-696">
         - File</span></span><br><span data-ttu-id="6a194-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-697">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-698">
         - PdfFile</span></span><br><span data-ttu-id="6a194-699">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-699">
         - Selection</span></span><br><span data-ttu-id="6a194-700">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-700">
         - Settings</span></span><br><span data-ttu-id="6a194-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-702">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-703">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-703">- Content</span></span><br><span data-ttu-id="6a194-704">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6a194-705">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="6a194-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-706">- ActiveView</span></span><br><span data-ttu-id="6a194-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-707">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-708">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-709">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-709">
         - File</span></span><br><span data-ttu-id="6a194-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-710">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-711">
         - PdfFile</span></span><br><span data-ttu-id="6a194-712">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-712">
         - Selection</span></span><br><span data-ttu-id="6a194-713">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-713">
         - Settings</span></span><br><span data-ttu-id="6a194-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-715">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-716">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-716">- Content</span></span><br><span data-ttu-id="6a194-717">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="6a194-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6a194-718">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="6a194-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-719">- ActiveView</span></span><br><span data-ttu-id="6a194-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-720">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-721">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-722">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-722">
         - File</span></span><br><span data-ttu-id="6a194-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-723">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-724">
         - PdfFile</span></span><br><span data-ttu-id="6a194-725">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-725">
         - Selection</span></span><br><span data-ttu-id="6a194-726">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-726">
         - Settings</span></span><br><span data-ttu-id="6a194-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-728">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="6a194-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="6a194-729">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-729">- Content</span></span><br><span data-ttu-id="6a194-730">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="6a194-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-732">- ActiveView</span></span><br><span data-ttu-id="6a194-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-733">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-734">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-735">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-735">
         - File</span></span><br><span data-ttu-id="6a194-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-736">
         - PdfFile</span></span><br><span data-ttu-id="6a194-737">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-737">
         - Selection</span></span><br><span data-ttu-id="6a194-738">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-738">
         - Settings</span></span><br><span data-ttu-id="6a194-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-739">
         - TextCoercion</span></span><br><span data-ttu-id="6a194-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-741">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-742">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-742">- Content</span></span><br><span data-ttu-id="6a194-743">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-743">
         - TaskPane</span></span><br><span data-ttu-id="6a194-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-745">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-746">- ActiveView</span></span><br><span data-ttu-id="6a194-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-747">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-748">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-749">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-749">
         - File</span></span><br><span data-ttu-id="6a194-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-750">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-751">
         - PdfFile</span></span><br><span data-ttu-id="6a194-752">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-752">
         - Selection</span></span><br><span data-ttu-id="6a194-753">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-753">
         - Settings</span></span><br><span data-ttu-id="6a194-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-755">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-756">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-756">- Content</span></span><br><span data-ttu-id="6a194-757">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-757">
         - TaskPane</span></span><br><span data-ttu-id="6a194-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-759">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-760">- ActiveView</span></span><br><span data-ttu-id="6a194-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-761">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-762">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-763">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-763">
         - File</span></span><br><span data-ttu-id="6a194-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-764">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-765">
         - PdfFile</span></span><br><span data-ttu-id="6a194-766">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-766">
         - Selection</span></span><br><span data-ttu-id="6a194-767">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-767">
         - Settings</span></span><br><span data-ttu-id="6a194-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-769">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="6a194-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="6a194-770">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-770">- Content</span></span><br><span data-ttu-id="6a194-771">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-771">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="6a194-772">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="6a194-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="6a194-773">- ActiveView</span></span><br><span data-ttu-id="6a194-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="6a194-774">
         - CompressedFile</span></span><br><span data-ttu-id="6a194-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-775">
         - DocumentEvents</span></span><br><span data-ttu-id="6a194-776">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="6a194-776">
         - File</span></span><br><span data-ttu-id="6a194-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-777">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="6a194-778">
         - PdfFile</span></span><br><span data-ttu-id="6a194-779">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-779">
         - Selection</span></span><br><span data-ttu-id="6a194-780">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-780">
         - Settings</span></span><br><span data-ttu-id="6a194-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="6a194-782">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="6a194-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="6a194-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="6a194-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6a194-784">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-784">Platform</span></span></th>
    <th><span data-ttu-id="6a194-785">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-785">Extension points</span></span></th>
    <th><span data-ttu-id="6a194-786">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="6a194-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-787"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="6a194-788">Office Online</span></span></td>
    <td> <span data-ttu-id="6a194-789">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="6a194-789">- Content</span></span><br><span data-ttu-id="6a194-790">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-790">
         - TaskPane</span></span><br><span data-ttu-id="6a194-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="6a194-791">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="6a194-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-792">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="6a194-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="6a194-794">- DocumentEvents</span></span><br><span data-ttu-id="6a194-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="6a194-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-796">
         - ImageCoercion</span></span><br><span data-ttu-id="6a194-797">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="6a194-797">
         - Settings</span></span><br><span data-ttu-id="6a194-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="6a194-799">Project</span><span class="sxs-lookup"><span data-stu-id="6a194-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="6a194-800">Plataforma</span><span class="sxs-lookup"><span data-stu-id="6a194-800">Platform</span></span></th>
    <th><span data-ttu-id="6a194-801">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="6a194-801">Extension points</span></span></th>
    <th><span data-ttu-id="6a194-802">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="6a194-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="6a194-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="6a194-803"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-804">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-805">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-806">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-807">- Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-807">- Selection</span></span><br><span data-ttu-id="6a194-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-809">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-810">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-812">- Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-812">- Selection</span></span><br><span data-ttu-id="6a194-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="6a194-814">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="6a194-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="6a194-815">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="6a194-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="6a194-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="6a194-816">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="6a194-817">- Seleção</span><span class="sxs-lookup"><span data-stu-id="6a194-817">- Selection</span></span><br><span data-ttu-id="6a194-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="6a194-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="6a194-819">Confira também</span><span class="sxs-lookup"><span data-stu-id="6a194-819">See also</span></span>

- [<span data-ttu-id="6a194-820">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="6a194-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="6a194-821">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="6a194-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="6a194-822">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="6a194-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="6a194-823">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="6a194-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="6a194-824">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="6a194-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="6a194-825">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="6a194-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="6a194-826">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6a194-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="6a194-827">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="6a194-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)