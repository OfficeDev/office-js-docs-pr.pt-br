---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 28a6d0e4c86d05855ed9d24461dbeb77454d2b48
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872127"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="cb0c8-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb0c8-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="cb0c8-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="cb0c8-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="cb0c8-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="cb0c8-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="cb0c8-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="cb0c8-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="cb0c8-108">O número de build para uma compra avulsa do Office 2019 é 16.0.10827.20150.</span><span class="sxs-lookup"><span data-stu-id="cb0c8-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="cb0c8-109">Excel</span><span class="sxs-lookup"><span data-stu-id="cb0c8-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="cb0c8-110">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="cb0c8-111">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="cb0c8-112">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="cb0c8-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-113"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="cb0c8-114">Office Online</span></span></td>
    <td> <span data-ttu-id="cb0c8-115">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-115">- TaskPane</span></span><br><span data-ttu-id="cb0c8-116">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-116">
        - Content</span></span><br><span data-ttu-id="cb0c8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="cb0c8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="cb0c8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-127">
        - BindingEvents</span></span><br><span data-ttu-id="cb0c8-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-128">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-129">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-130">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-130">
        - File</span></span><br><span data-ttu-id="cb0c8-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-131">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-133">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-133">
        - Selection</span></span><br><span data-ttu-id="cb0c8-134">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-134">
        - Settings</span></span><br><span data-ttu-id="cb0c8-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-135">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-136">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-137">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-139">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-140">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-140">- TaskPane</span></span><br><span data-ttu-id="cb0c8-141">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-141">
        - Content</span></span><br><span data-ttu-id="cb0c8-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="cb0c8-142">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="cb0c8-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-143">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-144">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-152">
        - BindingEvents</span></span><br><span data-ttu-id="cb0c8-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-153">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-154">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-155">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-155">
        - File</span></span><br><span data-ttu-id="cb0c8-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-156">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-158">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-158">
        - Selection</span></span><br><span data-ttu-id="cb0c8-159">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-159">
        - Settings</span></span><br><span data-ttu-id="cb0c8-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-160">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-161">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-162">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-164">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="cb0c8-165">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-165">- TaskPane</span></span><br><span data-ttu-id="cb0c8-166">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-166">
        - Content</span></span><br><span data-ttu-id="cb0c8-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-167">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="cb0c8-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-168">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-169">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-170">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-171">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-173">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-177">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-178">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-179">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-180">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-180">
        - File</span></span><br><span data-ttu-id="cb0c8-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-181">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-182">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-184">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-184">
        - Selection</span></span><br><span data-ttu-id="cb0c8-185">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-185">
        - Settings</span></span><br><span data-ttu-id="cb0c8-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-186">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-187">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-188">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-190">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="cb0c8-191">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-191">- TaskPane</span></span><br><span data-ttu-id="cb0c8-192">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-192">
        - Content</span></span></td>
    <td><span data-ttu-id="cb0c8-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-193">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-194">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="cb0c8-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-195">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-196">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-197">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-198">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-198">
        - File</span></span><br><span data-ttu-id="cb0c8-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-199">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-200">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-202">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-202">
        - Selection</span></span><br><span data-ttu-id="cb0c8-203">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-203">
        - Settings</span></span><br><span data-ttu-id="cb0c8-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-204">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-205">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-206">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-208">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="cb0c8-209">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-209">
        - TaskPane</span></span><br><span data-ttu-id="cb0c8-210">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="cb0c8-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-211">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="cb0c8-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-212">
        - BindingEvents</span></span><br><span data-ttu-id="cb0c8-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-213">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-214">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-215">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-215">
        - File</span></span><br><span data-ttu-id="cb0c8-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-216">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-217">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-219">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-219">
        - Selection</span></span><br><span data-ttu-id="cb0c8-220">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-220">
        - Settings</span></span><br><span data-ttu-id="cb0c8-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-221">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-222">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-223">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-225">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="cb0c8-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="cb0c8-226">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-226">- TaskPane</span></span><br><span data-ttu-id="cb0c8-227">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-227">
        - Content</span></span></td>
    <td><span data-ttu-id="cb0c8-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-228">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-229">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-230">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-231">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-232">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-233">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-234">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-235">
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-236">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-237">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-238">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-239">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-240">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-240">
        - File</span></span><br><span data-ttu-id="cb0c8-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-241">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-242">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-244">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-244">
        - Selection</span></span><br><span data-ttu-id="cb0c8-245">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-245">
        - Settings</span></span><br><span data-ttu-id="cb0c8-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-246">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-247">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-248">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-250">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="cb0c8-251">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-251">- TaskPane</span></span><br><span data-ttu-id="cb0c8-252">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-252">
        - Content</span></span><br><span data-ttu-id="cb0c8-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="cb0c8-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-254">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-256">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-257">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-258">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-259">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-260">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-261">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-262">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-263">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-264">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-265">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-266">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-266">
        - File</span></span><br><span data-ttu-id="cb0c8-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-267">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-268">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-270">
        - PdfFile</span></span><br><span data-ttu-id="cb0c8-271">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-271">
        - Selection</span></span><br><span data-ttu-id="cb0c8-272">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-272">
        - Settings</span></span><br><span data-ttu-id="cb0c8-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-273">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-274">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-275">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-277">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="cb0c8-278">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-278">- TaskPane</span></span><br><span data-ttu-id="cb0c8-279">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-279">
        - Content</span></span><br><span data-ttu-id="cb0c8-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="cb0c8-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-281">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="cb0c8-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="cb0c8-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="cb0c8-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-287">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="cb0c8-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-288">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="cb0c8-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-289">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="cb0c8-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-290">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-291">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-292">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-293">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-293">
        - File</span></span><br><span data-ttu-id="cb0c8-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-294">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-295">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-297">
        - PdfFile</span></span><br><span data-ttu-id="cb0c8-298">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-298">
        - Selection</span></span><br><span data-ttu-id="cb0c8-299">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-299">
        - Settings</span></span><br><span data-ttu-id="cb0c8-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-300">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-301">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-302">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-304">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="cb0c8-305">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-305">- TaskPane</span></span><br><span data-ttu-id="cb0c8-306">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-306">
        - Content</span></span></td>
    <td><span data-ttu-id="cb0c8-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-307">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="cb0c8-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-309">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-310">
        - CompressedFile</span></span><br><span data-ttu-id="cb0c8-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-311">
        - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-312">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-312">
        - File</span></span><br><span data-ttu-id="cb0c8-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-313">
        - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-314">
        - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-316">
        - PdfFile</span></span><br><span data-ttu-id="cb0c8-317">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-317">
        - Selection</span></span><br><span data-ttu-id="cb0c8-318">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-318">
        - Settings</span></span><br><span data-ttu-id="cb0c8-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-319">
        - TableBindings</span></span><br><span data-ttu-id="cb0c8-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-320">
        - TableCoercion</span></span><br><span data-ttu-id="cb0c8-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-321">
        - TextBindings</span></span><br><span data-ttu-id="cb0c8-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="cb0c8-323">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="cb0c8-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="cb0c8-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cb0c8-325">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-325">Platform</span></span></th>
    <th><span data-ttu-id="cb0c8-326">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-326">Extension points</span></span></th>
    <th><span data-ttu-id="cb0c8-327">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="cb0c8-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-328"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="cb0c8-329">Office Online</span></span></td>
    <td> <span data-ttu-id="cb0c8-330">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-330">- Mail Read</span></span><br><span data-ttu-id="cb0c8-331">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-331">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-333">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-334">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-335">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-336">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-337">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-338">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="cb0c8-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-339">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="cb0c8-340">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-341">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-342">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-342">- Mail Read</span></span><br><span data-ttu-id="cb0c8-343">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-343">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="cb0c8-345">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="cb0c8-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="cb0c8-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-346">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-347">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-348">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-349">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-350">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-351">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="cb0c8-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-352">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="cb0c8-353">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-354">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-355">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-355">- Mail Read</span></span><br><span data-ttu-id="cb0c8-356">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-356">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-357">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="cb0c8-358">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="cb0c8-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="cb0c8-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-359">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-360">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-361">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-362">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-363">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-364">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="cb0c8-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="cb0c8-366">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-367">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-368">- Mail Read</span></span><br><span data-ttu-id="cb0c8-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-369">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="cb0c8-371">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="cb0c8-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="cb0c8-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-372">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="cb0c8-376">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-377">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-378">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-378">- Mail Read</span></span><br><span data-ttu-id="cb0c8-379">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="cb0c8-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="cb0c8-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="cb0c8-384">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-385">Office 365 para iOS</span><span class="sxs-lookup"><span data-stu-id="cb0c8-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="cb0c8-386">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-386">- Mail Read</span></span><br><span data-ttu-id="cb0c8-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-388">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="cb0c8-393">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-394">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-395">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-395">- Mail Read</span></span><br><span data-ttu-id="cb0c8-396">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-396">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-397">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cb0c8-404">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-405">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-406">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-406">- Mail Read</span></span><br><span data-ttu-id="cb0c8-407">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-407">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-408">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cb0c8-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-416">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-417">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-417">- Mail Read</span></span><br><span data-ttu-id="cb0c8-418">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-418">
      - Mail Compose</span></span><br><span data-ttu-id="cb0c8-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="cb0c8-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="cb0c8-426">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-427">Office 365 para Android</span><span class="sxs-lookup"><span data-stu-id="cb0c8-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="cb0c8-428">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="cb0c8-428">- Mail Read</span></span><br><span data-ttu-id="cb0c8-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-429">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="cb0c8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="cb0c8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="cb0c8-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="cb0c8-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="cb0c8-435">Não disponível</span><span class="sxs-lookup"><span data-stu-id="cb0c8-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="cb0c8-436">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="cb0c8-437">Word</span><span class="sxs-lookup"><span data-stu-id="cb0c8-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cb0c8-438">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-438">Platform</span></span></th>
    <th><span data-ttu-id="cb0c8-439">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-439">Extension points</span></span></th>
    <th><span data-ttu-id="cb0c8-440">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="cb0c8-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-441"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="cb0c8-442">Office Online</span></span></td>
    <td> <span data-ttu-id="cb0c8-443">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-443">- TaskPane</span></span><br><span data-ttu-id="cb0c8-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-444">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-445">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-446">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-447">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-448">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-449">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-451">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-452">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-452">
         - File</span></span><br><span data-ttu-id="cb0c8-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-454">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-455">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-458">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-459">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-459">
         - Selection</span></span><br><span data-ttu-id="cb0c8-460">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-460">
         - Settings</span></span><br><span data-ttu-id="cb0c8-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-461">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-462">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-463">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-464">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-466">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-467">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-467">- TaskPane</span></span><br><span data-ttu-id="cb0c8-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-468">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-469">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-470">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-471">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-472">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-473">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-474">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-476">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-477">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-477">
         - File</span></span><br><span data-ttu-id="cb0c8-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-479">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-480">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-483">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-484">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-484">
         - Selection</span></span><br><span data-ttu-id="cb0c8-485">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-485">
         - Settings</span></span><br><span data-ttu-id="cb0c8-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-486">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-487">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-488">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-489">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-491">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-492">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-492">- TaskPane</span></span><br><span data-ttu-id="cb0c8-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-493">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-494">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-495">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-496">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-497">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-498">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-499">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-501">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-502">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-502">
         - File</span></span><br><span data-ttu-id="cb0c8-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-504">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-505">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-508">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-509">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-509">
         - Selection</span></span><br><span data-ttu-id="cb0c8-510">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-510">
         - Settings</span></span><br><span data-ttu-id="cb0c8-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-511">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-512">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-513">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-514">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-516">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-517">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-518">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-519">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="cb0c8-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-520">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-521">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-523">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-524">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-524">
         - File</span></span><br><span data-ttu-id="cb0c8-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-526">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-527">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-530">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-531">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-531">
         - Selection</span></span><br><span data-ttu-id="cb0c8-532">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-532">
         - Settings</span></span><br><span data-ttu-id="cb0c8-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-533">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-534">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-535">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-536">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-538">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-539">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-540">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="cb0c8-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-541">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-542">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-544">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-545">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-545">
         - File</span></span><br><span data-ttu-id="cb0c8-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-547">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-548">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-551">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-552">
         - Selection</span></span><br><span data-ttu-id="cb0c8-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-553">
         - Settings</span></span><br><span data-ttu-id="cb0c8-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-554">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-555">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-556">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-557">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-559">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="cb0c8-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="cb0c8-560">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-561">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-563">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cb0c8-564">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cb0c8-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-565">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-566">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-568">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-569">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-569">
         - File</span></span><br><span data-ttu-id="cb0c8-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-571">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-572">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-575">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-576">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-576">
         - Selection</span></span><br><span data-ttu-id="cb0c8-577">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-577">
         - Settings</span></span><br><span data-ttu-id="cb0c8-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-578">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-579">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-580">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-581">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-583">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-584">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-584">- TaskPane</span></span><br><span data-ttu-id="cb0c8-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cb0c8-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cb0c8-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-590">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-591">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-593">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-594">
         - File</span></span><br><span data-ttu-id="cb0c8-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-596">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-597">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-600">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-601">
         - Selection</span></span><br><span data-ttu-id="cb0c8-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-602">
         - Settings</span></span><br><span data-ttu-id="cb0c8-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-603">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-604">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-605">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-606">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-608">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-609">- TaskPane</span></span><br><span data-ttu-id="cb0c8-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-612">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="cb0c8-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-613">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="cb0c8-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="cb0c8-614">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="cb0c8-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-615">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-616">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-618">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-619">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-619">
         - File</span></span><br><span data-ttu-id="cb0c8-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-621">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-622">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-625">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-626">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-626">
         - Selection</span></span><br><span data-ttu-id="cb0c8-627">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-627">
         - Settings</span></span><br><span data-ttu-id="cb0c8-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-628">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-629">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-630">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-631">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-633">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-634">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="cb0c8-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-637">- BindingEvents</span></span><br><span data-ttu-id="cb0c8-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-638">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="cb0c8-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="cb0c8-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-640">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-641">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-641">
         - File</span></span><br><span data-ttu-id="cb0c8-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-643">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-644">
         - MatrixBindings</span></span><br><span data-ttu-id="cb0c8-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="cb0c8-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="cb0c8-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-647">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-648">
         - Selection</span></span><br><span data-ttu-id="cb0c8-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-649">
         - Settings</span></span><br><span data-ttu-id="cb0c8-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-650">
         - TableBindings</span></span><br><span data-ttu-id="cb0c8-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-651">
         - TableCoercion</span></span><br><span data-ttu-id="cb0c8-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="cb0c8-652">
         - TextBindings</span></span><br><span data-ttu-id="cb0c8-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-653">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="cb0c8-655">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="cb0c8-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cb0c8-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cb0c8-657">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-657">Platform</span></span></th>
    <th><span data-ttu-id="cb0c8-658">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-658">Extension points</span></span></th>
    <th><span data-ttu-id="cb0c8-659">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="cb0c8-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-660"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="cb0c8-661">Office Online</span></span></td>
    <td> <span data-ttu-id="cb0c8-662">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-662">- Content</span></span><br><span data-ttu-id="cb0c8-663">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-663">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-664">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-665">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-666">- ActiveView</span></span><br><span data-ttu-id="cb0c8-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-667">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-668">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-669">
         - File</span></span><br><span data-ttu-id="cb0c8-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-670">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-671">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-672">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-672">
         - Selection</span></span><br><span data-ttu-id="cb0c8-673">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-673">
         - Settings</span></span><br><span data-ttu-id="cb0c8-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-675">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-676">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-676">- Content</span></span><br><span data-ttu-id="cb0c8-677">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-677">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-679">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-680">- ActiveView</span></span><br><span data-ttu-id="cb0c8-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-681">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-682">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-683">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-683">
         - File</span></span><br><span data-ttu-id="cb0c8-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-684">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-685">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-686">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-686">
         - Selection</span></span><br><span data-ttu-id="cb0c8-687">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-687">
         - Settings</span></span><br><span data-ttu-id="cb0c8-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-689">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-690">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-690">- Content</span></span><br><span data-ttu-id="cb0c8-691">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-691">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-692">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-693">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-694">- ActiveView</span></span><br><span data-ttu-id="cb0c8-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-695">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-696">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-697">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-697">
         - File</span></span><br><span data-ttu-id="cb0c8-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-698">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-699">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-700">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-700">
         - Selection</span></span><br><span data-ttu-id="cb0c8-701">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-701">
         - Settings</span></span><br><span data-ttu-id="cb0c8-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-703">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-704">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-704">- Content</span></span><br><span data-ttu-id="cb0c8-705">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-706">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="cb0c8-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-707">- ActiveView</span></span><br><span data-ttu-id="cb0c8-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-708">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-709">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-710">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-710">
         - File</span></span><br><span data-ttu-id="cb0c8-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-711">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-712">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-713">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-713">
         - Selection</span></span><br><span data-ttu-id="cb0c8-714">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-714">
         - Settings</span></span><br><span data-ttu-id="cb0c8-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-716">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-717">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-717">- Content</span></span><br><span data-ttu-id="cb0c8-718">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="cb0c8-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-719">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="cb0c8-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-720">- ActiveView</span></span><br><span data-ttu-id="cb0c8-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-721">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-722">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-723">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-723">
         - File</span></span><br><span data-ttu-id="cb0c8-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-724">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-725">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-726">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-726">
         - Selection</span></span><br><span data-ttu-id="cb0c8-727">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-727">
         - Settings</span></span><br><span data-ttu-id="cb0c8-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-729">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="cb0c8-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="cb0c8-730">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-730">- Content</span></span><br><span data-ttu-id="cb0c8-731">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="cb0c8-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-733">- ActiveView</span></span><br><span data-ttu-id="cb0c8-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-734">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-735">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-736">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-736">
         - File</span></span><br><span data-ttu-id="cb0c8-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-737">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-738">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-738">
         - Selection</span></span><br><span data-ttu-id="cb0c8-739">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-739">
         - Settings</span></span><br><span data-ttu-id="cb0c8-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-740">
         - TextCoercion</span></span><br><span data-ttu-id="cb0c8-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-742">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-743">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-743">- Content</span></span><br><span data-ttu-id="cb0c8-744">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-744">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-747">- ActiveView</span></span><br><span data-ttu-id="cb0c8-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-748">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-749">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-750">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-750">
         - File</span></span><br><span data-ttu-id="cb0c8-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-751">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-752">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-753">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-753">
         - Selection</span></span><br><span data-ttu-id="cb0c8-754">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-754">
         - Settings</span></span><br><span data-ttu-id="cb0c8-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-756">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-757">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-757">- Content</span></span><br><span data-ttu-id="cb0c8-758">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-758">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-761">- ActiveView</span></span><br><span data-ttu-id="cb0c8-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-762">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-763">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-764">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-764">
         - File</span></span><br><span data-ttu-id="cb0c8-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-765">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-766">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-767">
         - Selection</span></span><br><span data-ttu-id="cb0c8-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-768">
         - Settings</span></span><br><span data-ttu-id="cb0c8-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-770">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="cb0c8-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="cb0c8-771">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-771">- Content</span></span><br><span data-ttu-id="cb0c8-772">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-773">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="cb0c8-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="cb0c8-774">- ActiveView</span></span><br><span data-ttu-id="cb0c8-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-775">
         - CompressedFile</span></span><br><span data-ttu-id="cb0c8-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-776">
         - DocumentEvents</span></span><br><span data-ttu-id="cb0c8-777">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-777">
         - File</span></span><br><span data-ttu-id="cb0c8-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-778">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="cb0c8-779">
         - PdfFile</span></span><br><span data-ttu-id="cb0c8-780">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-780">
         - Selection</span></span><br><span data-ttu-id="cb0c8-781">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-781">
         - Settings</span></span><br><span data-ttu-id="cb0c8-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="cb0c8-783">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="cb0c8-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="cb0c8-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="cb0c8-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cb0c8-785">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-785">Platform</span></span></th>
    <th><span data-ttu-id="cb0c8-786">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-786">Extension points</span></span></th>
    <th><span data-ttu-id="cb0c8-787">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="cb0c8-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-788"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="cb0c8-789">Office Online</span></span></td>
    <td> <span data-ttu-id="cb0c8-790">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="cb0c8-790">- Content</span></span><br><span data-ttu-id="cb0c8-791">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-791">
         - TaskPane</span></span><br><span data-ttu-id="cb0c8-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-793">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="cb0c8-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="cb0c8-795">- DocumentEvents</span></span><br><span data-ttu-id="cb0c8-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="cb0c8-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-797">
         - ImageCoercion</span></span><br><span data-ttu-id="cb0c8-798">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="cb0c8-798">
         - Settings</span></span><br><span data-ttu-id="cb0c8-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="cb0c8-800">Project</span><span class="sxs-lookup"><span data-stu-id="cb0c8-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="cb0c8-801">Plataforma</span><span class="sxs-lookup"><span data-stu-id="cb0c8-801">Platform</span></span></th>
    <th><span data-ttu-id="cb0c8-802">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="cb0c8-802">Extension points</span></span></th>
    <th><span data-ttu-id="cb0c8-803">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="cb0c8-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-804"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-805">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-806">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-808">- Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-808">- Selection</span></span><br><span data-ttu-id="cb0c8-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-810">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-811">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-813">- Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-813">- Selection</span></span><br><span data-ttu-id="cb0c8-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="cb0c8-815">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="cb0c8-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="cb0c8-816">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="cb0c8-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="cb0c8-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="cb0c8-817">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="cb0c8-818">- Seleção</span><span class="sxs-lookup"><span data-stu-id="cb0c8-818">- Selection</span></span><br><span data-ttu-id="cb0c8-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="cb0c8-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="cb0c8-820">Confira também</span><span class="sxs-lookup"><span data-stu-id="cb0c8-820">See also</span></span>

- [<span data-ttu-id="cb0c8-821">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="cb0c8-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="cb0c8-822">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="cb0c8-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="cb0c8-823">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="cb0c8-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="cb0c8-824">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="cb0c8-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
