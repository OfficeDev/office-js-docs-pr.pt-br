---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: 636c6290d8c67901beb195990593727485467460
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512878"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="268f0-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="268f0-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="268f0-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="268f0-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="268f0-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="268f0-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="268f0-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="268f0-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="268f0-108">Excel</span><span class="sxs-lookup"><span data-stu-id="268f0-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="268f0-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="268f0-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="268f0-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="268f0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="268f0-113">Office Online</span></span></td>
    <td> <span data-ttu-id="268f0-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-114">- TaskPane</span></span><br><span data-ttu-id="268f0-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-115">
        - Content</span></span><br><span data-ttu-id="268f0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="268f0-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="268f0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="268f0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="268f0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="268f0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="268f0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="268f0-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="268f0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="268f0-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="268f0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="268f0-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-126">
        - BindingEvents</span></span><br><span data-ttu-id="268f0-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-127">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-128">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-129">
        - File</span></span><br><span data-ttu-id="268f0-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-130">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-132">
        - Selection</span></span><br><span data-ttu-id="268f0-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-133">
        - Settings</span></span><br><span data-ttu-id="268f0-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-134">
        - TableBindings</span></span><br><span data-ttu-id="268f0-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-135">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-136">
        - TextBindings</span></span><br><span data-ttu-id="268f0-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="268f0-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-139">
        - TaskPane</span></span><br><span data-ttu-id="268f0-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="268f0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="268f0-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="268f0-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-142">
        - BindingEvents</span></span><br><span data-ttu-id="268f0-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-143">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-144">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-145">
        - File</span></span><br><span data-ttu-id="268f0-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-146">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-147">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-149">
        - Selection</span></span><br><span data-ttu-id="268f0-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-150">
        - Settings</span></span><br><span data-ttu-id="268f0-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-151">
        - TableBindings</span></span><br><span data-ttu-id="268f0-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-152">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-153">
        - TextBindings</span></span><br><span data-ttu-id="268f0-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="268f0-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-156">- TaskPane</span></span><br><span data-ttu-id="268f0-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-157">
        - Content</span></span></td>
    <td><span data-ttu-id="268f0-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="268f0-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="268f0-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-160">- BindingEvents</span></span><br><span data-ttu-id="268f0-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-161">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-162">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-163">
        - File</span></span><br><span data-ttu-id="268f0-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-164">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-165">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-167">
        - Selection</span></span><br><span data-ttu-id="268f0-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-168">
        - Settings</span></span><br><span data-ttu-id="268f0-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-169">
        - TableBindings</span></span><br><span data-ttu-id="268f0-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-170">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-171">
        - TextBindings</span></span><br><span data-ttu-id="268f0-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-173">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="268f0-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-174">- TaskPane</span></span><br><span data-ttu-id="268f0-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-175">
        - Content</span></span><br><span data-ttu-id="268f0-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="268f0-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="268f0-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="268f0-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="268f0-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="268f0-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="268f0-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="268f0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="268f0-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="268f0-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="268f0-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-186">- BindingEvents</span></span><br><span data-ttu-id="268f0-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-187">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-188">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-189">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-189">
        - File</span></span><br><span data-ttu-id="268f0-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-190">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-191">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-193">
        - Selection</span></span><br><span data-ttu-id="268f0-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-194">
        - Settings</span></span><br><span data-ttu-id="268f0-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-195">
        - TableBindings</span></span><br><span data-ttu-id="268f0-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-196">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-197">
        - TextBindings</span></span><br><span data-ttu-id="268f0-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-199">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="268f0-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="268f0-200">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-200">- TaskPane</span></span><br><span data-ttu-id="268f0-201">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-201">
        - Content</span></span></td>
    <td><span data-ttu-id="268f0-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="268f0-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="268f0-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="268f0-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="268f0-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="268f0-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="268f0-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="268f0-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="268f0-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="268f0-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-211">- BindingEvents</span></span><br><span data-ttu-id="268f0-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-212">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-213">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-214">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-214">
        - File</span></span><br><span data-ttu-id="268f0-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-215">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-216">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-218">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-218">
        - Selection</span></span><br><span data-ttu-id="268f0-219">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-219">
        - Settings</span></span><br><span data-ttu-id="268f0-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-220">
        - TableBindings</span></span><br><span data-ttu-id="268f0-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-221">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-222">
        - TextBindings</span></span><br><span data-ttu-id="268f0-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-224">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="268f0-225">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-225">- TaskPane</span></span><br><span data-ttu-id="268f0-226">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-226">
        - Content</span></span></td>
    <td><span data-ttu-id="268f0-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="268f0-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="268f0-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-229">- BindingEvents</span></span><br><span data-ttu-id="268f0-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-230">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-231">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-232">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-232">
        - File</span></span><br><span data-ttu-id="268f0-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-233">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-234">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-236">
        - PdfFile</span></span><br><span data-ttu-id="268f0-237">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-237">
        - Selection</span></span><br><span data-ttu-id="268f0-238">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-238">
        - Settings</span></span><br><span data-ttu-id="268f0-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-239">
        - TableBindings</span></span><br><span data-ttu-id="268f0-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-240">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-241">
        - TextBindings</span></span><br><span data-ttu-id="268f0-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-243">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="268f0-244">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-244">- TaskPane</span></span><br><span data-ttu-id="268f0-245">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-245">
        - Content</span></span><br><span data-ttu-id="268f0-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="268f0-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="268f0-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="268f0-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="268f0-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="268f0-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="268f0-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="268f0-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="268f0-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="268f0-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="268f0-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="268f0-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-256">- BindingEvents</span></span><br><span data-ttu-id="268f0-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-257">
        - CompressedFile</span></span><br><span data-ttu-id="268f0-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-258">
        - DocumentEvents</span></span><br><span data-ttu-id="268f0-259">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-259">
        - File</span></span><br><span data-ttu-id="268f0-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-260">
        - ImageCoercion</span></span><br><span data-ttu-id="268f0-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-261">
        - MatrixBindings</span></span><br><span data-ttu-id="268f0-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="268f0-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-263">
        - PdfFile</span></span><br><span data-ttu-id="268f0-264">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-264">
        - Selection</span></span><br><span data-ttu-id="268f0-265">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-265">
        - Settings</span></span><br><span data-ttu-id="268f0-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-266">
        - TableBindings</span></span><br><span data-ttu-id="268f0-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-267">
        - TableCoercion</span></span><br><span data-ttu-id="268f0-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-268">
        - TextBindings</span></span><br><span data-ttu-id="268f0-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="268f0-270">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="268f0-270">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="268f0-271">Outlook</span><span class="sxs-lookup"><span data-stu-id="268f0-271">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="268f0-272">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-272">Platform</span></span></th>
    <th><span data-ttu-id="268f0-273">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-273">Extension points</span></span></th>
    <th><span data-ttu-id="268f0-274">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-274">API requirement sets</span></span></th>
    <th><span data-ttu-id="268f0-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-275"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-276">Office Online</span><span class="sxs-lookup"><span data-stu-id="268f0-276">Office Online</span></span></td>
    <td> <span data-ttu-id="268f0-277">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-277">- Mail Read</span></span><br><span data-ttu-id="268f0-278">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-278">
      - Mail Compose</span></span><br><span data-ttu-id="268f0-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-279">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-280">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="268f0-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="268f0-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-286">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="268f0-287">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-288">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-288">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-289">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-289">- Mail Read</span></span><br><span data-ttu-id="268f0-290">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-290">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="268f0-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-291">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="268f0-295">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-295">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-296">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-296">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-297">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-297">- Mail Read</span></span><br><span data-ttu-id="268f0-298">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-298">
      - Mail Compose</span></span><br><span data-ttu-id="268f0-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="268f0-300">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="268f0-300">
      - Modules</span></span></td>
    <td> <span data-ttu-id="268f0-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-301">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="268f0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="268f0-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-307">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="268f0-308">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-308">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-309">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-309">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-310">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-310">- Mail Read</span></span><br><span data-ttu-id="268f0-311">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-311">
      - Mail Compose</span></span><br><span data-ttu-id="268f0-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="268f0-313">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="268f0-313">
      - Modules</span></span></td>
    <td> <span data-ttu-id="268f0-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-314">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="268f0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="268f0-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="268f0-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="268f0-321">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-321">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-322">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="268f0-322">Office for iOS</span></span></td>
    <td> <span data-ttu-id="268f0-323">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-323">- Mail Read</span></span><br><span data-ttu-id="268f0-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-325">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-329">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="268f0-330">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-330">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-331">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-331">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-332">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-332">- Mail Read</span></span><br><span data-ttu-id="268f0-333">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-333">
      - Mail Compose</span></span><br><span data-ttu-id="268f0-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-335">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="268f0-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="268f0-341">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-341">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-342">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-342">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-343">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-343">- Mail Read</span></span><br><span data-ttu-id="268f0-344">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="268f0-344">
      - Mail Compose</span></span><br><span data-ttu-id="268f0-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="268f0-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="268f0-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="268f0-352">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-353">Office para Android</span><span class="sxs-lookup"><span data-stu-id="268f0-353">Office for Android</span></span></td>
    <td> <span data-ttu-id="268f0-354">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="268f0-354">- Mail Read</span></span><br><span data-ttu-id="268f0-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-356">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="268f0-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="268f0-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="268f0-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="268f0-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="268f0-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="268f0-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="268f0-361">Não disponível</span><span class="sxs-lookup"><span data-stu-id="268f0-361">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="268f0-362">Word</span><span class="sxs-lookup"><span data-stu-id="268f0-362">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="268f0-363">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-363">Platform</span></span></th>
    <th><span data-ttu-id="268f0-364">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-364">Extension points</span></span></th>
    <th><span data-ttu-id="268f0-365">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-365">API requirement sets</span></span></th>
    <th><span data-ttu-id="268f0-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-366"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-367">Office Online</span><span class="sxs-lookup"><span data-stu-id="268f0-367">Office Online</span></span></td>
    <td> <span data-ttu-id="268f0-368">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-368">- TaskPane</span></span><br><span data-ttu-id="268f0-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-369">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-370">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="268f0-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="268f0-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-373">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-374">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-374">- BindingEvents</span></span><br><span data-ttu-id="268f0-375">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-375">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-376">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-376">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-377">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-377">
         - File</span></span><br><span data-ttu-id="268f0-378">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-378">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-379">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-379">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-380">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-380">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-381">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-381">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-382">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-382">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-383">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-383">
         - PdfFile</span></span><br><span data-ttu-id="268f0-384">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-384">
         - Selection</span></span><br><span data-ttu-id="268f0-385">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-385">
         - Settings</span></span><br><span data-ttu-id="268f0-386">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-386">
         - TableBindings</span></span><br><span data-ttu-id="268f0-387">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-387">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-388">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-388">
         - TextBindings</span></span><br><span data-ttu-id="268f0-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-389">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-390">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-390">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-391">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-392">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-392">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="268f0-393">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="268f0-394">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-394">- BindingEvents</span></span><br><span data-ttu-id="268f0-395">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-395">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-396">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-396">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-397">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-398">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-398">
         - File</span></span><br><span data-ttu-id="268f0-399">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-399">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-400">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-400">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-401">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-401">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-402">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-402">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-403">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-403">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-404">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-404">
         - PdfFile</span></span><br><span data-ttu-id="268f0-405">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-405">
         - Selection</span></span><br><span data-ttu-id="268f0-406">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-406">
         - Settings</span></span><br><span data-ttu-id="268f0-407">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-407">
         - TableBindings</span></span><br><span data-ttu-id="268f0-408">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-408">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-409">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-409">
         - TextBindings</span></span><br><span data-ttu-id="268f0-410">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-410">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-411">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-411">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-412">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-412">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-413">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-413">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-414">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="268f0-415">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="268f0-416">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-416">- BindingEvents</span></span><br><span data-ttu-id="268f0-417">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-417">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-418">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-418">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-419">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-419">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-420">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-420">
         - File</span></span><br><span data-ttu-id="268f0-421">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-421">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-422">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-422">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-423">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-423">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-424">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-424">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-425">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-425">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-426">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-426">
         - PdfFile</span></span><br><span data-ttu-id="268f0-427">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-427">
         - Selection</span></span><br><span data-ttu-id="268f0-428">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-428">
         - Settings</span></span><br><span data-ttu-id="268f0-429">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-429">
         - TableBindings</span></span><br><span data-ttu-id="268f0-430">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-430">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-431">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-431">
         - TextBindings</span></span><br><span data-ttu-id="268f0-432">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-432">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-433">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-433">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-434">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-434">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-435">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-435">- TaskPane</span></span><br><span data-ttu-id="268f0-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-436">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-437">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="268f0-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="268f0-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-440">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-441">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-441">- BindingEvents</span></span><br><span data-ttu-id="268f0-442">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-442">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-443">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-443">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-444">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-444">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-445">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-445">
         - File</span></span><br><span data-ttu-id="268f0-446">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-446">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-447">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-447">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-448">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-448">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-449">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-449">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-450">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-450">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-451">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-451">
         - PdfFile</span></span><br><span data-ttu-id="268f0-452">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-452">
         - Selection</span></span><br><span data-ttu-id="268f0-453">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-453">
         - Settings</span></span><br><span data-ttu-id="268f0-454">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-454">
         - TableBindings</span></span><br><span data-ttu-id="268f0-455">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-455">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-456">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-456">
         - TextBindings</span></span><br><span data-ttu-id="268f0-457">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-457">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-458">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-458">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-459">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="268f0-459">Office for iPad</span></span></td>
    <td> <span data-ttu-id="268f0-460">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-460">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-461">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="268f0-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="268f0-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="268f0-464">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="268f0-465">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-465">- BindingEvents</span></span><br><span data-ttu-id="268f0-466">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-466">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-467">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-467">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-468">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-468">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-469">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-469">
         - File</span></span><br><span data-ttu-id="268f0-470">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-470">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-471">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-471">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-472">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-472">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-473">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-473">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-474">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-474">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-475">
         - PdfFile</span></span><br><span data-ttu-id="268f0-476">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-476">
         - Selection</span></span><br><span data-ttu-id="268f0-477">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-477">
         - Settings</span></span><br><span data-ttu-id="268f0-478">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-478">
         - TableBindings</span></span><br><span data-ttu-id="268f0-479">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-479">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-480">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-480">
         - TextBindings</span></span><br><span data-ttu-id="268f0-481">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-481">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-482">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-482">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-483">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-483">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-484">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-484">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-485">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="268f0-486">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="268f0-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-487">- BindingEvents</span></span><br><span data-ttu-id="268f0-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-488">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-489">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-490">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-491">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-491">
         - File</span></span><br><span data-ttu-id="268f0-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-492">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-493">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-494">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-495">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-496">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-497">
         - PdfFile</span></span><br><span data-ttu-id="268f0-498">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-498">
         - Selection</span></span><br><span data-ttu-id="268f0-499">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-499">
         - Settings</span></span><br><span data-ttu-id="268f0-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-500">
         - TableBindings</span></span><br><span data-ttu-id="268f0-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-501">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-502">
         - TextBindings</span></span><br><span data-ttu-id="268f0-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-503">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-504">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-505">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-505">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-506">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-506">- TaskPane</span></span><br><span data-ttu-id="268f0-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="268f0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="268f0-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="268f0-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="268f0-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="268f0-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="268f0-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="268f0-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-512">- BindingEvents</span></span><br><span data-ttu-id="268f0-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-513">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="268f0-514">
         - CustomXmlParts</span></span><br><span data-ttu-id="268f0-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-515">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-516">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-516">
         - File</span></span><br><span data-ttu-id="268f0-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-517">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-518">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-519">
         - MatrixBindings</span></span><br><span data-ttu-id="268f0-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-520">
         - MatrixCoercion</span></span><br><span data-ttu-id="268f0-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-521">
         - OoxmlCoercion</span></span><br><span data-ttu-id="268f0-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-522">
         - PdfFile</span></span><br><span data-ttu-id="268f0-523">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-523">
         - Selection</span></span><br><span data-ttu-id="268f0-524">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-524">
         - Settings</span></span><br><span data-ttu-id="268f0-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-525">
         - TableBindings</span></span><br><span data-ttu-id="268f0-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-526">
         - TableCoercion</span></span><br><span data-ttu-id="268f0-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="268f0-527">
         - TextBindings</span></span><br><span data-ttu-id="268f0-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-528">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="268f0-529">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="268f0-530">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="268f0-530">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="268f0-531">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="268f0-531">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="268f0-532">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-532">Platform</span></span></th>
    <th><span data-ttu-id="268f0-533">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-533">Extension points</span></span></th>
    <th><span data-ttu-id="268f0-534">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-534">API requirement sets</span></span></th>
    <th><span data-ttu-id="268f0-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-535"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-536">Office Online</span><span class="sxs-lookup"><span data-stu-id="268f0-536">Office Online</span></span></td>
    <td> <span data-ttu-id="268f0-537">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-537">- Content</span></span><br><span data-ttu-id="268f0-538">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-538">
         - TaskPane</span></span><br><span data-ttu-id="268f0-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-539">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-541">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-541">- ActiveView</span></span><br><span data-ttu-id="268f0-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-542">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-543">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-544">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-544">
         - File</span></span><br><span data-ttu-id="268f0-545">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-545">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-546">
         - PdfFile</span></span><br><span data-ttu-id="268f0-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-547">
         - Selection</span></span><br><span data-ttu-id="268f0-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-548">
         - Settings</span></span><br><span data-ttu-id="268f0-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-549">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-550">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-550">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-551">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-551">- Content</span></span><br><span data-ttu-id="268f0-552">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-552">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="268f0-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="268f0-553">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="268f0-554">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-554">- ActiveView</span></span><br><span data-ttu-id="268f0-555">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-555">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-556">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-557">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-557">
         - File</span></span><br><span data-ttu-id="268f0-558">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-558">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-559">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-559">
         - PdfFile</span></span><br><span data-ttu-id="268f0-560">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-560">
         - Selection</span></span><br><span data-ttu-id="268f0-561">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-561">
         - Settings</span></span><br><span data-ttu-id="268f0-562">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-562">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-563">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-563">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-564">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-564">- Content</span></span><br><span data-ttu-id="268f0-565">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-565">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="268f0-566">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="268f0-567">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-567">- ActiveView</span></span><br><span data-ttu-id="268f0-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-568">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-569">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-570">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-570">
         - File</span></span><br><span data-ttu-id="268f0-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-571">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-572">
         - PdfFile</span></span><br><span data-ttu-id="268f0-573">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-573">
         - Selection</span></span><br><span data-ttu-id="268f0-574">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-574">
         - Settings</span></span><br><span data-ttu-id="268f0-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-575">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-576">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-576">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-577">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-577">- Content</span></span><br><span data-ttu-id="268f0-578">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-578">
         - TaskPane</span></span><br><span data-ttu-id="268f0-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-579">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-580">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-581">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-581">- ActiveView</span></span><br><span data-ttu-id="268f0-582">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-582">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-583">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-583">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-584">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-584">
         - File</span></span><br><span data-ttu-id="268f0-585">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-585">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-586">
         - PdfFile</span></span><br><span data-ttu-id="268f0-587">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-587">
         - Selection</span></span><br><span data-ttu-id="268f0-588">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-588">
         - Settings</span></span><br><span data-ttu-id="268f0-589">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-589">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-590">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="268f0-590">Office for iPad</span></span></td>
    <td> <span data-ttu-id="268f0-591">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-591">- Content</span></span><br><span data-ttu-id="268f0-592">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-592">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-593">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="268f0-594">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-594">- ActiveView</span></span><br><span data-ttu-id="268f0-595">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-595">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-596">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-596">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-597">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-597">
         - File</span></span><br><span data-ttu-id="268f0-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-598">
         - PdfFile</span></span><br><span data-ttu-id="268f0-599">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-599">
         - Selection</span></span><br><span data-ttu-id="268f0-600">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-600">
         - Settings</span></span><br><span data-ttu-id="268f0-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-601">
         - TextCoercion</span></span><br><span data-ttu-id="268f0-602">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-602">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-603">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-603">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-604">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-604">- Content</span></span><br><span data-ttu-id="268f0-605">
         - TaskPane/td></span><span class="sxs-lookup"><span data-stu-id="268f0-605">
         - TaskPane/td></span></span> <td> <span data-ttu-id="268f0-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="268f0-606">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="268f0-607">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-607">- ActiveView</span></span><br><span data-ttu-id="268f0-608">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-608">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-609">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-610">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-610">
         - File</span></span><br><span data-ttu-id="268f0-611">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-611">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-612">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-612">
         - PdfFile</span></span><br><span data-ttu-id="268f0-613">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-613">
         - Selection</span></span><br><span data-ttu-id="268f0-614">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-614">
         - Settings</span></span><br><span data-ttu-id="268f0-615">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-615">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-616">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="268f0-616">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="268f0-617">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-617">- Content</span></span><br><span data-ttu-id="268f0-618">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-618">
         - TaskPane</span></span><br><span data-ttu-id="268f0-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-619">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-620">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-621">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="268f0-621">- ActiveView</span></span><br><span data-ttu-id="268f0-622">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="268f0-622">
         - CompressedFile</span></span><br><span data-ttu-id="268f0-623">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-623">
         - DocumentEvents</span></span><br><span data-ttu-id="268f0-624">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="268f0-624">
         - File</span></span><br><span data-ttu-id="268f0-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-625">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-626">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="268f0-626">
         - PdfFile</span></span><br><span data-ttu-id="268f0-627">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-627">
         - Selection</span></span><br><span data-ttu-id="268f0-628">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-628">
         - Settings</span></span><br><span data-ttu-id="268f0-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-629">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="268f0-630">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="268f0-630">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="268f0-631">OneNote</span><span class="sxs-lookup"><span data-stu-id="268f0-631">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="268f0-632">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-632">Platform</span></span></th>
    <th><span data-ttu-id="268f0-633">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-633">Extension points</span></span></th>
    <th><span data-ttu-id="268f0-634">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-634">API requirement sets</span></span></th>
    <th><span data-ttu-id="268f0-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-635"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-636">Office Online</span><span class="sxs-lookup"><span data-stu-id="268f0-636">Office Online</span></span></td>
    <td> <span data-ttu-id="268f0-637">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="268f0-637">- Content</span></span><br><span data-ttu-id="268f0-638">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-638">
         - TaskPane</span></span><br><span data-ttu-id="268f0-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="268f0-639">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="268f0-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-640">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="268f0-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-641">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-642">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="268f0-642">- DocumentEvents</span></span><br><span data-ttu-id="268f0-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="268f0-644">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-644">
         - ImageCoercion</span></span><br><span data-ttu-id="268f0-645">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="268f0-645">
         - Settings</span></span><br><span data-ttu-id="268f0-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-646">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="268f0-647">Project</span><span class="sxs-lookup"><span data-stu-id="268f0-647">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="268f0-648">Plataforma</span><span class="sxs-lookup"><span data-stu-id="268f0-648">Platform</span></span></th>
    <th><span data-ttu-id="268f0-649">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="268f0-649">Extension points</span></span></th>
    <th><span data-ttu-id="268f0-650">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="268f0-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="268f0-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="268f0-651"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-652">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-652">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-653">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-653">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-654">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-655">- Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-655">- Selection</span></span><br><span data-ttu-id="268f0-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-656">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-657">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-657">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-658">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-658">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-659">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-660">- Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-660">- Selection</span></span><br><span data-ttu-id="268f0-661">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-661">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="268f0-662">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="268f0-662">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="268f0-663">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="268f0-663">- TaskPane</span></span></td>
    <td> <span data-ttu-id="268f0-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="268f0-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="268f0-665">- Seleção</span><span class="sxs-lookup"><span data-stu-id="268f0-665">- Selection</span></span><br><span data-ttu-id="268f0-666">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="268f0-666">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="268f0-667">Confira também</span><span class="sxs-lookup"><span data-stu-id="268f0-667">See also</span></span>

- [<span data-ttu-id="268f0-668">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="268f0-668">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="268f0-669">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="268f0-669">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="268f0-670">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="268f0-670">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="268f0-671">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="268f0-671">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
