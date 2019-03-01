---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 02/20/2019
localization_priority: Priority
ms.openlocfilehash: a3e9c508a5bae0e7eb660458835b9242d0602818
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/28/2019
ms.locfileid: "30199610"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="218af-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="218af-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="218af-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="218af-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="218af-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="218af-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="218af-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="218af-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="218af-108">Excel</span><span class="sxs-lookup"><span data-stu-id="218af-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="218af-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="218af-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="218af-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="218af-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="218af-113">Office Online</span></span></td>
    <td> <span data-ttu-id="218af-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-114">- TaskPane</span></span><br><span data-ttu-id="218af-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-115">
        - Content</span></span><br><span data-ttu-id="218af-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="218af-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="218af-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="218af-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="218af-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="218af-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="218af-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="218af-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="218af-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="218af-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="218af-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="218af-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-126">
        - BindingEvents</span></span><br><span data-ttu-id="218af-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-127">
        - CompressedFile</span></span><br><span data-ttu-id="218af-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-128">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-129">
        - File</span></span><br><span data-ttu-id="218af-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-130">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-132">
        - Selection</span></span><br><span data-ttu-id="218af-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-133">
        - Settings</span></span><br><span data-ttu-id="218af-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-134">
        - TableBindings</span></span><br><span data-ttu-id="218af-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-135">
        - TableCoercion</span></span><br><span data-ttu-id="218af-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-136">
        - TextBindings</span></span><br><span data-ttu-id="218af-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="218af-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-139">
        - TaskPane</span></span><br><span data-ttu-id="218af-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="218af-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="218af-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="218af-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-142">
        - BindingEvents</span></span><br><span data-ttu-id="218af-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-143">
        - CompressedFile</span></span><br><span data-ttu-id="218af-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-144">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-145">
        - File</span></span><br><span data-ttu-id="218af-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-146">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-147">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-149">
        - Selection</span></span><br><span data-ttu-id="218af-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-150">
        - Settings</span></span><br><span data-ttu-id="218af-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-151">
        - TableBindings</span></span><br><span data-ttu-id="218af-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-152">
        - TableCoercion</span></span><br><span data-ttu-id="218af-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-153">
        - TextBindings</span></span><br><span data-ttu-id="218af-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="218af-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-156">- TaskPane</span></span><br><span data-ttu-id="218af-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-157">
        - Content</span></span></td>
    <td><span data-ttu-id="218af-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-158">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="218af-159">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="218af-160">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-160">- BindingEvents</span></span><br><span data-ttu-id="218af-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-161">
        - CompressedFile</span></span><br><span data-ttu-id="218af-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-162">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-163">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-163">
        - File</span></span><br><span data-ttu-id="218af-164">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-164">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-165">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-167">
        - Selection</span></span><br><span data-ttu-id="218af-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-168">
        - Settings</span></span><br><span data-ttu-id="218af-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-169">
        - TableBindings</span></span><br><span data-ttu-id="218af-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-170">
        - TableCoercion</span></span><br><span data-ttu-id="218af-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-171">
        - TextBindings</span></span><br><span data-ttu-id="218af-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-173">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-173">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="218af-174">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-174">- TaskPane</span></span><br><span data-ttu-id="218af-175">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-175">
        - Content</span></span><br><span data-ttu-id="218af-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="218af-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-177">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-178">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="218af-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-179">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="218af-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-180">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="218af-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-181">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="218af-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-182">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="218af-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-183">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="218af-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="218af-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="218af-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-185">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="218af-186">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-186">- BindingEvents</span></span><br><span data-ttu-id="218af-187">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-187">
        - CompressedFile</span></span><br><span data-ttu-id="218af-188">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-188">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-189">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-189">
        - File</span></span><br><span data-ttu-id="218af-190">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-190">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-191">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-193">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-193">
        - Selection</span></span><br><span data-ttu-id="218af-194">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-194">
        - Settings</span></span><br><span data-ttu-id="218af-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-195">
        - TableBindings</span></span><br><span data-ttu-id="218af-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-196">
        - TableCoercion</span></span><br><span data-ttu-id="218af-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-197">
        - TextBindings</span></span><br><span data-ttu-id="218af-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-199">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="218af-199">Office for iPad</span></span></td>
    <td><span data-ttu-id="218af-200">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-200">- TaskPane</span></span><br><span data-ttu-id="218af-201">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-201">
        - Content</span></span></td>
    <td><span data-ttu-id="218af-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-202">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-203">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="218af-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-204">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="218af-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-205">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="218af-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-206">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="218af-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-207">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="218af-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-208">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="218af-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="218af-209">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="218af-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-210">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="218af-211">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-211">- BindingEvents</span></span><br><span data-ttu-id="218af-212">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-212">
        - CompressedFile</span></span><br><span data-ttu-id="218af-213">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-213">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-214">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-214">
        - File</span></span><br><span data-ttu-id="218af-215">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-215">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-216">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-216">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-217">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-217">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-218">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-218">
        - Selection</span></span><br><span data-ttu-id="218af-219">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-219">
        - Settings</span></span><br><span data-ttu-id="218af-220">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-220">
        - TableBindings</span></span><br><span data-ttu-id="218af-221">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-221">
        - TableCoercion</span></span><br><span data-ttu-id="218af-222">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-222">
        - TextBindings</span></span><br><span data-ttu-id="218af-223">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-223">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-224">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-224">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="218af-225">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-225">- TaskPane</span></span><br><span data-ttu-id="218af-226">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-226">
        - Content</span></span></td>
    <td><span data-ttu-id="218af-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-227">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="218af-228">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="218af-229">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-229">- BindingEvents</span></span><br><span data-ttu-id="218af-230">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-230">
        - CompressedFile</span></span><br><span data-ttu-id="218af-231">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-231">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-232">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-232">
        - File</span></span><br><span data-ttu-id="218af-233">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-233">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-234">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-234">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-235">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-235">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-236">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-236">
        - PdfFile</span></span><br><span data-ttu-id="218af-237">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-237">
        - Selection</span></span><br><span data-ttu-id="218af-238">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-238">
        - Settings</span></span><br><span data-ttu-id="218af-239">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-239">
        - TableBindings</span></span><br><span data-ttu-id="218af-240">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-240">
        - TableCoercion</span></span><br><span data-ttu-id="218af-241">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-241">
        - TextBindings</span></span><br><span data-ttu-id="218af-242">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-242">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-243">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-243">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="218af-244">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-244">- TaskPane</span></span><br><span data-ttu-id="218af-245">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-245">
        - Content</span></span><br><span data-ttu-id="218af-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="218af-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-247">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="218af-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-248">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="218af-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-249">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="218af-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-250">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="218af-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-251">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="218af-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-252">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="218af-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="218af-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="218af-254">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="218af-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="218af-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-256">- BindingEvents</span></span><br><span data-ttu-id="218af-257">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-257">
        - CompressedFile</span></span><br><span data-ttu-id="218af-258">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-258">
        - DocumentEvents</span></span><br><span data-ttu-id="218af-259">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-259">
        - File</span></span><br><span data-ttu-id="218af-260">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-260">
        - ImageCoercion</span></span><br><span data-ttu-id="218af-261">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-261">
        - MatrixBindings</span></span><br><span data-ttu-id="218af-262">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-262">
        - MatrixCoercion</span></span><br><span data-ttu-id="218af-263">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-263">
        - PdfFile</span></span><br><span data-ttu-id="218af-264">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-264">
        - Selection</span></span><br><span data-ttu-id="218af-265">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-265">
        - Settings</span></span><br><span data-ttu-id="218af-266">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-266">
        - TableBindings</span></span><br><span data-ttu-id="218af-267">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-267">
        - TableCoercion</span></span><br><span data-ttu-id="218af-268">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-268">
        - TextBindings</span></span><br><span data-ttu-id="218af-269">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-269">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="218af-270">Outlook</span><span class="sxs-lookup"><span data-stu-id="218af-270">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="218af-271">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-271">Platform</span></span></th>
    <th><span data-ttu-id="218af-272">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-272">Extension points</span></span></th>
    <th><span data-ttu-id="218af-273">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-273">API requirement sets</span></span></th>
    <th><span data-ttu-id="218af-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-274"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-275">Office Online</span><span class="sxs-lookup"><span data-stu-id="218af-275">Office Online</span></span></td>
    <td> <span data-ttu-id="218af-276">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-276">- Mail Read</span></span><br><span data-ttu-id="218af-277">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-277">
      - Mail Compose</span></span><br><span data-ttu-id="218af-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-278">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-279">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-280">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-281">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-282">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-283">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="218af-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-284">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="218af-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-285">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="218af-286">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-286">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-287">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-288">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-288">- Mail Read</span></span><br><span data-ttu-id="218af-289">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-289">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="218af-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-290">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-291">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-292">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-293">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="218af-294">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-294">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-295">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-295">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-296">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-296">- Mail Read</span></span><br><span data-ttu-id="218af-297">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-297">
      - Mail Compose</span></span><br><span data-ttu-id="218af-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="218af-299">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="218af-299">
      - Modules</span></span></td>
    <td> <span data-ttu-id="218af-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-300">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-304">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="218af-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-305">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="218af-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="218af-307">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-307">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-308">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-308">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-309">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-309">- Mail Read</span></span><br><span data-ttu-id="218af-310">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-310">
      - Mail Compose</span></span><br><span data-ttu-id="218af-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="218af-312">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="218af-312">
      - Modules</span></span></td>
    <td> <span data-ttu-id="218af-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-313">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-314">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-316">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="218af-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="218af-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="218af-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="218af-320">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-320">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-321">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="218af-321">Office for iOS</span></span></td>
    <td> <span data-ttu-id="218af-322">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-322">- Mail Read</span></span><br><span data-ttu-id="218af-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-324">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-326">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-327">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="218af-329">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-329">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-330">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-330">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-331">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-331">- Mail Read</span></span><br><span data-ttu-id="218af-332">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-332">
      - Mail Compose</span></span><br><span data-ttu-id="218af-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-334">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="218af-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="218af-340">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-341">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-341">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-342">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-342">- Mail Read</span></span><br><span data-ttu-id="218af-343">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="218af-343">
      - Mail Compose</span></span><br><span data-ttu-id="218af-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-345">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="218af-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="218af-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="218af-351">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-352">Office para Android</span><span class="sxs-lookup"><span data-stu-id="218af-352">Office for Android</span></span></td>
    <td> <span data-ttu-id="218af-353">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="218af-353">- Mail Read</span></span><br><span data-ttu-id="218af-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-355">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="218af-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="218af-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="218af-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="218af-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="218af-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="218af-359">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="218af-360">Não disponível</span><span class="sxs-lookup"><span data-stu-id="218af-360">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="218af-361">Word</span><span class="sxs-lookup"><span data-stu-id="218af-361">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="218af-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-362">Platform</span></span></th>
    <th><span data-ttu-id="218af-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-363">Extension points</span></span></th>
    <th><span data-ttu-id="218af-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="218af-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-365"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-366">Office Online</span><span class="sxs-lookup"><span data-stu-id="218af-366">Office Online</span></span></td>
    <td> <span data-ttu-id="218af-367">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-367">- TaskPane</span></span><br><span data-ttu-id="218af-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-368">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-369">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-370">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="218af-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-371">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="218af-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-372">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-373">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-373">- BindingEvents</span></span><br><span data-ttu-id="218af-374">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-374">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-375">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-376">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-376">
         - File</span></span><br><span data-ttu-id="218af-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-377">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-378">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-379">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-379">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-380">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-380">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-381">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-381">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-382">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-382">
         - PdfFile</span></span><br><span data-ttu-id="218af-383">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-383">
         - Selection</span></span><br><span data-ttu-id="218af-384">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-384">
         - Settings</span></span><br><span data-ttu-id="218af-385">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-385">
         - TableBindings</span></span><br><span data-ttu-id="218af-386">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-386">
         - TableCoercion</span></span><br><span data-ttu-id="218af-387">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-387">
         - TextBindings</span></span><br><span data-ttu-id="218af-388">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-388">
         - TextCoercion</span></span><br><span data-ttu-id="218af-389">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-389">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-390">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-390">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-391">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-391">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="218af-392">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="218af-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-393">- BindingEvents</span></span><br><span data-ttu-id="218af-394">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-394">
         - CompressedFile</span></span><br><span data-ttu-id="218af-395">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-395">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-396">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-396">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-397">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-397">
         - File</span></span><br><span data-ttu-id="218af-398">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-398">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-399">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-399">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-400">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-400">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-401">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-401">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-402">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-402">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-403">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-403">
         - PdfFile</span></span><br><span data-ttu-id="218af-404">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-404">
         - Selection</span></span><br><span data-ttu-id="218af-405">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-405">
         - Settings</span></span><br><span data-ttu-id="218af-406">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-406">
         - TableBindings</span></span><br><span data-ttu-id="218af-407">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-407">
         - TableCoercion</span></span><br><span data-ttu-id="218af-408">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-408">
         - TextBindings</span></span><br><span data-ttu-id="218af-409">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-409">
         - TextCoercion</span></span><br><span data-ttu-id="218af-410">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-410">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-411">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-411">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-412">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-412">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-413">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="218af-414">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="218af-415">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-415">- BindingEvents</span></span><br><span data-ttu-id="218af-416">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-416">
         - CompressedFile</span></span><br><span data-ttu-id="218af-417">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-417">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-418">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-418">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-419">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-419">
         - File</span></span><br><span data-ttu-id="218af-420">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-420">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-421">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-421">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-422">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-422">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-423">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-423">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-424">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-424">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-425">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-425">
         - PdfFile</span></span><br><span data-ttu-id="218af-426">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-426">
         - Selection</span></span><br><span data-ttu-id="218af-427">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-427">
         - Settings</span></span><br><span data-ttu-id="218af-428">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-428">
         - TableBindings</span></span><br><span data-ttu-id="218af-429">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-429">
         - TableCoercion</span></span><br><span data-ttu-id="218af-430">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-430">
         - TextBindings</span></span><br><span data-ttu-id="218af-431">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-431">
         - TextCoercion</span></span><br><span data-ttu-id="218af-432">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-432">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-433">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-433">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-434">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-434">- TaskPane</span></span><br><span data-ttu-id="218af-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-435">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-436">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="218af-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-438">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="218af-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-439">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-440">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-440">- BindingEvents</span></span><br><span data-ttu-id="218af-441">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-441">
         - CompressedFile</span></span><br><span data-ttu-id="218af-442">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-442">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-443">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-443">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-444">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-444">
         - File</span></span><br><span data-ttu-id="218af-445">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-445">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-446">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-446">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-447">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-447">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-448">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-448">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-449">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-449">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-450">
         - PdfFile</span></span><br><span data-ttu-id="218af-451">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-451">
         - Selection</span></span><br><span data-ttu-id="218af-452">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-452">
         - Settings</span></span><br><span data-ttu-id="218af-453">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-453">
         - TableBindings</span></span><br><span data-ttu-id="218af-454">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-454">
         - TableCoercion</span></span><br><span data-ttu-id="218af-455">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-455">
         - TextBindings</span></span><br><span data-ttu-id="218af-456">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-456">
         - TextCoercion</span></span><br><span data-ttu-id="218af-457">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-457">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-458">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="218af-458">Office for iPad</span></span></td>
    <td> <span data-ttu-id="218af-459">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-459">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-460">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-461">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="218af-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-462">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="218af-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="218af-463">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="218af-464">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-464">- BindingEvents</span></span><br><span data-ttu-id="218af-465">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-465">
         - CompressedFile</span></span><br><span data-ttu-id="218af-466">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-466">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-467">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-467">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-468">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-468">
         - File</span></span><br><span data-ttu-id="218af-469">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-469">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-470">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-470">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-471">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-471">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-472">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-472">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-473">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-473">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-474">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-474">
         - PdfFile</span></span><br><span data-ttu-id="218af-475">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-475">
         - Selection</span></span><br><span data-ttu-id="218af-476">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-476">
         - Settings</span></span><br><span data-ttu-id="218af-477">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-477">
         - TableBindings</span></span><br><span data-ttu-id="218af-478">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-478">
         - TableCoercion</span></span><br><span data-ttu-id="218af-479">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-479">
         - TextBindings</span></span><br><span data-ttu-id="218af-480">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-480">
         - TextCoercion</span></span><br><span data-ttu-id="218af-481">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-481">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-482">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-482">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-483">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-483">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-484">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="218af-485">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="218af-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-486">- BindingEvents</span></span><br><span data-ttu-id="218af-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-487">
         - CompressedFile</span></span><br><span data-ttu-id="218af-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-488">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-489">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-490">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-490">
         - File</span></span><br><span data-ttu-id="218af-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-491">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-492">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-493">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-494">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-495">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-496">
         - PdfFile</span></span><br><span data-ttu-id="218af-497">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-497">
         - Selection</span></span><br><span data-ttu-id="218af-498">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-498">
         - Settings</span></span><br><span data-ttu-id="218af-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-499">
         - TableBindings</span></span><br><span data-ttu-id="218af-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-500">
         - TableCoercion</span></span><br><span data-ttu-id="218af-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-501">
         - TextBindings</span></span><br><span data-ttu-id="218af-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-502">
         - TextCoercion</span></span><br><span data-ttu-id="218af-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-503">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-504">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-504">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-505">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-505">- TaskPane</span></span><br><span data-ttu-id="218af-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="218af-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="218af-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="218af-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="218af-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="218af-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="218af-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="218af-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="218af-511">- BindingEvents</span></span><br><span data-ttu-id="218af-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-512">
         - CompressedFile</span></span><br><span data-ttu-id="218af-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="218af-513">
         - CustomXmlParts</span></span><br><span data-ttu-id="218af-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-514">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-515">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-515">
         - File</span></span><br><span data-ttu-id="218af-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-516">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-517">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="218af-518">
         - MatrixBindings</span></span><br><span data-ttu-id="218af-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-519">
         - MatrixCoercion</span></span><br><span data-ttu-id="218af-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-520">
         - OoxmlCoercion</span></span><br><span data-ttu-id="218af-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-521">
         - PdfFile</span></span><br><span data-ttu-id="218af-522">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-522">
         - Selection</span></span><br><span data-ttu-id="218af-523">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-523">
         - Settings</span></span><br><span data-ttu-id="218af-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="218af-524">
         - TableBindings</span></span><br><span data-ttu-id="218af-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-525">
         - TableCoercion</span></span><br><span data-ttu-id="218af-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="218af-526">
         - TextBindings</span></span><br><span data-ttu-id="218af-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-527">
         - TextCoercion</span></span><br><span data-ttu-id="218af-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="218af-528">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="218af-529">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="218af-529">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="218af-530">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-530">Platform</span></span></th>
    <th><span data-ttu-id="218af-531">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-531">Extension points</span></span></th>
    <th><span data-ttu-id="218af-532">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-532">API requirement sets</span></span></th>
    <th><span data-ttu-id="218af-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-533"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-534">Office Online</span><span class="sxs-lookup"><span data-stu-id="218af-534">Office Online</span></span></td>
    <td> <span data-ttu-id="218af-535">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-535">- Content</span></span><br><span data-ttu-id="218af-536">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-536">
         - TaskPane</span></span><br><span data-ttu-id="218af-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-537">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-538">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-539">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-539">- ActiveView</span></span><br><span data-ttu-id="218af-540">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-540">
         - CompressedFile</span></span><br><span data-ttu-id="218af-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-541">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-542">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-542">
         - File</span></span><br><span data-ttu-id="218af-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-543">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-544">
         - PdfFile</span></span><br><span data-ttu-id="218af-545">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-545">
         - Selection</span></span><br><span data-ttu-id="218af-546">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-546">
         - Settings</span></span><br><span data-ttu-id="218af-547">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-547">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-548">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-548">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-549">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-549">- Content</span></span><br><span data-ttu-id="218af-550">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-550">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="218af-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="218af-551">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="218af-552">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-552">- ActiveView</span></span><br><span data-ttu-id="218af-553">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-553">
         - CompressedFile</span></span><br><span data-ttu-id="218af-554">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-554">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-555">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-555">
         - File</span></span><br><span data-ttu-id="218af-556">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-556">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-557">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-557">
         - PdfFile</span></span><br><span data-ttu-id="218af-558">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-558">
         - Selection</span></span><br><span data-ttu-id="218af-559">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-559">
         - Settings</span></span><br><span data-ttu-id="218af-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-560">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-561">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-561">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-562">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-562">- Content</span></span><br><span data-ttu-id="218af-563">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-563">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="218af-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="218af-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-565">- ActiveView</span></span><br><span data-ttu-id="218af-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-566">
         - CompressedFile</span></span><br><span data-ttu-id="218af-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-567">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-568">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-568">
         - File</span></span><br><span data-ttu-id="218af-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-569">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-570">
         - PdfFile</span></span><br><span data-ttu-id="218af-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-571">
         - Selection</span></span><br><span data-ttu-id="218af-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-572">
         - Settings</span></span><br><span data-ttu-id="218af-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-573">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-574">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-574">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-575">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-575">- Content</span></span><br><span data-ttu-id="218af-576">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-576">
         - TaskPane</span></span><br><span data-ttu-id="218af-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-577">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-578">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-579">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-579">- ActiveView</span></span><br><span data-ttu-id="218af-580">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-580">
         - CompressedFile</span></span><br><span data-ttu-id="218af-581">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-581">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-582">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-582">
         - File</span></span><br><span data-ttu-id="218af-583">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-583">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-584">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-584">
         - PdfFile</span></span><br><span data-ttu-id="218af-585">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-585">
         - Selection</span></span><br><span data-ttu-id="218af-586">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-586">
         - Settings</span></span><br><span data-ttu-id="218af-587">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-587">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-588">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="218af-588">Office for iPad</span></span></td>
    <td> <span data-ttu-id="218af-589">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-589">- Content</span></span><br><span data-ttu-id="218af-590">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-590">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="218af-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-592">- ActiveView</span></span><br><span data-ttu-id="218af-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-593">
         - CompressedFile</span></span><br><span data-ttu-id="218af-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-594">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-595">
         - File</span></span><br><span data-ttu-id="218af-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-596">
         - PdfFile</span></span><br><span data-ttu-id="218af-597">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-597">
         - Selection</span></span><br><span data-ttu-id="218af-598">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-598">
         - Settings</span></span><br><span data-ttu-id="218af-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-599">
         - TextCoercion</span></span><br><span data-ttu-id="218af-600">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-600">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-601">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-601">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-602">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-602">- Content</span></span><br><span data-ttu-id="218af-603">
         - TaskPane/td></span><span class="sxs-lookup"><span data-stu-id="218af-603">
         - TaskPane/td></span></span> <td> <span data-ttu-id="218af-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="218af-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="218af-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-605">- ActiveView</span></span><br><span data-ttu-id="218af-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-606">
         - CompressedFile</span></span><br><span data-ttu-id="218af-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-607">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-608">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-608">
         - File</span></span><br><span data-ttu-id="218af-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-609">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-610">
         - PdfFile</span></span><br><span data-ttu-id="218af-611">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-611">
         - Selection</span></span><br><span data-ttu-id="218af-612">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-612">
         - Settings</span></span><br><span data-ttu-id="218af-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-613">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-614">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="218af-614">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="218af-615">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-615">- Content</span></span><br><span data-ttu-id="218af-616">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-616">
         - TaskPane</span></span><br><span data-ttu-id="218af-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-617">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="218af-619">- ActiveView</span></span><br><span data-ttu-id="218af-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="218af-620">
         - CompressedFile</span></span><br><span data-ttu-id="218af-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-621">
         - DocumentEvents</span></span><br><span data-ttu-id="218af-622">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="218af-622">
         - File</span></span><br><span data-ttu-id="218af-623">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-623">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="218af-624">
         - PdfFile</span></span><br><span data-ttu-id="218af-625">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-625">
         - Selection</span></span><br><span data-ttu-id="218af-626">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-626">
         - Settings</span></span><br><span data-ttu-id="218af-627">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-627">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="218af-628">OneNote</span><span class="sxs-lookup"><span data-stu-id="218af-628">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="218af-629">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-629">Platform</span></span></th>
    <th><span data-ttu-id="218af-630">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-630">Extension points</span></span></th>
    <th><span data-ttu-id="218af-631">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-631">API requirement sets</span></span></th>
    <th><span data-ttu-id="218af-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-632"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-633">Office Online</span><span class="sxs-lookup"><span data-stu-id="218af-633">Office Online</span></span></td>
    <td> <span data-ttu-id="218af-634">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="218af-634">- Content</span></span><br><span data-ttu-id="218af-635">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-635">
         - TaskPane</span></span><br><span data-ttu-id="218af-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="218af-636">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="218af-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-637">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="218af-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-638">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-639">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="218af-639">- DocumentEvents</span></span><br><span data-ttu-id="218af-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="218af-641">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-641">
         - ImageCoercion</span></span><br><span data-ttu-id="218af-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="218af-642">
         - Settings</span></span><br><span data-ttu-id="218af-643">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-643">
         - TextCoercion</span></span></td>
  </tr>
</table><span data-ttu-id="218af-644">
\*&ast; – Adicionado com atualizações pós-lançamento.*

</span><span class="sxs-lookup"><span data-stu-id="218af-644">
\*&ast; - Added with post-release updates.*

</span></span><br/>

## <a name="project"></a><span data-ttu-id="218af-645">Projeto</span><span class="sxs-lookup"><span data-stu-id="218af-645">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="218af-646">Plataforma</span><span class="sxs-lookup"><span data-stu-id="218af-646">Platform</span></span></th>
    <th><span data-ttu-id="218af-647">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="218af-647">Extension points</span></span></th>
    <th><span data-ttu-id="218af-648">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="218af-648">API requirement sets</span></span></th>
    <th><span data-ttu-id="218af-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="218af-649"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-650">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-650">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-651">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-652">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-653">- Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-653">- Selection</span></span><br><span data-ttu-id="218af-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-654">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-655">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-655">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-656">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-656">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-657">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-658">- Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-658">- Selection</span></span><br><span data-ttu-id="218af-659">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-659">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="218af-660">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="218af-660">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="218af-661">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="218af-661">- TaskPane</span></span></td>
    <td> <span data-ttu-id="218af-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="218af-662">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="218af-663">- Seleção</span><span class="sxs-lookup"><span data-stu-id="218af-663">- Selection</span></span><br><span data-ttu-id="218af-664">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="218af-664">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="218af-665">Confira também</span><span class="sxs-lookup"><span data-stu-id="218af-665">See also</span></span>

- [<span data-ttu-id="218af-666">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="218af-666">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="218af-667">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="218af-667">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="218af-668">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="218af-668">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="218af-669">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="218af-669">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
