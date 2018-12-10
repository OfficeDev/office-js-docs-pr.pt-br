---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 11/07/2018
ms.openlocfilehash: c601eac5ed3fcad76b63fff5ae6eeadb7662c8b7
ms.sourcegitcommit: 0adc31ceaba92cb15dc6430c00fe7a96c107c9de
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/09/2018
ms.locfileid: "27210102"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e37eb-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e37eb-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e37eb-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="e37eb-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="e37eb-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e37eb-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e37eb-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="e37eb-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e37eb-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e37eb-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e37eb-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e37eb-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e37eb-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e37eb-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="e37eb-113">Office Online</span></span></td>
    <td> <span data-ttu-id="e37eb-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-114">- TaskPane</span></span><br><span data-ttu-id="e37eb-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-115">
        - Content</span></span><br><span data-ttu-id="e37eb-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="e37eb-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e37eb-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-126">
        - BindingEvents</span></span><br><span data-ttu-id="e37eb-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-127">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-128">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-129">
        - File</span></span><br><span data-ttu-id="e37eb-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-130">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-132">
        - Selection</span></span><br><span data-ttu-id="e37eb-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-133">
        - Settings</span></span><br><span data-ttu-id="e37eb-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-134">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-135">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-136">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e37eb-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-139">
        - TaskPane</span></span><br><span data-ttu-id="e37eb-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e37eb-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-142">
        - BindingEvents</span></span><br><span data-ttu-id="e37eb-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-143">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-144">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-145">
        - File</span></span><br><span data-ttu-id="e37eb-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-146">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-147">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-149">
        - Selection</span></span><br><span data-ttu-id="e37eb-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-150">
        - Settings</span></span><br><span data-ttu-id="e37eb-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-151">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-152">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-153">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e37eb-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-156">- TaskPane</span></span><br><span data-ttu-id="e37eb-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-157">
        - Content</span></span><br><span data-ttu-id="e37eb-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e37eb-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-168">- BindingEvents</span></span><br><span data-ttu-id="e37eb-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-169">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-170">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-171">
        - File</span></span><br><span data-ttu-id="e37eb-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-172">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-173">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-175">
        - Selection</span></span><br><span data-ttu-id="e37eb-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-176">
        - Settings</span></span><br><span data-ttu-id="e37eb-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-177">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-178">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-179">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="e37eb-182">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-182">- TaskPane</span></span><br><span data-ttu-id="e37eb-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-183">
        - Content</span></span><br><span data-ttu-id="e37eb-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e37eb-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-194">- BindingEvents</span></span><br><span data-ttu-id="e37eb-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-195">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-196">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-197">
        - File</span></span><br><span data-ttu-id="e37eb-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-198">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-199">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-201">
        - Selection</span></span><br><span data-ttu-id="e37eb-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-202">
        - Settings</span></span><br><span data-ttu-id="e37eb-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-203">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-204">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-205">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-207">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="e37eb-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="e37eb-208">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-208">- TaskPane</span></span><br><span data-ttu-id="e37eb-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-209">
        - Content</span></span></td>
    <td><span data-ttu-id="e37eb-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-219">- BindingEvents</span></span><br><span data-ttu-id="e37eb-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-220">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-221">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-222">
        - File</span></span><br><span data-ttu-id="e37eb-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-223">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-224">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-226">
        - Selection</span></span><br><span data-ttu-id="e37eb-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-227">
        - Settings</span></span><br><span data-ttu-id="e37eb-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-228">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-229">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-230">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-232">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e37eb-233">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-233">- TaskPane</span></span><br><span data-ttu-id="e37eb-234">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-234">
        - Content</span></span><br><span data-ttu-id="e37eb-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e37eb-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-245">- BindingEvents</span></span><br><span data-ttu-id="e37eb-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-246">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-247">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-248">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-248">
        - File</span></span><br><span data-ttu-id="e37eb-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-249">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-250">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-252">
        - PdfFile</span></span><br><span data-ttu-id="e37eb-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-253">
        - Selection</span></span><br><span data-ttu-id="e37eb-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-254">
        - Settings</span></span><br><span data-ttu-id="e37eb-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-255">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-256">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-257">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-259">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="e37eb-260">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-260">- TaskPane</span></span><br><span data-ttu-id="e37eb-261">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-261">
        - Content</span></span><br><span data-ttu-id="e37eb-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e37eb-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e37eb-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e37eb-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e37eb-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e37eb-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e37eb-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e37eb-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e37eb-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e37eb-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e37eb-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-272">- BindingEvents</span></span><br><span data-ttu-id="e37eb-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-273">
        - CompressedFile</span></span><br><span data-ttu-id="e37eb-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-274">
        - DocumentEvents</span></span><br><span data-ttu-id="e37eb-275">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-275">
        - File</span></span><br><span data-ttu-id="e37eb-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-276">
        - ImageCoercion</span></span><br><span data-ttu-id="e37eb-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-277">
        - MatrixBindings</span></span><br><span data-ttu-id="e37eb-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-279">
        - PdfFile</span></span><br><span data-ttu-id="e37eb-280">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-280">
        - Selection</span></span><br><span data-ttu-id="e37eb-281">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-281">
        - Settings</span></span><br><span data-ttu-id="e37eb-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-282">
        - TableBindings</span></span><br><span data-ttu-id="e37eb-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-283">
        - TableCoercion</span></span><br><span data-ttu-id="e37eb-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-284">
        - TextBindings</span></span><br><span data-ttu-id="e37eb-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e37eb-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="e37eb-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e37eb-287">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-287">Platform</span></span></th>
    <th><span data-ttu-id="e37eb-288">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-288">Extension points</span></span></th>
    <th><span data-ttu-id="e37eb-289">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="e37eb-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="e37eb-291">Office Online</span></span></td>
    <td> <span data-ttu-id="e37eb-292">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-292">- Mail Read</span></span><br><span data-ttu-id="e37eb-293">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-293">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e37eb-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e37eb-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e37eb-302">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-303">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-304">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-304">- Mail Read</span></span><br><span data-ttu-id="e37eb-305">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-305">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e37eb-311">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-312">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-313">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-313">- Mail Read</span></span><br><span data-ttu-id="e37eb-314">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-314">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e37eb-316">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="e37eb-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e37eb-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e37eb-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e37eb-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e37eb-324">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-325">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-326">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-326">- Mail Read</span></span><br><span data-ttu-id="e37eb-327">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-327">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e37eb-329">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="e37eb-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e37eb-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e37eb-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e37eb-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e37eb-337">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-338">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e37eb-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e37eb-339">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-339">- Mail Read</span></span><br><span data-ttu-id="e37eb-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e37eb-346">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-347">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-348">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-348">- Mail Read</span></span><br><span data-ttu-id="e37eb-349">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-349">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e37eb-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e37eb-357">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-358">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-359">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-359">- Mail Read</span></span><br><span data-ttu-id="e37eb-360">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-360">
      - Mail Compose</span></span><br><span data-ttu-id="e37eb-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e37eb-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e37eb-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e37eb-369">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-370">Office para Android</span><span class="sxs-lookup"><span data-stu-id="e37eb-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="e37eb-371">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e37eb-371">- Mail Read</span></span><br><span data-ttu-id="e37eb-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e37eb-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e37eb-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e37eb-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e37eb-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e37eb-378">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e37eb-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e37eb-379">Word</span><span class="sxs-lookup"><span data-stu-id="e37eb-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e37eb-380">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-380">Platform</span></span></th>
    <th><span data-ttu-id="e37eb-381">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-381">Extension points</span></span></th>
    <th><span data-ttu-id="e37eb-382">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="e37eb-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="e37eb-384">Office Online</span></span></td>
    <td> <span data-ttu-id="e37eb-385">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-385">- TaskPane</span></span><br><span data-ttu-id="e37eb-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-391">- BindingEvents</span></span><br><span data-ttu-id="e37eb-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-393">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-394">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-394">
         - File</span></span><br><span data-ttu-id="e37eb-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-396">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-397">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-400">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-401">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-401">
         - Selection</span></span><br><span data-ttu-id="e37eb-402">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-402">
         - Settings</span></span><br><span data-ttu-id="e37eb-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-403">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-404">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-405">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-406">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-408">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-409">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-411">- BindingEvents</span></span><br><span data-ttu-id="e37eb-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-412">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-414">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-415">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-415">
         - File</span></span><br><span data-ttu-id="e37eb-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-417">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-418">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-421">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-422">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-422">
         - Selection</span></span><br><span data-ttu-id="e37eb-423">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-423">
         - Settings</span></span><br><span data-ttu-id="e37eb-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-424">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-425">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-426">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-427">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-429">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-430">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-430">- TaskPane</span></span><br><span data-ttu-id="e37eb-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-436">- BindingEvents</span></span><br><span data-ttu-id="e37eb-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-437">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-439">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-440">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-440">
         - File</span></span><br><span data-ttu-id="e37eb-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-442">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-443">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-446">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-447">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-447">
         - Selection</span></span><br><span data-ttu-id="e37eb-448">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-448">
         - Settings</span></span><br><span data-ttu-id="e37eb-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-449">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-450">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-451">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-452">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-454">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-455">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-455">- TaskPane</span></span><br><span data-ttu-id="e37eb-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-461">- BindingEvents</span></span><br><span data-ttu-id="e37eb-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-462">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-464">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-465">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-465">
         - File</span></span><br><span data-ttu-id="e37eb-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-467">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-468">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-471">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-472">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-472">
         - Selection</span></span><br><span data-ttu-id="e37eb-473">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-473">
         - Settings</span></span><br><span data-ttu-id="e37eb-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-474">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-475">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-476">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-477">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-479">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="e37eb-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e37eb-480">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e37eb-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e37eb-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-485">- BindingEvents</span></span><br><span data-ttu-id="e37eb-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-486">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-488">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-489">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-489">
         - File</span></span><br><span data-ttu-id="e37eb-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-491">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-492">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-495">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-496">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-496">
         - Selection</span></span><br><span data-ttu-id="e37eb-497">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-497">
         - Settings</span></span><br><span data-ttu-id="e37eb-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-498">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-499">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-500">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-501">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-503">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-504">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-504">- TaskPane</span></span><br><span data-ttu-id="e37eb-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e37eb-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e37eb-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-510">- BindingEvents</span></span><br><span data-ttu-id="e37eb-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-511">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-513">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-514">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-514">
         - File</span></span><br><span data-ttu-id="e37eb-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-516">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-517">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-520">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-521">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-521">
         - Selection</span></span><br><span data-ttu-id="e37eb-522">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-522">
         - Settings</span></span><br><span data-ttu-id="e37eb-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-523">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-524">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-525">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-526">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-528">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-529">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-529">- TaskPane</span></span><br><span data-ttu-id="e37eb-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e37eb-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e37eb-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e37eb-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e37eb-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e37eb-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-535">- BindingEvents</span></span><br><span data-ttu-id="e37eb-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-536">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e37eb-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="e37eb-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-538">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-539">
         - File</span></span><br><span data-ttu-id="e37eb-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-541">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-542">
         - MatrixBindings</span></span><br><span data-ttu-id="e37eb-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="e37eb-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e37eb-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-545">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-546">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-546">
         - Selection</span></span><br><span data-ttu-id="e37eb-547">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-547">
         - Settings</span></span><br><span data-ttu-id="e37eb-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-548">
         - TableBindings</span></span><br><span data-ttu-id="e37eb-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-549">
         - TableCoercion</span></span><br><span data-ttu-id="e37eb-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e37eb-550">
         - TextBindings</span></span><br><span data-ttu-id="e37eb-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-551">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e37eb-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e37eb-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e37eb-554">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-554">Platform</span></span></th>
    <th><span data-ttu-id="e37eb-555">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-555">Extension points</span></span></th>
    <th><span data-ttu-id="e37eb-556">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="e37eb-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="e37eb-558">Office Online</span></span></td>
    <td> <span data-ttu-id="e37eb-559">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-559">- Content</span></span><br><span data-ttu-id="e37eb-560">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-560">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-563">- ActiveView</span></span><br><span data-ttu-id="e37eb-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-564">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-565">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-566">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-566">
         - File</span></span><br><span data-ttu-id="e37eb-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-567">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-568">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-569">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-569">
         - Selection</span></span><br><span data-ttu-id="e37eb-570">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-570">
         - Settings</span></span><br><span data-ttu-id="e37eb-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-572">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-573">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-573">- Content</span></span><br><span data-ttu-id="e37eb-574">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e37eb-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e37eb-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e37eb-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-576">- ActiveView</span></span><br><span data-ttu-id="e37eb-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-577">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-578">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-579">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-579">
         - File</span></span><br><span data-ttu-id="e37eb-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-580">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-581">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-582">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-582">
         - Selection</span></span><br><span data-ttu-id="e37eb-583">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-583">
         - Settings</span></span><br><span data-ttu-id="e37eb-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-585">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-586">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-586">- Content</span></span><br><span data-ttu-id="e37eb-587">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-587">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-590">- ActiveView</span></span><br><span data-ttu-id="e37eb-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-591">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-592">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-593">
         - File</span></span><br><span data-ttu-id="e37eb-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-594">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-595">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-596">
         - Selection</span></span><br><span data-ttu-id="e37eb-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-597">
         - Settings</span></span><br><span data-ttu-id="e37eb-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-599">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-600">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-600">- Content</span></span><br><span data-ttu-id="e37eb-601">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-601">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-604">- ActiveView</span></span><br><span data-ttu-id="e37eb-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-605">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-606">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-607">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-607">
         - File</span></span><br><span data-ttu-id="e37eb-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-608">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-609">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-610">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-610">
         - Selection</span></span><br><span data-ttu-id="e37eb-611">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-611">
         - Settings</span></span><br><span data-ttu-id="e37eb-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-613">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="e37eb-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e37eb-614">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-614">- Content</span></span><br><span data-ttu-id="e37eb-615">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e37eb-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-617">- ActiveView</span></span><br><span data-ttu-id="e37eb-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-618">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-619">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-620">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-620">
         - File</span></span><br><span data-ttu-id="e37eb-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-621">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-622">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-622">
         - Selection</span></span><br><span data-ttu-id="e37eb-623">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-623">
         - Settings</span></span><br><span data-ttu-id="e37eb-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-624">
         - TextCoercion</span></span><br><span data-ttu-id="e37eb-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-626">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-627">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-627">- Content</span></span><br><span data-ttu-id="e37eb-628">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-628">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-631">- ActiveView</span></span><br><span data-ttu-id="e37eb-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-632">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-633">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-634">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-634">
         - File</span></span><br><span data-ttu-id="e37eb-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-635">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-636">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-637">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-637">
         - Selection</span></span><br><span data-ttu-id="e37eb-638">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-638">
         - Settings</span></span><br><span data-ttu-id="e37eb-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-640">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e37eb-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e37eb-641">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-641">- Content</span></span><br><span data-ttu-id="e37eb-642">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-642">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e37eb-645">- ActiveView</span></span><br><span data-ttu-id="e37eb-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-646">
         - CompressedFile</span></span><br><span data-ttu-id="e37eb-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-647">
         - DocumentEvents</span></span><br><span data-ttu-id="e37eb-648">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e37eb-648">
         - File</span></span><br><span data-ttu-id="e37eb-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-649">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e37eb-650">
         - PdfFile</span></span><br><span data-ttu-id="e37eb-651">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-651">
         - Selection</span></span><br><span data-ttu-id="e37eb-652">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-652">
         - Settings</span></span><br><span data-ttu-id="e37eb-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e37eb-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="e37eb-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e37eb-655">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-655">Platform</span></span></th>
    <th><span data-ttu-id="e37eb-656">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-656">Extension points</span></span></th>
    <th><span data-ttu-id="e37eb-657">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="e37eb-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="e37eb-659">Office Online</span></span></td>
    <td> <span data-ttu-id="e37eb-660">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e37eb-660">- Content</span></span><br><span data-ttu-id="e37eb-661">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-661">
         - TaskPane</span></span><br><span data-ttu-id="e37eb-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e37eb-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e37eb-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e37eb-665">- DocumentEvents</span></span><br><span data-ttu-id="e37eb-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="e37eb-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-667">
         - ImageCoercion</span></span><br><span data-ttu-id="e37eb-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e37eb-668">
         - Settings</span></span><br><span data-ttu-id="e37eb-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e37eb-670">Project</span><span class="sxs-lookup"><span data-stu-id="e37eb-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e37eb-671">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e37eb-671">Platform</span></span></th>
    <th><span data-ttu-id="e37eb-672">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e37eb-672">Extension points</span></span></th>
    <th><span data-ttu-id="e37eb-673">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="e37eb-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e37eb-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-675">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-678">- Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-678">- Selection</span></span><br><span data-ttu-id="e37eb-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-680">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-681">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-683">- Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-683">- Selection</span></span><br><span data-ttu-id="e37eb-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e37eb-685">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e37eb-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e37eb-686">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="e37eb-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e37eb-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e37eb-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e37eb-688">- Seleção</span><span class="sxs-lookup"><span data-stu-id="e37eb-688">- Selection</span></span><br><span data-ttu-id="e37eb-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e37eb-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e37eb-690">Confira também</span><span class="sxs-lookup"><span data-stu-id="e37eb-690">See also</span></span>

- [<span data-ttu-id="e37eb-691">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e37eb-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e37eb-692">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="e37eb-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e37eb-693">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="e37eb-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e37eb-694">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="e37eb-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
