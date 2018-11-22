---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 11/07/2018
ms.openlocfilehash: c3da40be21c0e569028dd10e93e33760ba2bd39d
ms.sourcegitcommit: 3e84d616e69f39eeeeea773f2431e7d674c4a9f5
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/22/2018
ms.locfileid: "26644750"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="83756-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="83756-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="83756-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="83756-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="83756-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="83756-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="83756-p102">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="83756-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="83756-108">Excel</span><span class="sxs-lookup"><span data-stu-id="83756-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="83756-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="83756-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="83756-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="83756-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="83756-113">Office Online</span></span></td>
    <td> <span data-ttu-id="83756-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-114">- TaskPane</span></span><br><span data-ttu-id="83756-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-115">
        - Content</span></span><br><span data-ttu-id="83756-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="83756-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="83756-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-126">
        - BindingEvents</span></span><br><span data-ttu-id="83756-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-127">
        - CompressedFile</span></span><br><span data-ttu-id="83756-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-128">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-129">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-129">
        - File</span></span><br><span data-ttu-id="83756-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-130">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-132">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-132">
        - Selection</span></span><br><span data-ttu-id="83756-133">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-133">
        - Settings</span></span><br><span data-ttu-id="83756-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-134">
        - TableBindings</span></span><br><span data-ttu-id="83756-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-135">
        - TableCoercion</span></span><br><span data-ttu-id="83756-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-136">
        - TextBindings</span></span><br><span data-ttu-id="83756-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-138">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="83756-139">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-139">
        - TaskPane</span></span><br><span data-ttu-id="83756-140">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="83756-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-142">
        - BindingEvents</span></span><br><span data-ttu-id="83756-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-143">
        - CompressedFile</span></span><br><span data-ttu-id="83756-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-144">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-145">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-145">
        - File</span></span><br><span data-ttu-id="83756-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-146">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-147">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-149">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-149">
        - Selection</span></span><br><span data-ttu-id="83756-150">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-150">
        - Settings</span></span><br><span data-ttu-id="83756-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-151">
        - TableBindings</span></span><br><span data-ttu-id="83756-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-152">
        - TableCoercion</span></span><br><span data-ttu-id="83756-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-153">
        - TextBindings</span></span><br><span data-ttu-id="83756-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-155">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="83756-156">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-156">- TaskPane</span></span><br><span data-ttu-id="83756-157">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-157">
        - Content</span></span><br><span data-ttu-id="83756-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="83756-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-168">- BindingEvents</span></span><br><span data-ttu-id="83756-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-169">
        - CompressedFile</span></span><br><span data-ttu-id="83756-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-170">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-171">
        - File</span></span><br><span data-ttu-id="83756-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-172">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-173">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-175">
        - Selection</span></span><br><span data-ttu-id="83756-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-176">
        - Settings</span></span><br><span data-ttu-id="83756-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-177">
        - TableBindings</span></span><br><span data-ttu-id="83756-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-178">
        - TableCoercion</span></span><br><span data-ttu-id="83756-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-179">
        - TextBindings</span></span><br><span data-ttu-id="83756-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="83756-182">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-182">- TaskPane</span></span><br><span data-ttu-id="83756-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-183">
        - Content</span></span><br><span data-ttu-id="83756-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="83756-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-194">- BindingEvents</span></span><br><span data-ttu-id="83756-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-195">
        - CompressedFile</span></span><br><span data-ttu-id="83756-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-196">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-197">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-197">
        - File</span></span><br><span data-ttu-id="83756-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-198">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-199">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-201">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-201">
        - Selection</span></span><br><span data-ttu-id="83756-202">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-202">
        - Settings</span></span><br><span data-ttu-id="83756-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-203">
        - TableBindings</span></span><br><span data-ttu-id="83756-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-204">
        - TableCoercion</span></span><br><span data-ttu-id="83756-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-205">
        - TextBindings</span></span><br><span data-ttu-id="83756-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-207">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="83756-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="83756-208">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-208">- TaskPane</span></span><br><span data-ttu-id="83756-209">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-209">
        - Content</span></span></td>
    <td><span data-ttu-id="83756-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-219">- BindingEvents</span></span><br><span data-ttu-id="83756-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-220">
        - CompressedFile</span></span><br><span data-ttu-id="83756-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-221">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-222">
        - File</span></span><br><span data-ttu-id="83756-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-223">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-224">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-226">
        - Selection</span></span><br><span data-ttu-id="83756-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-227">
        - Settings</span></span><br><span data-ttu-id="83756-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-228">
        - TableBindings</span></span><br><span data-ttu-id="83756-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-229">
        - TableCoercion</span></span><br><span data-ttu-id="83756-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-230">
        - TextBindings</span></span><br><span data-ttu-id="83756-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-232">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="83756-233">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-233">- TaskPane</span></span><br><span data-ttu-id="83756-234">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-234">
        - Content</span></span><br><span data-ttu-id="83756-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="83756-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-245">- BindingEvents</span></span><br><span data-ttu-id="83756-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-246">
        - CompressedFile</span></span><br><span data-ttu-id="83756-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-247">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-248">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-248">
        - File</span></span><br><span data-ttu-id="83756-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-249">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-250">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-252">
        - PdfFile</span></span><br><span data-ttu-id="83756-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-253">
        - Selection</span></span><br><span data-ttu-id="83756-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-254">
        - Settings</span></span><br><span data-ttu-id="83756-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-255">
        - TableBindings</span></span><br><span data-ttu-id="83756-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-256">
        - TableCoercion</span></span><br><span data-ttu-id="83756-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-257">
        - TextBindings</span></span><br><span data-ttu-id="83756-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-259">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="83756-260">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-260">- TaskPane</span></span><br><span data-ttu-id="83756-261">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-261">
        - Content</span></span><br><span data-ttu-id="83756-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="83756-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="83756-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="83756-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="83756-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="83756-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="83756-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="83756-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="83756-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="83756-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="83756-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="83756-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-272">- BindingEvents</span></span><br><span data-ttu-id="83756-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-273">
        - CompressedFile</span></span><br><span data-ttu-id="83756-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-274">
        - DocumentEvents</span></span><br><span data-ttu-id="83756-275">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-275">
        - File</span></span><br><span data-ttu-id="83756-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-276">
        - ImageCoercion</span></span><br><span data-ttu-id="83756-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-277">
        - MatrixBindings</span></span><br><span data-ttu-id="83756-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="83756-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-279">
        - PdfFile</span></span><br><span data-ttu-id="83756-280">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-280">
        - Selection</span></span><br><span data-ttu-id="83756-281">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-281">
        - Settings</span></span><br><span data-ttu-id="83756-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-282">
        - TableBindings</span></span><br><span data-ttu-id="83756-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-283">
        - TableCoercion</span></span><br><span data-ttu-id="83756-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-284">
        - TextBindings</span></span><br><span data-ttu-id="83756-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="83756-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="83756-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="83756-287">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-287">Platform</span></span></th>
    <th><span data-ttu-id="83756-288">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-288">Extension points</span></span></th>
    <th><span data-ttu-id="83756-289">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="83756-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="83756-291">Office Online</span></span></td>
    <td> <span data-ttu-id="83756-292">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-292">- Mail Read</span></span><br><span data-ttu-id="83756-293">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-293">
      - Mail Compose</span></span><br><span data-ttu-id="83756-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="83756-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="83756-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="83756-302">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-303">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-304">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-304">- Mail Read</span></span><br><span data-ttu-id="83756-305">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-305">
      - Mail Compose</span></span><br><span data-ttu-id="83756-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="83756-311">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-312">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-313">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-313">- Mail Read</span></span><br><span data-ttu-id="83756-314">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-314">
      - Mail Compose</span></span><br><span data-ttu-id="83756-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="83756-316">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="83756-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="83756-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="83756-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="83756-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="83756-324">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-325">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-326">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-326">- Mail Read</span></span><br><span data-ttu-id="83756-327">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-327">
      - Mail Compose</span></span><br><span data-ttu-id="83756-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="83756-329">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="83756-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="83756-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="83756-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="83756-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="83756-337">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-338">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="83756-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="83756-339">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-339">- Mail Read</span></span><br><span data-ttu-id="83756-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="83756-346">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-347">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-348">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-348">- Mail Read</span></span><br><span data-ttu-id="83756-349">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-349">
      - Mail Compose</span></span><br><span data-ttu-id="83756-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="83756-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="83756-357">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-358">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-359">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-359">- Mail Read</span></span><br><span data-ttu-id="83756-360">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="83756-360">
      - Mail Compose</span></span><br><span data-ttu-id="83756-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="83756-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="83756-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="83756-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="83756-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="83756-369">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-370">Office para Android</span><span class="sxs-lookup"><span data-stu-id="83756-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="83756-371">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="83756-371">- Mail Read</span></span><br><span data-ttu-id="83756-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="83756-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="83756-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="83756-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="83756-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="83756-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="83756-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="83756-378">Não disponível</span><span class="sxs-lookup"><span data-stu-id="83756-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="83756-379">Word</span><span class="sxs-lookup"><span data-stu-id="83756-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="83756-380">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-380">Platform</span></span></th>
    <th><span data-ttu-id="83756-381">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-381">Extension points</span></span></th>
    <th><span data-ttu-id="83756-382">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="83756-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="83756-384">Office Online</span></span></td>
    <td> <span data-ttu-id="83756-385">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-385">- TaskPane</span></span><br><span data-ttu-id="83756-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-391">- BindingEvents</span></span><br><span data-ttu-id="83756-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-393">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-394">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-394">
         - File</span></span><br><span data-ttu-id="83756-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-396">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-397">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-400">
         - PdfFile</span></span><br><span data-ttu-id="83756-401">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-401">
         - Selection</span></span><br><span data-ttu-id="83756-402">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-402">
         - Settings</span></span><br><span data-ttu-id="83756-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-403">
         - TableBindings</span></span><br><span data-ttu-id="83756-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-404">
         - TableCoercion</span></span><br><span data-ttu-id="83756-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-405">
         - TextBindings</span></span><br><span data-ttu-id="83756-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-406">
         - TextCoercion</span></span><br><span data-ttu-id="83756-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-408">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-409">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-411">- BindingEvents</span></span><br><span data-ttu-id="83756-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-412">
         - CompressedFile</span></span><br><span data-ttu-id="83756-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-414">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-415">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-415">
         - File</span></span><br><span data-ttu-id="83756-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-417">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-418">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-421">
         - PdfFile</span></span><br><span data-ttu-id="83756-422">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-422">
         - Selection</span></span><br><span data-ttu-id="83756-423">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-423">
         - Settings</span></span><br><span data-ttu-id="83756-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-424">
         - TableBindings</span></span><br><span data-ttu-id="83756-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-425">
         - TableCoercion</span></span><br><span data-ttu-id="83756-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-426">
         - TextBindings</span></span><br><span data-ttu-id="83756-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-427">
         - TextCoercion</span></span><br><span data-ttu-id="83756-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-429">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-430">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-430">- TaskPane</span></span><br><span data-ttu-id="83756-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-436">- BindingEvents</span></span><br><span data-ttu-id="83756-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-437">
         - CompressedFile</span></span><br><span data-ttu-id="83756-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-439">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-440">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-440">
         - File</span></span><br><span data-ttu-id="83756-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-442">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-443">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-446">
         - PdfFile</span></span><br><span data-ttu-id="83756-447">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-447">
         - Selection</span></span><br><span data-ttu-id="83756-448">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-448">
         - Settings</span></span><br><span data-ttu-id="83756-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-449">
         - TableBindings</span></span><br><span data-ttu-id="83756-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-450">
         - TableCoercion</span></span><br><span data-ttu-id="83756-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-451">
         - TextBindings</span></span><br><span data-ttu-id="83756-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-452">
         - TextCoercion</span></span><br><span data-ttu-id="83756-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-454">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-455">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-455">- TaskPane</span></span><br><span data-ttu-id="83756-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-461">- BindingEvents</span></span><br><span data-ttu-id="83756-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-462">
         - CompressedFile</span></span><br><span data-ttu-id="83756-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-464">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-465">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-465">
         - File</span></span><br><span data-ttu-id="83756-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-467">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-468">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-471">
         - PdfFile</span></span><br><span data-ttu-id="83756-472">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-472">
         - Selection</span></span><br><span data-ttu-id="83756-473">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-473">
         - Settings</span></span><br><span data-ttu-id="83756-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-474">
         - TableBindings</span></span><br><span data-ttu-id="83756-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-475">
         - TableCoercion</span></span><br><span data-ttu-id="83756-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-476">
         - TextBindings</span></span><br><span data-ttu-id="83756-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-477">
         - TextCoercion</span></span><br><span data-ttu-id="83756-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-479">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="83756-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="83756-480">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="83756-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="83756-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-485">- BindingEvents</span></span><br><span data-ttu-id="83756-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-486">
         - CompressedFile</span></span><br><span data-ttu-id="83756-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-488">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-489">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-489">
         - File</span></span><br><span data-ttu-id="83756-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-491">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-492">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-495">
         - PdfFile</span></span><br><span data-ttu-id="83756-496">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-496">
         - Selection</span></span><br><span data-ttu-id="83756-497">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-497">
         - Settings</span></span><br><span data-ttu-id="83756-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-498">
         - TableBindings</span></span><br><span data-ttu-id="83756-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-499">
         - TableCoercion</span></span><br><span data-ttu-id="83756-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-500">
         - TextBindings</span></span><br><span data-ttu-id="83756-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-501">
         - TextCoercion</span></span><br><span data-ttu-id="83756-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-503">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-504">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-504">- TaskPane</span></span><br><span data-ttu-id="83756-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="83756-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="83756-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-510">- BindingEvents</span></span><br><span data-ttu-id="83756-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-511">
         - CompressedFile</span></span><br><span data-ttu-id="83756-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-513">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-514">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-514">
         - File</span></span><br><span data-ttu-id="83756-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-516">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-517">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-520">
         - PdfFile</span></span><br><span data-ttu-id="83756-521">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-521">
         - Selection</span></span><br><span data-ttu-id="83756-522">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-522">
         - Settings</span></span><br><span data-ttu-id="83756-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-523">
         - TableBindings</span></span><br><span data-ttu-id="83756-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-524">
         - TableCoercion</span></span><br><span data-ttu-id="83756-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-525">
         - TextBindings</span></span><br><span data-ttu-id="83756-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-526">
         - TextCoercion</span></span><br><span data-ttu-id="83756-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-528">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-529">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-529">- TaskPane</span></span><br><span data-ttu-id="83756-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="83756-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="83756-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="83756-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="83756-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="83756-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="83756-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="83756-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="83756-535">- BindingEvents</span></span><br><span data-ttu-id="83756-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-536">
         - CompressedFile</span></span><br><span data-ttu-id="83756-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="83756-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="83756-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-538">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-539">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-539">
         - File</span></span><br><span data-ttu-id="83756-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-541">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="83756-542">
         - MatrixBindings</span></span><br><span data-ttu-id="83756-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="83756-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="83756-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-545">
         - PdfFile</span></span><br><span data-ttu-id="83756-546">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-546">
         - Selection</span></span><br><span data-ttu-id="83756-547">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-547">
         - Settings</span></span><br><span data-ttu-id="83756-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="83756-548">
         - TableBindings</span></span><br><span data-ttu-id="83756-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-549">
         - TableCoercion</span></span><br><span data-ttu-id="83756-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="83756-550">
         - TextBindings</span></span><br><span data-ttu-id="83756-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-551">
         - TextCoercion</span></span><br><span data-ttu-id="83756-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="83756-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="83756-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="83756-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="83756-554">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-554">Platform</span></span></th>
    <th><span data-ttu-id="83756-555">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-555">Extension points</span></span></th>
    <th><span data-ttu-id="83756-556">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="83756-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="83756-558">Office Online</span></span></td>
    <td> <span data-ttu-id="83756-559">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-559">- Content</span></span><br><span data-ttu-id="83756-560">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-560">
         - TaskPane</span></span><br><span data-ttu-id="83756-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-563">- ActiveView</span></span><br><span data-ttu-id="83756-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-564">
         - CompressedFile</span></span><br><span data-ttu-id="83756-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-565">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-566">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-566">
         - File</span></span><br><span data-ttu-id="83756-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-567">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-568">
         - PdfFile</span></span><br><span data-ttu-id="83756-569">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-569">
         - Selection</span></span><br><span data-ttu-id="83756-570">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-570">
         - Settings</span></span><br><span data-ttu-id="83756-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-572">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-573">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-573">- Content</span></span><br><span data-ttu-id="83756-574">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="83756-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="83756-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="83756-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-576">- ActiveView</span></span><br><span data-ttu-id="83756-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-577">
         - CompressedFile</span></span><br><span data-ttu-id="83756-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-578">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-579">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-579">
         - File</span></span><br><span data-ttu-id="83756-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-580">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-581">
         - PdfFile</span></span><br><span data-ttu-id="83756-582">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-582">
         - Selection</span></span><br><span data-ttu-id="83756-583">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-583">
         - Settings</span></span><br><span data-ttu-id="83756-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-585">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-586">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-586">- Content</span></span><br><span data-ttu-id="83756-587">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-587">
         - TaskPane</span></span><br><span data-ttu-id="83756-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-590">- ActiveView</span></span><br><span data-ttu-id="83756-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-591">
         - CompressedFile</span></span><br><span data-ttu-id="83756-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-592">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-593">
         - File</span></span><br><span data-ttu-id="83756-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-594">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-595">
         - PdfFile</span></span><br><span data-ttu-id="83756-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-596">
         - Selection</span></span><br><span data-ttu-id="83756-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-597">
         - Settings</span></span><br><span data-ttu-id="83756-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-599">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-600">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-600">- Content</span></span><br><span data-ttu-id="83756-601">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-601">
         - TaskPane</span></span><br><span data-ttu-id="83756-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-604">- ActiveView</span></span><br><span data-ttu-id="83756-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-605">
         - CompressedFile</span></span><br><span data-ttu-id="83756-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-606">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-607">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-607">
         - File</span></span><br><span data-ttu-id="83756-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-608">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-609">
         - PdfFile</span></span><br><span data-ttu-id="83756-610">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-610">
         - Selection</span></span><br><span data-ttu-id="83756-611">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-611">
         - Settings</span></span><br><span data-ttu-id="83756-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-613">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="83756-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="83756-614">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-614">- Content</span></span><br><span data-ttu-id="83756-615">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="83756-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-617">- ActiveView</span></span><br><span data-ttu-id="83756-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-618">
         - CompressedFile</span></span><br><span data-ttu-id="83756-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-619">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-620">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-620">
         - File</span></span><br><span data-ttu-id="83756-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-621">
         - PdfFile</span></span><br><span data-ttu-id="83756-622">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-622">
         - Selection</span></span><br><span data-ttu-id="83756-623">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-623">
         - Settings</span></span><br><span data-ttu-id="83756-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-624">
         - TextCoercion</span></span><br><span data-ttu-id="83756-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-626">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-627">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-627">- Content</span></span><br><span data-ttu-id="83756-628">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-628">
         - TaskPane</span></span><br><span data-ttu-id="83756-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-631">- ActiveView</span></span><br><span data-ttu-id="83756-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-632">
         - CompressedFile</span></span><br><span data-ttu-id="83756-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-633">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-634">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-634">
         - File</span></span><br><span data-ttu-id="83756-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-635">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-636">
         - PdfFile</span></span><br><span data-ttu-id="83756-637">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-637">
         - Selection</span></span><br><span data-ttu-id="83756-638">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-638">
         - Settings</span></span><br><span data-ttu-id="83756-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-640">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="83756-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="83756-641">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-641">- Content</span></span><br><span data-ttu-id="83756-642">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-642">
         - TaskPane</span></span><br><span data-ttu-id="83756-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="83756-645">- ActiveView</span></span><br><span data-ttu-id="83756-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="83756-646">
         - CompressedFile</span></span><br><span data-ttu-id="83756-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-647">
         - DocumentEvents</span></span><br><span data-ttu-id="83756-648">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="83756-648">
         - File</span></span><br><span data-ttu-id="83756-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-649">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="83756-650">
         - PdfFile</span></span><br><span data-ttu-id="83756-651">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-651">
         - Selection</span></span><br><span data-ttu-id="83756-652">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-652">
         - Settings</span></span><br><span data-ttu-id="83756-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="83756-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="83756-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="83756-655">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-655">Platform</span></span></th>
    <th><span data-ttu-id="83756-656">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-656">Extension points</span></span></th>
    <th><span data-ttu-id="83756-657">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="83756-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="83756-659">Office Online</span></span></td>
    <td> <span data-ttu-id="83756-660">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="83756-660">- Content</span></span><br><span data-ttu-id="83756-661">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-661">
         - TaskPane</span></span><br><span data-ttu-id="83756-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="83756-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="83756-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="83756-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="83756-665">- DocumentEvents</span></span><br><span data-ttu-id="83756-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="83756-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-667">
         - ImageCoercion</span></span><br><span data-ttu-id="83756-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="83756-668">
         - Settings</span></span><br><span data-ttu-id="83756-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="83756-670">Project</span><span class="sxs-lookup"><span data-stu-id="83756-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="83756-671">Plataforma</span><span class="sxs-lookup"><span data-stu-id="83756-671">Platform</span></span></th>
    <th><span data-ttu-id="83756-672">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="83756-672">Extension points</span></span></th>
    <th><span data-ttu-id="83756-673">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="83756-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="83756-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="83756-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-675">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-678">- Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-678">- Selection</span></span><br><span data-ttu-id="83756-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-680">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-681">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-683">- Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-683">- Selection</span></span><br><span data-ttu-id="83756-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="83756-685">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="83756-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="83756-686">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="83756-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="83756-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="83756-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="83756-688">- Seleção</span><span class="sxs-lookup"><span data-stu-id="83756-688">- Selection</span></span><br><span data-ttu-id="83756-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="83756-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="83756-690">Confira também</span><span class="sxs-lookup"><span data-stu-id="83756-690">See also</span></span>

- [<span data-ttu-id="83756-691">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="83756-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="83756-692">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="83756-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="83756-693">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="83756-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="83756-694">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="83756-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
