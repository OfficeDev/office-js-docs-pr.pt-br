---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 39a80f322c282e29e6e8c4363f0c82522b33b75d
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579923"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="82927-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="82927-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="82927-p101">Para funcionar como esperado, o suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API. As tabelas a seguir contêm a plataforma disponível, os pontos de extensão, os conjuntos de requisitos da API e os conjuntos de requisitos de API comuns que são atualmente suportados para cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="82927-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="82927-p102">Se uma célula da tabela contiver um asterisco ( \* ), isso significa que estamos trabalhando nela. Para conjuntos de requisitos para Project ou Access, confira [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="82927-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="82927-p103">O número do build para o Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão contém apenas os conjuntos de requisitos do ExcelApi 1.1, WordApi 1.1 e API comum.</span><span class="sxs-lookup"><span data-stu-id="82927-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="82927-110">Excel</span><span class="sxs-lookup"><span data-stu-id="82927-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="82927-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="82927-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="82927-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="82927-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="82927-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="82927-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="82927-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="82927-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="82927-115">Office Online</span></span></td>
    <td> <span data-ttu-id="82927-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-116">- Taskpane</span></span><br><span data-ttu-id="82927-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-117">
        - Content</span></span><br><span data-ttu-id="82927-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="82927-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="82927-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-128">
        -BindingEvents</span></span><br><span data-ttu-id="82927-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-129">
        -CompressedFile</span></span><br><span data-ttu-id="82927-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-130">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-131">
        - File</span></span><br><span data-ttu-id="82927-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-132">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-134">
        - Selection</span></span><br><span data-ttu-id="82927-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-135">
        - Settings</span></span><br><span data-ttu-id="82927-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-136">
        -TableBindings</span></span><br><span data-ttu-id="82927-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-137">
        -TableCoercion</span></span><br><span data-ttu-id="82927-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-138">
        -TextBindings</span></span><br><span data-ttu-id="82927-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-140">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="82927-141">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-141">
        - Taskpane</span></span><br><span data-ttu-id="82927-142">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="82927-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-144">
        -BindingEvents</span></span><br><span data-ttu-id="82927-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-145">
        -CompressedFile</span></span><br><span data-ttu-id="82927-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-146">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-147">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-147">
        - File</span></span><br><span data-ttu-id="82927-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-148">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-149">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-151">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-151">
        - Selection</span></span><br><span data-ttu-id="82927-152">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-152">
        - Settings</span></span><br><span data-ttu-id="82927-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-153">
        -TableBindings</span></span><br><span data-ttu-id="82927-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-154">
        -TableCoercion</span></span><br><span data-ttu-id="82927-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-155">
        -TextBindings</span></span><br><span data-ttu-id="82927-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-157">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="82927-158">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-158">- Taskpane</span></span><br><span data-ttu-id="82927-159">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-159">
        - Content</span></span><br><span data-ttu-id="82927-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="82927-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-170">-BindingEvents</span></span><br><span data-ttu-id="82927-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-171">
        -CompressedFile</span></span><br><span data-ttu-id="82927-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-172">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-173">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-173">
        - File</span></span><br><span data-ttu-id="82927-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-174">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-175">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-177">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-177">
        - Selection</span></span><br><span data-ttu-id="82927-178">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-178">
        - Settings</span></span><br><span data-ttu-id="82927-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-179">
        -TableBindings</span></span><br><span data-ttu-id="82927-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-180">
        -TableCoercion</span></span><br><span data-ttu-id="82927-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-181">
        -TextBindings</span></span><br><span data-ttu-id="82927-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-183">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="82927-184">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-184">- Taskpane</span></span><br><span data-ttu-id="82927-185">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-185">
        - Content</span></span><br><span data-ttu-id="82927-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="82927-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-196">-BindingEvents</span></span><br><span data-ttu-id="82927-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-197">
        -CompressedFile</span></span><br><span data-ttu-id="82927-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-198">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-199">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-199">
        - File</span></span><br><span data-ttu-id="82927-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-200">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-201">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-203">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-203">
        - Selection</span></span><br><span data-ttu-id="82927-204">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-204">
        - Settings</span></span><br><span data-ttu-id="82927-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-205">
        -TableBindings</span></span><br><span data-ttu-id="82927-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-206">
        -TableCoercion</span></span><br><span data-ttu-id="82927-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-207">
        -TextBindings</span></span><br><span data-ttu-id="82927-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-209">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="82927-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="82927-210">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-210">- Taskpane</span></span><br><span data-ttu-id="82927-211">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-211">
        - Content</span></span></td>
    <td><span data-ttu-id="82927-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-221">-BindingEvents</span></span><br><span data-ttu-id="82927-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-222">
        -CompressedFile</span></span><br><span data-ttu-id="82927-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-223">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-224">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-224">
        - File</span></span><br><span data-ttu-id="82927-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-225">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-226">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-228">
        - Selection</span></span><br><span data-ttu-id="82927-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-229">
        - Settings</span></span><br><span data-ttu-id="82927-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-230">
        -TableBindings</span></span><br><span data-ttu-id="82927-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-231">
        -TableCoercion</span></span><br><span data-ttu-id="82927-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-232">
        -TextBindings</span></span><br><span data-ttu-id="82927-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-234">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="82927-235">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-235">- Taskpane</span></span><br><span data-ttu-id="82927-236">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-236">
        - Content</span></span><br><span data-ttu-id="82927-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="82927-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-247">-BindingEvents</span></span><br><span data-ttu-id="82927-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-248">
        -CompressedFile</span></span><br><span data-ttu-id="82927-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-249">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-250">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-250">
        - File</span></span><br><span data-ttu-id="82927-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-251">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-252">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-254">
        -PdfFile</span></span><br><span data-ttu-id="82927-255">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-255">
        - Selection</span></span><br><span data-ttu-id="82927-256">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-256">
        - Settings</span></span><br><span data-ttu-id="82927-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-257">
        -TableBindings</span></span><br><span data-ttu-id="82927-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-258">
        -TableCoercion</span></span><br><span data-ttu-id="82927-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-259">
        -TextBindings</span></span><br><span data-ttu-id="82927-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-261">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="82927-262">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-262">- Taskpane</span></span><br><span data-ttu-id="82927-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-263">
        - Content</span></span><br><span data-ttu-id="82927-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="82927-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="82927-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="82927-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="82927-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="82927-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="82927-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="82927-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="82927-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="82927-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="82927-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="82927-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-274">-BindingEvents</span></span><br><span data-ttu-id="82927-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-275">
        -CompressedFile</span></span><br><span data-ttu-id="82927-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-276">
        -DocumentEvents</span></span><br><span data-ttu-id="82927-277">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-277">
        - File</span></span><br><span data-ttu-id="82927-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-278">
        -ImageCoercion</span></span><br><span data-ttu-id="82927-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-279">
        -MatrixBindings</span></span><br><span data-ttu-id="82927-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="82927-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-281">
        -PdfFile</span></span><br><span data-ttu-id="82927-282">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-282">
        - Selection</span></span><br><span data-ttu-id="82927-283">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-283">
        - Settings</span></span><br><span data-ttu-id="82927-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-284">
        -TableBindings</span></span><br><span data-ttu-id="82927-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-285">
        -TableCoercion</span></span><br><span data-ttu-id="82927-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-286">
        -TextBindings</span></span><br><span data-ttu-id="82927-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="82927-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="82927-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="82927-289">Plataforma</span><span class="sxs-lookup"><span data-stu-id="82927-289">Platform</span></span></th>
    <th><span data-ttu-id="82927-290">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="82927-290">Extension points</span></span></th>
    <th><span data-ttu-id="82927-291">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="82927-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="82927-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="82927-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="82927-293">Office Online</span></span></td>
    <td> <span data-ttu-id="82927-294">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-294">- Mail Read</span></span><br><span data-ttu-id="82927-295">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-295">
      - Mail Compose</span></span><br><span data-ttu-id="82927-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="82927-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="82927-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="82927-304">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-305">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-306">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-306">- Mail Read</span></span><br><span data-ttu-id="82927-307">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-307">
      - Mail Compose</span></span><br><span data-ttu-id="82927-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="82927-313">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-314">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-315">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-315">- Mail Read</span></span><br><span data-ttu-id="82927-316">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-316">
      - Mail Compose</span></span><br><span data-ttu-id="82927-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="82927-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="82927-318">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="82927-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="82927-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="82927-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="82927-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="82927-326">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-327">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="82927-328">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-328">- Mail Read</span></span><br><span data-ttu-id="82927-329">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-329">
      - Mail Compose</span></span><br><span data-ttu-id="82927-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="82927-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="82927-331">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="82927-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="82927-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="82927-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="82927-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="82927-339">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-340">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="82927-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="82927-341">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-341">- Mail Read</span></span><br><span data-ttu-id="82927-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="82927-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="82927-348">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-349">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="82927-350">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-350">- Mail Read</span></span><br><span data-ttu-id="82927-351">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-351">
      - Mail Compose</span></span><br><span data-ttu-id="82927-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="82927-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="82927-359">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-360">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="82927-361">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-361">- Mail Read</span></span><br><span data-ttu-id="82927-362">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="82927-362">
      - Mail Compose</span></span><br><span data-ttu-id="82927-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="82927-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="82927-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="82927-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="82927-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="82927-371">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-372">Office para Android</span><span class="sxs-lookup"><span data-stu-id="82927-372">Office for Android</span></span></td>
    <td> <span data-ttu-id="82927-373">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="82927-373">- Mail Read</span></span><br><span data-ttu-id="82927-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="82927-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-375">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="82927-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="82927-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="82927-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="82927-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="82927-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="82927-379">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="82927-380">Não disponível</span><span class="sxs-lookup"><span data-stu-id="82927-380">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="82927-381">Word</span><span class="sxs-lookup"><span data-stu-id="82927-381">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="82927-382">Plataforma</span><span class="sxs-lookup"><span data-stu-id="82927-382">Platform</span></span></th>
    <th><span data-ttu-id="82927-383">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="82927-383">Extension points</span></span></th>
    <th><span data-ttu-id="82927-384">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="82927-384">API requirement sets</span></span></th>
    <th><span data-ttu-id="82927-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="82927-385"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="82927-386">Office Online</span><span class="sxs-lookup"><span data-stu-id="82927-386">Office Online</span></span></td>
    <td> <span data-ttu-id="82927-387">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-387">- Taskpane</span></span><br><span data-ttu-id="82927-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-389">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-392">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-393">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-393">-BindingEvents</span></span><br><span data-ttu-id="82927-394">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-394">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-395">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-395">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-396">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-396">
         - File</span></span><br><span data-ttu-id="82927-397">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-397">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-398">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-398">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-399">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-399">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-400">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-400">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-401">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-401">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-402">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-402">
         -PdfFile</span></span><br><span data-ttu-id="82927-403">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-403">
         - Selection</span></span><br><span data-ttu-id="82927-404">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-404">
         - Settings</span></span><br><span data-ttu-id="82927-405">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-405">
         -TableBindings</span></span><br><span data-ttu-id="82927-406">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-406">
         -TableCoercion</span></span><br><span data-ttu-id="82927-407">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-407">
         -TextBindings</span></span><br><span data-ttu-id="82927-408">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-408">
         -TextCoercion</span></span><br><span data-ttu-id="82927-409">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-409">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-410">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-410">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-411">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-411">- Taskpane</span></span></td>
    <td> <span data-ttu-id="82927-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-412">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-413">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-413">-BindingEvents</span></span><br><span data-ttu-id="82927-414">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-414">
         -CompressedFile</span></span><br><span data-ttu-id="82927-415">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-415">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-416">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-416">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-417">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-417">
         - File</span></span><br><span data-ttu-id="82927-418">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-418">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-419">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-419">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-420">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-420">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-421">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-421">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-422">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-422">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-423">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-423">
         -PdfFile</span></span><br><span data-ttu-id="82927-424">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-424">
         - Selection</span></span><br><span data-ttu-id="82927-425">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-425">
         - Settings</span></span><br><span data-ttu-id="82927-426">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-426">
         -TableBindings</span></span><br><span data-ttu-id="82927-427">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-427">
         -TableCoercion</span></span><br><span data-ttu-id="82927-428">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-428">
         -TextBindings</span></span><br><span data-ttu-id="82927-429">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-429">
         -TextCoercion</span></span><br><span data-ttu-id="82927-430">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-430">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-431">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-431">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-432">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-432">- Taskpane</span></span><br><span data-ttu-id="82927-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-433">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-434">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-437">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-438">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-438">-BindingEvents</span></span><br><span data-ttu-id="82927-439">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-439">
         -CompressedFile</span></span><br><span data-ttu-id="82927-440">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-440">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-441">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-441">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-442">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-442">
         - File</span></span><br><span data-ttu-id="82927-443">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-443">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-444">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-444">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-445">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-445">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-446">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-446">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-447">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-447">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-448">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-448">
         -PdfFile</span></span><br><span data-ttu-id="82927-449">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-449">
         - Selection</span></span><br><span data-ttu-id="82927-450">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-450">
         - Settings</span></span><br><span data-ttu-id="82927-451">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-451">
         -TableBindings</span></span><br><span data-ttu-id="82927-452">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-452">
         -TableCoercion</span></span><br><span data-ttu-id="82927-453">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-453">
         -TextBindings</span></span><br><span data-ttu-id="82927-454">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-454">
         -TextCoercion</span></span><br><span data-ttu-id="82927-455">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-455">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-456">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-456">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="82927-457">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-457">- Taskpane</span></span><br><span data-ttu-id="82927-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-458">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-459">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-462">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-463">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-463">-BindingEvents</span></span><br><span data-ttu-id="82927-464">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-464">
         -CompressedFile</span></span><br><span data-ttu-id="82927-465">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-465">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-466">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-466">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-467">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-467">
         - File</span></span><br><span data-ttu-id="82927-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-469">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-470">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-470">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-471">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-471">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-472">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-472">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-473">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-473">
         -PdfFile</span></span><br><span data-ttu-id="82927-474">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-474">
         - Selection</span></span><br><span data-ttu-id="82927-475">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-475">
         - Settings</span></span><br><span data-ttu-id="82927-476">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-476">
         -TableBindings</span></span><br><span data-ttu-id="82927-477">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-477">
         -TableCoercion</span></span><br><span data-ttu-id="82927-478">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-478">
         -TextBindings</span></span><br><span data-ttu-id="82927-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-479">
         -TextCoercion</span></span><br><span data-ttu-id="82927-480">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-480">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-481">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="82927-481">Office for iOS</span></span></td>
    <td> <span data-ttu-id="82927-482">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-482">- Taskpane</span></span></td>
    <td> <span data-ttu-id="82927-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-483">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="82927-486">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="82927-487">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-487">-BindingEvents</span></span><br><span data-ttu-id="82927-488">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-488">
         -CompressedFile</span></span><br><span data-ttu-id="82927-489">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-489">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-490">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-490">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-491">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-491">
         - File</span></span><br><span data-ttu-id="82927-492">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-492">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-493">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-493">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-494">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-494">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-495">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-495">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-496">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-496">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-497">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-497">
         -PdfFile</span></span><br><span data-ttu-id="82927-498">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-498">
         - Selection</span></span><br><span data-ttu-id="82927-499">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-499">
         - Settings</span></span><br><span data-ttu-id="82927-500">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-500">
         -TableBindings</span></span><br><span data-ttu-id="82927-501">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-501">
         -TableCoercion</span></span><br><span data-ttu-id="82927-502">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-502">
         -TextBindings</span></span><br><span data-ttu-id="82927-503">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-503">
         -TextCoercion</span></span><br><span data-ttu-id="82927-504">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-504">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-505">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-505">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="82927-506">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-506">- Taskpane</span></span><br><span data-ttu-id="82927-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-507">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-508">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="82927-511">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="82927-512">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-512">-BindingEvents</span></span><br><span data-ttu-id="82927-513">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-513">
         -CompressedFile</span></span><br><span data-ttu-id="82927-514">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-514">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-515">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-515">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-516">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-516">
         - File</span></span><br><span data-ttu-id="82927-517">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-517">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-518">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-518">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-519">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-519">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-520">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-520">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-521">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-521">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-522">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-522">
         -PdfFile</span></span><br><span data-ttu-id="82927-523">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-523">
         - Selection</span></span><br><span data-ttu-id="82927-524">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-524">
         - Settings</span></span><br><span data-ttu-id="82927-525">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-525">
         -TableBindings</span></span><br><span data-ttu-id="82927-526">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-526">
         -TableCoercion</span></span><br><span data-ttu-id="82927-527">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-527">
         -TextBindings</span></span><br><span data-ttu-id="82927-528">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-528">
         -TextCoercion</span></span><br><span data-ttu-id="82927-529">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-529">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-530">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-530">Office for Mac</span></span></td>
    <td> <span data-ttu-id="82927-531">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-531">- Taskpane</span></span><br><span data-ttu-id="82927-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-532">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-533">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="82927-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="82927-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="82927-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="82927-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="82927-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="82927-536">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="82927-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="82927-537">-BindingEvents</span></span><br><span data-ttu-id="82927-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-538">
         -CompressedFile</span></span><br><span data-ttu-id="82927-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="82927-539">
         -CustomXmlParts</span></span><br><span data-ttu-id="82927-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-540">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-541">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-541">
         - File</span></span><br><span data-ttu-id="82927-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-542">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-543">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-543">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="82927-544">
         -MatrixBindings</span></span><br><span data-ttu-id="82927-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-545">
         -MatrixCoercion</span></span><br><span data-ttu-id="82927-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-546">
         -OoxmlCoercion</span></span><br><span data-ttu-id="82927-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-547">
         -PdfFile</span></span><br><span data-ttu-id="82927-548">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-548">
         - Selection</span></span><br><span data-ttu-id="82927-549">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-549">
         - Settings</span></span><br><span data-ttu-id="82927-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="82927-550">
         -TableBindings</span></span><br><span data-ttu-id="82927-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-551">
         -TableCoercion</span></span><br><span data-ttu-id="82927-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="82927-552">
         -TextBindings</span></span><br><span data-ttu-id="82927-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-553">
         -TextCoercion</span></span><br><span data-ttu-id="82927-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="82927-554">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="82927-555">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="82927-555">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="82927-556">Plataforma</span><span class="sxs-lookup"><span data-stu-id="82927-556">Platform</span></span></th>
    <th><span data-ttu-id="82927-557">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="82927-557">Extension points</span></span></th>
    <th><span data-ttu-id="82927-558">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="82927-558">API requirement sets</span></span></th>
    <th><span data-ttu-id="82927-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="82927-559"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="82927-560">Office Online</span><span class="sxs-lookup"><span data-stu-id="82927-560">Office Online</span></span></td>
    <td> <span data-ttu-id="82927-561">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-561">- Content</span></span><br><span data-ttu-id="82927-562">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-562">
         - Taskpane</span></span><br><span data-ttu-id="82927-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-564">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-565">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-565">-ActiveView</span></span><br><span data-ttu-id="82927-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-566">
         -CompressedFile</span></span><br><span data-ttu-id="82927-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-567">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-568">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-568">
         - File</span></span><br><span data-ttu-id="82927-569">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-569">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-570">
         -PdfFile</span></span><br><span data-ttu-id="82927-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-571">
         - Selection</span></span><br><span data-ttu-id="82927-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-572">
         - Settings</span></span><br><span data-ttu-id="82927-573">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-573">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-574">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-574">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-575">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-575">- Content</span></span><br><span data-ttu-id="82927-576">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-576">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="82927-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="82927-577">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="82927-578">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-578">-ActiveView</span></span><br><span data-ttu-id="82927-579">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-579">
         -CompressedFile</span></span><br><span data-ttu-id="82927-580">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-580">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-581">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-581">
         - File</span></span><br><span data-ttu-id="82927-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-582">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-583">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-583">
         -PdfFile</span></span><br><span data-ttu-id="82927-584">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-584">
         - Selection</span></span><br><span data-ttu-id="82927-585">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-585">
         - Settings</span></span><br><span data-ttu-id="82927-586">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-586">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-587">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-587">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="82927-588">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-588">- Content</span></span><br><span data-ttu-id="82927-589">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-589">
         - Taskpane</span></span><br><span data-ttu-id="82927-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-590">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-591">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-592">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-592">-ActiveView</span></span><br><span data-ttu-id="82927-593">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-593">
         -CompressedFile</span></span><br><span data-ttu-id="82927-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-594">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-595">
         - File</span></span><br><span data-ttu-id="82927-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-596">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-597">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-597">
         -PdfFile</span></span><br><span data-ttu-id="82927-598">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-598">
         - Selection</span></span><br><span data-ttu-id="82927-599">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-599">
         - Settings</span></span><br><span data-ttu-id="82927-600">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-600">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-601">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="82927-601">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="82927-602">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-602">- Content</span></span><br><span data-ttu-id="82927-603">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-603">
         - Taskpane</span></span><br><span data-ttu-id="82927-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-604">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-605">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-606">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-606">-ActiveView</span></span><br><span data-ttu-id="82927-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-607">
         -CompressedFile</span></span><br><span data-ttu-id="82927-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-608">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-609">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-609">
         - File</span></span><br><span data-ttu-id="82927-610">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-610">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-611">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-611">
         -PdfFile</span></span><br><span data-ttu-id="82927-612">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-612">
         - Selection</span></span><br><span data-ttu-id="82927-613">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-613">
         - Settings</span></span><br><span data-ttu-id="82927-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-614">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-615">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="82927-615">Office for iOS</span></span></td>
    <td> <span data-ttu-id="82927-616">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-616">- Content</span></span><br><span data-ttu-id="82927-617">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-617">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="82927-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-618">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="82927-619">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-619">-ActiveView</span></span><br><span data-ttu-id="82927-620">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-620">
         -CompressedFile</span></span><br><span data-ttu-id="82927-621">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-621">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-622">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-622">
         - File</span></span><br><span data-ttu-id="82927-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-623">
         -PdfFile</span></span><br><span data-ttu-id="82927-624">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-624">
         - Selection</span></span><br><span data-ttu-id="82927-625">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-625">
         - Settings</span></span><br><span data-ttu-id="82927-626">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-626">
         -TextCoercion</span></span><br><span data-ttu-id="82927-627">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-627">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-628">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-628">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="82927-629">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-629">- Content</span></span><br><span data-ttu-id="82927-630">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-630">
         - Taskpane</span></span><br><span data-ttu-id="82927-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-631">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-632">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-633">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-633">-ActiveView</span></span><br><span data-ttu-id="82927-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-634">
         -CompressedFile</span></span><br><span data-ttu-id="82927-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-635">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-636">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-636">
         - File</span></span><br><span data-ttu-id="82927-637">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-637">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-638">
         -PdfFile</span></span><br><span data-ttu-id="82927-639">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-639">
         - Selection</span></span><br><span data-ttu-id="82927-640">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-640">
         - Settings</span></span><br><span data-ttu-id="82927-641">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-641">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="82927-642">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="82927-642">Office for Mac</span></span></td>
    <td> <span data-ttu-id="82927-643">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-643">- Content</span></span><br><span data-ttu-id="82927-644">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-644">
         - Taskpane</span></span><br><span data-ttu-id="82927-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-645">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-646">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-647">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="82927-647">-ActiveView</span></span><br><span data-ttu-id="82927-648">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="82927-648">
         -CompressedFile</span></span><br><span data-ttu-id="82927-649">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-649">
         -DocumentEvents</span></span><br><span data-ttu-id="82927-650">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="82927-650">
         - File</span></span><br><span data-ttu-id="82927-651">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-651">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-652">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="82927-652">
         -PdfFile</span></span><br><span data-ttu-id="82927-653">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="82927-653">
         - Selection</span></span><br><span data-ttu-id="82927-654">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-654">
         - Settings</span></span><br><span data-ttu-id="82927-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-655">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="82927-656">OneNote</span><span class="sxs-lookup"><span data-stu-id="82927-656">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="82927-657">Plataforma</span><span class="sxs-lookup"><span data-stu-id="82927-657">Platform</span></span></th>
    <th><span data-ttu-id="82927-658">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="82927-658">Extension points</span></span></th>
    <th><span data-ttu-id="82927-659">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="82927-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="82927-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="82927-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="82927-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="82927-661">Office Online</span></span></td>
    <td> <span data-ttu-id="82927-662">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="82927-662">- Content</span></span><br><span data-ttu-id="82927-663">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="82927-663">
         - Taskpane</span></span><br><span data-ttu-id="82927-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="82927-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="82927-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="82927-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="82927-666">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="82927-667">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="82927-667">-DocumentEvents</span></span><br><span data-ttu-id="82927-668">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-668">
         -HtmlCoercion</span></span><br><span data-ttu-id="82927-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-669">
         -ImageCoercion</span></span><br><span data-ttu-id="82927-670">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="82927-670">
         - Settings</span></span><br><span data-ttu-id="82927-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="82927-671">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="82927-672">Confira também</span><span class="sxs-lookup"><span data-stu-id="82927-672">See also</span></span>

- [<span data-ttu-id="82927-673">Visão geral da plataforma de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="82927-673">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="82927-674">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="82927-674">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="82927-675">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="82927-675">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="82927-676">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="82927-676">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
