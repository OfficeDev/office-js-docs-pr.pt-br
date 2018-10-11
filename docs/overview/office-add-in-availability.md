---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: 6f7b5b565773457e6cd8a9eee69eb304784a29a9
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459312"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d4d93-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d4d93-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d4d93-p101">Para funcionar como esperado, o suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API. As tabelas a seguir contêm a plataforma disponível, os pontos de extensão, os conjuntos de requisitos da API e os conjuntos de requisitos de API comuns que são atualmente suportados para cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="d4d93-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="d4d93-p102">Se uma célula da tabela contiver um asterisco ( \* ), isso significa que estamos trabalhando nela. Para conjuntos de requisitos para Project ou Access, confira [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d4d93-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="d4d93-p103">O número do build para o Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão contém apenas os conjuntos de requisitos do ExcelApi 1.1, WordApi 1.1 e API comum.</span><span class="sxs-lookup"><span data-stu-id="d4d93-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="d4d93-110">Excel</span><span class="sxs-lookup"><span data-stu-id="d4d93-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d4d93-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d4d93-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d4d93-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d4d93-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="d4d93-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="d4d93-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d4d93-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4d93-115">Office Online</span></span></td>
    <td> <span data-ttu-id="d4d93-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-116">- Taskpane</span></span><br><span data-ttu-id="d4d93-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-117">
        - Content</span></span><br><span data-ttu-id="d4d93-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="d4d93-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d4d93-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-127">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-128">
        -BindingEvents</span></span><br><span data-ttu-id="d4d93-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-129">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-130">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-131">
        - File</span></span><br><span data-ttu-id="d4d93-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-132">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-134">
        - Selection</span></span><br><span data-ttu-id="d4d93-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-135">
        - Settings</span></span><br><span data-ttu-id="d4d93-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-136">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-137">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-138">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-140">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="d4d93-141">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-141">
        - Taskpane</span></span><br><span data-ttu-id="d4d93-142">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d4d93-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-143">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-144">
        -BindingEvents</span></span><br><span data-ttu-id="d4d93-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-145">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-146">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-147">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-147">
        - File</span></span><br><span data-ttu-id="d4d93-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-148">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-149">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-151">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-151">
        - Selection</span></span><br><span data-ttu-id="d4d93-152">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-152">
        - Settings</span></span><br><span data-ttu-id="d4d93-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-153">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-154">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-155">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-157">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="d4d93-158">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-158">- Taskpane</span></span><br><span data-ttu-id="d4d93-159">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-159">
        - Content</span></span><br><span data-ttu-id="d4d93-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-160">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4d93-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-161">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-168">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-169">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-170">-BindingEvents</span></span><br><span data-ttu-id="d4d93-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-171">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-172">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-173">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-173">
        - File</span></span><br><span data-ttu-id="d4d93-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-174">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-175">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-177">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-177">
        - Selection</span></span><br><span data-ttu-id="d4d93-178">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-178">
        - Settings</span></span><br><span data-ttu-id="d4d93-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-179">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-180">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-181">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-183">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="d4d93-184">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-184">- Taskpane</span></span><br><span data-ttu-id="d4d93-185">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-185">
        - Content</span></span><br><span data-ttu-id="d4d93-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4d93-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-187">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-193">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-194">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-195">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-196">-BindingEvents</span></span><br><span data-ttu-id="d4d93-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-197">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-198">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-199">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-199">
        - File</span></span><br><span data-ttu-id="d4d93-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-200">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-201">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-203">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-203">
        - Selection</span></span><br><span data-ttu-id="d4d93-204">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-204">
        - Settings</span></span><br><span data-ttu-id="d4d93-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-205">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-206">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-207">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-209">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d4d93-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="d4d93-210">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-210">- Taskpane</span></span><br><span data-ttu-id="d4d93-211">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-211">
        - Content</span></span></td>
    <td><span data-ttu-id="d4d93-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-212">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-217">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-218">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-219">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-220">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-221">-BindingEvents</span></span><br><span data-ttu-id="d4d93-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-222">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-223">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-224">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-224">
        - File</span></span><br><span data-ttu-id="d4d93-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-225">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-226">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-228">
        - Selection</span></span><br><span data-ttu-id="d4d93-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-229">
        - Settings</span></span><br><span data-ttu-id="d4d93-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-230">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-231">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-232">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-234">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="d4d93-235">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-235">- Taskpane</span></span><br><span data-ttu-id="d4d93-236">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-236">
        - Content</span></span><br><span data-ttu-id="d4d93-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4d93-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-238">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-242">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-243">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-244">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-245">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-246">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-247">-BindingEvents</span></span><br><span data-ttu-id="d4d93-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-248">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-249">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-250">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-250">
        - File</span></span><br><span data-ttu-id="d4d93-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-251">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-252">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-254">
        -PdfFile</span></span><br><span data-ttu-id="d4d93-255">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-255">
        - Selection</span></span><br><span data-ttu-id="d4d93-256">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-256">
        - Settings</span></span><br><span data-ttu-id="d4d93-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-257">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-258">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-259">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-261">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="d4d93-262">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-262">- Taskpane</span></span><br><span data-ttu-id="d4d93-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-263">
        - Content</span></span><br><span data-ttu-id="d4d93-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d4d93-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-265">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d4d93-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d4d93-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d4d93-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-268">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d4d93-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-269">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="d4d93-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-270">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="d4d93-271">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="d4d93-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-272">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="d4d93-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-273">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d4d93-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-274">-BindingEvents</span></span><br><span data-ttu-id="d4d93-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-275">
        -CompressedFile</span></span><br><span data-ttu-id="d4d93-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-276">
        -DocumentEvents</span></span><br><span data-ttu-id="d4d93-277">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-277">
        - File</span></span><br><span data-ttu-id="d4d93-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-278">
        -ImageCoercion</span></span><br><span data-ttu-id="d4d93-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-279">
        -MatrixBindings</span></span><br><span data-ttu-id="d4d93-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-281">
        -PdfFile</span></span><br><span data-ttu-id="d4d93-282">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-282">
        - Selection</span></span><br><span data-ttu-id="d4d93-283">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-283">
        - Settings</span></span><br><span data-ttu-id="d4d93-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-284">
        -TableBindings</span></span><br><span data-ttu-id="d4d93-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-285">
        -TableCoercion</span></span><br><span data-ttu-id="d4d93-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-286">
        -TextBindings</span></span><br><span data-ttu-id="d4d93-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="d4d93-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="d4d93-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4d93-289">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d4d93-289">Platform</span></span></th>
    <th><span data-ttu-id="d4d93-290">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d4d93-290">Extension points</span></span></th>
    <th><span data-ttu-id="d4d93-291">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4d93-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d4d93-292"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4d93-293">Office Online</span></span></td>
    <td> <span data-ttu-id="d4d93-294">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-294">- Mail Read</span></span><br><span data-ttu-id="d4d93-295">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-295">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-296">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-297">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-298">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-299">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-300">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-301">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4d93-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-302">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4d93-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4d93-304">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-305">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-306">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-306">- Mail Read</span></span><br><span data-ttu-id="d4d93-307">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-307">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-308">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-309">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-310">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-311">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-312">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="d4d93-313">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-314">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-315">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-315">- Mail Read</span></span><br><span data-ttu-id="d4d93-316">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-316">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-317">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d4d93-318">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="d4d93-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d4d93-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-319">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-320">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-321">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-322">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-323">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4d93-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-324">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4d93-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-325">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4d93-326">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-327">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="d4d93-328">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-328">- Mail Read</span></span><br><span data-ttu-id="d4d93-329">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-329">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-330">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d4d93-331">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="d4d93-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d4d93-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-332">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-333">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-334">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-335">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4d93-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="d4d93-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="d4d93-339">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-340">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d4d93-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d4d93-341">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-341">- Mail Read</span></span><br><span data-ttu-id="d4d93-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-342">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-343">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-344">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-345">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d4d93-348">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-349">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-350">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-350">- Mail Read</span></span><br><span data-ttu-id="d4d93-351">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-351">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-352">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-353">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-354">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-355">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-356">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4d93-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d4d93-359">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-360">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-361">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-361">- Mail Read</span></span><br><span data-ttu-id="d4d93-362">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-362">
      - Mail Compose</span></span><br><span data-ttu-id="d4d93-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-363">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-364">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-365">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-366">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d4d93-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d4d93-370">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-371">Office para Android</span><span class="sxs-lookup"><span data-stu-id="d4d93-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="d4d93-372">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d4d93-372">- Mail Read</span></span><br><span data-ttu-id="d4d93-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-373">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-374">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d4d93-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-375">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d4d93-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-376">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d4d93-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-377">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d4d93-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-378">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d4d93-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="d4d93-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="d4d93-380">Word</span><span class="sxs-lookup"><span data-stu-id="d4d93-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4d93-381">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d4d93-381">Platform</span></span></th>
    <th><span data-ttu-id="d4d93-382">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d4d93-382">Extension points</span></span></th>
    <th><span data-ttu-id="d4d93-383">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4d93-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d4d93-384"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4d93-385">Office Online</span></span></td>
    <td> <span data-ttu-id="d4d93-386">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-386">- Taskpane</span></span><br><span data-ttu-id="d4d93-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-387">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-388">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-389">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-390">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-391">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-392">-BindingEvents</span></span><br><span data-ttu-id="d4d93-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-393">customXmlParts</span></span><br><span data-ttu-id="d4d93-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-394">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-395">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-395">
         - File</span></span><br><span data-ttu-id="d4d93-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-397">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-398">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-401">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-402">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-402">
         - Selection</span></span><br><span data-ttu-id="d4d93-403">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-403">
         - Settings</span></span><br><span data-ttu-id="d4d93-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-404">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-405">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-406">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-407">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-409">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-410">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="d4d93-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-411">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-412">-BindingEvents</span></span><br><span data-ttu-id="d4d93-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-413">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-414">customXmlParts</span></span><br><span data-ttu-id="d4d93-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-415">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-416">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-416">
         - File</span></span><br><span data-ttu-id="d4d93-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-418">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-419">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-422">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-423">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-423">
         - Selection</span></span><br><span data-ttu-id="d4d93-424">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-424">
         - Settings</span></span><br><span data-ttu-id="d4d93-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-425">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-426">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-427">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-428">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-430">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-431">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-431">- Taskpane</span></span><br><span data-ttu-id="d4d93-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-432">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-433">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-434">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-435">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-436">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-437">-BindingEvents</span></span><br><span data-ttu-id="d4d93-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-438">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-439">customXmlParts</span></span><br><span data-ttu-id="d4d93-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-440">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-441">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-441">
         - File</span></span><br><span data-ttu-id="d4d93-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-443">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-444">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-447">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-448">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-448">
         - Selection</span></span><br><span data-ttu-id="d4d93-449">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-449">
         - Settings</span></span><br><span data-ttu-id="d4d93-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-450">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-451">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-452">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-453">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-455">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="d4d93-456">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-456">- Taskpane</span></span><br><span data-ttu-id="d4d93-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-459">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-460">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-461">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-462">-BindingEvents</span></span><br><span data-ttu-id="d4d93-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-463">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-464">customXmlParts</span></span><br><span data-ttu-id="d4d93-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-465">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-466">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-466">
         - File</span></span><br><span data-ttu-id="d4d93-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-468">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-469">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-472">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-473">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-473">
         - Selection</span></span><br><span data-ttu-id="d4d93-474">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-474">
         - Settings</span></span><br><span data-ttu-id="d4d93-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-475">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-476">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-477">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-478">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-480">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d4d93-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d4d93-481">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="d4d93-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-482">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4d93-485">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4d93-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-486">-BindingEvents</span></span><br><span data-ttu-id="d4d93-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-487">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-488">customXmlParts</span></span><br><span data-ttu-id="d4d93-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-489">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-490">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-490">
         - File</span></span><br><span data-ttu-id="d4d93-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-492">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-493">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-496">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-497">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-497">
         - Selection</span></span><br><span data-ttu-id="d4d93-498">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-498">
         - Settings</span></span><br><span data-ttu-id="d4d93-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-499">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-500">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-501">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-502">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-504">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-505">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-505">- Taskpane</span></span><br><span data-ttu-id="d4d93-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-506">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-507">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-508">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-509">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4d93-510">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4d93-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-511">-BindingEvents</span></span><br><span data-ttu-id="d4d93-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-512">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-513">customXmlParts</span></span><br><span data-ttu-id="d4d93-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-514">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-515">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-515">
         - File</span></span><br><span data-ttu-id="d4d93-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-517">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-518">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-521">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-522">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-522">
         - Selection</span></span><br><span data-ttu-id="d4d93-523">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-523">
         - Settings</span></span><br><span data-ttu-id="d4d93-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-524">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-525">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-526">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-527">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-529">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-530">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-530">- Taskpane</span></span><br><span data-ttu-id="d4d93-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-531">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-532">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d4d93-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-533">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d4d93-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-534">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d4d93-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4d93-535">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4d93-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-536">-BindingEvents</span></span><br><span data-ttu-id="d4d93-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-537">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d4d93-538">customXmlParts</span></span><br><span data-ttu-id="d4d93-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-539">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-540">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-540">
         - File</span></span><br><span data-ttu-id="d4d93-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-542">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-543">
         -MatrixBindings</span></span><br><span data-ttu-id="d4d93-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="d4d93-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d4d93-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-546">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-547">
         - Selection</span></span><br><span data-ttu-id="d4d93-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-548">
         - Settings</span></span><br><span data-ttu-id="d4d93-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-549">
         -TableBindings</span></span><br><span data-ttu-id="d4d93-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-550">
         -TableCoercion</span></span><br><span data-ttu-id="d4d93-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d4d93-551">
         -TextBindings</span></span><br><span data-ttu-id="d4d93-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-552">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d4d93-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d4d93-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4d93-555">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d4d93-555">Platform</span></span></th>
    <th><span data-ttu-id="d4d93-556">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d4d93-556">Extension points</span></span></th>
    <th><span data-ttu-id="d4d93-557">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4d93-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d4d93-558"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4d93-559">Office Online</span></span></td>
    <td> <span data-ttu-id="d4d93-560">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-560">- Content</span></span><br><span data-ttu-id="d4d93-561">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-561">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-562">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-563">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-564">-ActiveView</span></span><br><span data-ttu-id="d4d93-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-565">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-566">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-567">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-567">
         - File</span></span><br><span data-ttu-id="d4d93-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-568">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-569">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-570">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-570">
         - Selection</span></span><br><span data-ttu-id="d4d93-571">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-571">
         - Settings</span></span><br><span data-ttu-id="d4d93-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-573">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-574">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-574">- Content</span></span><br><span data-ttu-id="d4d93-575">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="d4d93-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d4d93-576">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d4d93-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-577">-ActiveView</span></span><br><span data-ttu-id="d4d93-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-578">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-579">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-580">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-580">
         - File</span></span><br><span data-ttu-id="d4d93-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-581">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-582">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-583">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-583">
         - Selection</span></span><br><span data-ttu-id="d4d93-584">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-584">
         - Settings</span></span><br><span data-ttu-id="d4d93-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-586">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d4d93-587">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-587">- Content</span></span><br><span data-ttu-id="d4d93-588">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-588">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-589">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-590">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-591">-ActiveView</span></span><br><span data-ttu-id="d4d93-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-592">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-593">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-594">
         - File</span></span><br><span data-ttu-id="d4d93-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-595">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-596">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-597">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-597">
         - Selection</span></span><br><span data-ttu-id="d4d93-598">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-598">
         - Settings</span></span><br><span data-ttu-id="d4d93-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-600">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="d4d93-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="d4d93-601">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-601">- Content</span></span><br><span data-ttu-id="d4d93-602">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-602">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-603">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-604">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-605">-ActiveView</span></span><br><span data-ttu-id="d4d93-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-606">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-607">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-608">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-608">
         - File</span></span><br><span data-ttu-id="d4d93-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-609">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-610">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-611">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-611">
         - Selection</span></span><br><span data-ttu-id="d4d93-612">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-612">
         - Settings</span></span><br><span data-ttu-id="d4d93-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-614">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d4d93-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d4d93-615">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-615">- Content</span></span><br><span data-ttu-id="d4d93-616">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="d4d93-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-617">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="d4d93-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-618">-ActiveView</span></span><br><span data-ttu-id="d4d93-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-619">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-620">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-621">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-621">
         - File</span></span><br><span data-ttu-id="d4d93-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-622">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-623">
         - Selection</span></span><br><span data-ttu-id="d4d93-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-624">
         - Settings</span></span><br><span data-ttu-id="d4d93-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-625">
         -TextCoercion</span></span><br><span data-ttu-id="d4d93-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-627">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-628">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-628">- Content</span></span><br><span data-ttu-id="d4d93-629">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-629">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-630">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-631">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-632">-ActiveView</span></span><br><span data-ttu-id="d4d93-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-633">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-634">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-635">
         - File</span></span><br><span data-ttu-id="d4d93-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-636">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-637">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-638">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-638">
         - Selection</span></span><br><span data-ttu-id="d4d93-639">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-639">
         - Settings</span></span><br><span data-ttu-id="d4d93-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-641">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="d4d93-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="d4d93-642">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-642">- Content</span></span><br><span data-ttu-id="d4d93-643">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-643">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-644">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-645">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d4d93-646">-ActiveView</span></span><br><span data-ttu-id="d4d93-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-647">
         -CompressedFile</span></span><br><span data-ttu-id="d4d93-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-648">
         -DocumentEvents</span></span><br><span data-ttu-id="d4d93-649">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d4d93-649">
         - File</span></span><br><span data-ttu-id="d4d93-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-650">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="d4d93-651">
         -PdfFile</span></span><br><span data-ttu-id="d4d93-652">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d4d93-652">
         - Selection</span></span><br><span data-ttu-id="d4d93-653">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-653">
         - Settings</span></span><br><span data-ttu-id="d4d93-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="d4d93-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="d4d93-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d4d93-656">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d4d93-656">Platform</span></span></th>
    <th><span data-ttu-id="d4d93-657">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d4d93-657">Extension points</span></span></th>
    <th><span data-ttu-id="d4d93-658">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="d4d93-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d4d93-659"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d4d93-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="d4d93-660">Office Online</span></span></td>
    <td> <span data-ttu-id="d4d93-661">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d4d93-661">- Content</span></span><br><span data-ttu-id="d4d93-662">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d4d93-662">
         - Taskpane</span></span><br><span data-ttu-id="d4d93-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-663">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d4d93-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-664">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d4d93-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d4d93-665">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d4d93-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d4d93-666">-DocumentEvents</span></span><br><span data-ttu-id="d4d93-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="d4d93-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-668">
         -ImageCoercion</span></span><br><span data-ttu-id="d4d93-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d4d93-669">
         - Settings</span></span><br><span data-ttu-id="d4d93-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d4d93-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="d4d93-671">Confira também</span><span class="sxs-lookup"><span data-stu-id="d4d93-671">See also</span></span>

- [<span data-ttu-id="d4d93-672">Visão geral da plataforma de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d4d93-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d4d93-673">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="d4d93-673">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d4d93-674">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="d4d93-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="d4d93-675">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="d4d93-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
