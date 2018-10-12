---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 10/03/2018
ms.openlocfilehash: bc7ac5c97c041a546c160c05cffc2c80db1ff1b1
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506347"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e9059-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e9059-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e9059-p101">Para funcionar como esperado, o suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API. As tabelas a seguir contêm a plataforma disponível, os pontos de extensão, os conjuntos de requisitos da API e os conjuntos de requisitos de API comuns que são atualmente suportados para cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e9059-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="e9059-p102">Se uma célula da tabela contiver um asterisco ( \* ), isso significa que estamos trabalhando nela. Para conjuntos de requisitos para Project ou Access, confira [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e9059-p102">If a table cell contains an asterisk ( \* ), that means we’re working on it. For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="e9059-p103">O número do build para o Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão contém apenas os conjuntos de requisitos do ExcelApi 1.1, WordApi 1.1 e API comum.</span><span class="sxs-lookup"><span data-stu-id="e9059-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e9059-110">Excel</span><span class="sxs-lookup"><span data-stu-id="e9059-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e9059-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e9059-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e9059-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e9059-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e9059-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e9059-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e9059-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e9059-114"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9059-115">Office Online</span></span></td>
    <td> <span data-ttu-id="e9059-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-116">- Taskpane</span></span><br><span data-ttu-id="e9059-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-117">
        - Content</span></span><br><span data-ttu-id="e9059-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a>
    </span><span class="sxs-lookup"><span data-stu-id="e9059-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e9059-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-127">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-128">
        -BindingEvents</span></span><br><span data-ttu-id="e9059-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-129">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-130">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-131">
        - File</span></span><br><span data-ttu-id="e9059-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-132">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-133">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-134">
        - Selection</span></span><br><span data-ttu-id="e9059-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-135">
        - Settings</span></span><br><span data-ttu-id="e9059-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-136">
        -TableBindings</span></span><br><span data-ttu-id="e9059-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-137">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-138">
        -TextBindings</span></span><br><span data-ttu-id="e9059-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-139">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-140">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-140">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e9059-141">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-141">
        - Taskpane</span></span><br><span data-ttu-id="e9059-142">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-142">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e9059-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-143">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-144">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-144">
        -BindingEvents</span></span><br><span data-ttu-id="e9059-145">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-145">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-146">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-146">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-147">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-147">
        - File</span></span><br><span data-ttu-id="e9059-148">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-148">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-149">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-149">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-150">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-150">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-151">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-151">
        - Selection</span></span><br><span data-ttu-id="e9059-152">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-152">
        - Settings</span></span><br><span data-ttu-id="e9059-153">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-153">
        -TableBindings</span></span><br><span data-ttu-id="e9059-154">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-154">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-155">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-155">
        -TextBindings</span></span><br><span data-ttu-id="e9059-156">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-156">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-157">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-157">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e9059-158">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-158">- Taskpane</span></span><br><span data-ttu-id="e9059-159">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-159">
        - Content</span></span><br><span data-ttu-id="e9059-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9059-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-161">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-167">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-168">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-170">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-170">-BindingEvents</span></span><br><span data-ttu-id="e9059-171">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-171">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-172">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-172">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-173">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-173">
        - File</span></span><br><span data-ttu-id="e9059-174">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-174">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-175">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-175">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-176">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-176">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-177">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-177">
        - Selection</span></span><br><span data-ttu-id="e9059-178">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-178">
        - Settings</span></span><br><span data-ttu-id="e9059-179">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-179">
        -TableBindings</span></span><br><span data-ttu-id="e9059-180">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-180">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-181">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-181">
        -TextBindings</span></span><br><span data-ttu-id="e9059-182">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-182">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-183">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-183">Office for Windows Desktop.</span></span></td>
    <td><span data-ttu-id="e9059-184">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-184">- Taskpane</span></span><br><span data-ttu-id="e9059-185">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-185">
        - Content</span></span><br><span data-ttu-id="e9059-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9059-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-187">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-193">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-195">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-196">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-196">-BindingEvents</span></span><br><span data-ttu-id="e9059-197">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-197">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-198">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-198">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-199">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-199">
        - File</span></span><br><span data-ttu-id="e9059-200">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-200">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-201">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-201">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-202">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-202">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-203">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-203">
        - Selection</span></span><br><span data-ttu-id="e9059-204">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-204">
        - Settings</span></span><br><span data-ttu-id="e9059-205">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-205">
        -TableBindings</span></span><br><span data-ttu-id="e9059-206">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-206">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-207">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-207">
        -TextBindings</span></span><br><span data-ttu-id="e9059-208">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-208">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-209">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e9059-209">Office for iOS</span></span></td>
    <td><span data-ttu-id="e9059-210">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-210">- Taskpane</span></span><br><span data-ttu-id="e9059-211">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-211">
        - Content</span></span></td>
    <td><span data-ttu-id="e9059-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-212">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-217">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-218">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-219">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-220">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-221">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-221">-BindingEvents</span></span><br><span data-ttu-id="e9059-222">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-222">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-223">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-223">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-224">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-224">
        - File</span></span><br><span data-ttu-id="e9059-225">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-225">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-226">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-227">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-228">
        - Selection</span></span><br><span data-ttu-id="e9059-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-229">
        - Settings</span></span><br><span data-ttu-id="e9059-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-230">
        -TableBindings</span></span><br><span data-ttu-id="e9059-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-231">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-232">
        -TextBindings</span></span><br><span data-ttu-id="e9059-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-233">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-234">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-234">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e9059-235">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-235">- Taskpane</span></span><br><span data-ttu-id="e9059-236">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-236">
        - Content</span></span><br><span data-ttu-id="e9059-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9059-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-238">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-244">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-245">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-246">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-247">-BindingEvents</span></span><br><span data-ttu-id="e9059-248">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-248">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-249">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-249">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-250">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-250">
        - File</span></span><br><span data-ttu-id="e9059-251">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-251">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-252">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-253">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-254">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-254">
        -PdfFile</span></span><br><span data-ttu-id="e9059-255">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-255">
        - Selection</span></span><br><span data-ttu-id="e9059-256">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-256">
        - Settings</span></span><br><span data-ttu-id="e9059-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-257">
        -TableBindings</span></span><br><span data-ttu-id="e9059-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-258">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-259">
        -TextBindings</span></span><br><span data-ttu-id="e9059-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-260">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-261">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-261">Office for Mac</span></span></td>
    <td><span data-ttu-id="e9059-262">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-262">- Taskpane</span></span><br><span data-ttu-id="e9059-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-263">
        - Content</span></span><br><span data-ttu-id="e9059-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e9059-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-265">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e9059-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e9059-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e9059-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e9059-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e9059-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e9059-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-271">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e9059-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e9059-272">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e9059-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-273">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e9059-274">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-274">-BindingEvents</span></span><br><span data-ttu-id="e9059-275">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-275">
        -CompressedFile</span></span><br><span data-ttu-id="e9059-276">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-276">
        -DocumentEvents</span></span><br><span data-ttu-id="e9059-277">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-277">
        - File</span></span><br><span data-ttu-id="e9059-278">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-278">
        -ImageCoercion</span></span><br><span data-ttu-id="e9059-279">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-279">
        -MatrixBindings</span></span><br><span data-ttu-id="e9059-280">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-280">
        -MatrixCoercion</span></span><br><span data-ttu-id="e9059-281">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-281">
        -PdfFile</span></span><br><span data-ttu-id="e9059-282">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-282">
        - Selection</span></span><br><span data-ttu-id="e9059-283">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-283">
        - Settings</span></span><br><span data-ttu-id="e9059-284">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-284">
        -TableBindings</span></span><br><span data-ttu-id="e9059-285">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-285">
        -TableCoercion</span></span><br><span data-ttu-id="e9059-286">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-286">
        -TextBindings</span></span><br><span data-ttu-id="e9059-287">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-287">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e9059-288">Outlook</span><span class="sxs-lookup"><span data-stu-id="e9059-288">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9059-289">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e9059-289">Platform</span></span></th>
    <th><span data-ttu-id="e9059-290">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e9059-290">Extension points</span></span></th>
    <th><span data-ttu-id="e9059-291">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e9059-291">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9059-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e9059-292"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-293">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9059-293">Office Online</span></span></td>
    <td> <span data-ttu-id="e9059-294">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-294">- Mail Read</span></span><br><span data-ttu-id="e9059-295">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-295">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-297">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9059-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-302">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9059-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-303">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9059-304">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-304">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-305">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-305">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-306">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-306">- Mail Read</span></span><br><span data-ttu-id="e9059-307">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-307">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-309">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-311">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-312">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e9059-313">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-313">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-314">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-314">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-315">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-315">- Mail Read</span></span><br><span data-ttu-id="e9059-316">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-316">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-317">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e9059-318">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="e9059-318">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e9059-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-319">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9059-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-324">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9059-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-325">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9059-326">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-326">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-327">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-327">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e9059-328">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-328">- Mail Read</span></span><br><span data-ttu-id="e9059-329">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-329">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-330">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e9059-331">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="e9059-331">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e9059-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-332">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9059-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e9059-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e9059-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e9059-339">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-340">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e9059-340">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e9059-341">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-341">- Mail Read</span></span><br><span data-ttu-id="e9059-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-343">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-346">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e9059-348">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-348">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-349">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-349">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-350">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-350">- Mail Read</span></span><br><span data-ttu-id="e9059-351">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-351">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-353">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9059-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-358">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e9059-359">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-359">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-360">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-360">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-361">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-361">- Mail Read</span></span><br><span data-ttu-id="e9059-362">
      - Redação de email</span><span class="sxs-lookup"><span data-stu-id="e9059-362">
      - Mail Compose</span></span><br><span data-ttu-id="e9059-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-364">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e9059-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e9059-369">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e9059-370">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-370">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-371">Office para Android</span><span class="sxs-lookup"><span data-stu-id="e9059-371">Office for Android</span></span></td>
    <td> <span data-ttu-id="e9059-372">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e9059-372">- Mail Read</span></span><br><span data-ttu-id="e9059-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-374">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e9059-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e9059-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e9059-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e9059-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e9059-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e9059-378">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e9059-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e9059-379">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e9059-380">Word</span><span class="sxs-lookup"><span data-stu-id="e9059-380">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9059-381">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e9059-381">Platform</span></span></th>
    <th><span data-ttu-id="e9059-382">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e9059-382">Extension points</span></span></th>
    <th><span data-ttu-id="e9059-383">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e9059-383">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9059-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e9059-384"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-385">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9059-385">Office Online</span></span></td>
    <td> <span data-ttu-id="e9059-386">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-386">- Taskpane</span></span><br><span data-ttu-id="e9059-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-387">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-391">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-392">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-392">-BindingEvents</span></span><br><span data-ttu-id="e9059-393">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-393">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-394">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-394">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-395">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-395">
         - File</span></span><br><span data-ttu-id="e9059-396">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-396">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-397">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-397">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-398">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-398">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-399">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-399">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-400">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-400">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-401">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-401">
         -PdfFile</span></span><br><span data-ttu-id="e9059-402">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-402">
         - Selection</span></span><br><span data-ttu-id="e9059-403">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-403">
         - Settings</span></span><br><span data-ttu-id="e9059-404">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-404">
         -TableBindings</span></span><br><span data-ttu-id="e9059-405">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-405">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-406">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-406">
         -TextBindings</span></span><br><span data-ttu-id="e9059-407">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-407">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-408">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-408">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-409">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-409">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-410">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-410">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e9059-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-411">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-412">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-412">-BindingEvents</span></span><br><span data-ttu-id="e9059-413">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-413">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-414">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-414">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-415">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-415">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-416">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-416">
         - File</span></span><br><span data-ttu-id="e9059-417">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-417">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-418">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-418">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-419">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-419">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-420">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-420">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-421">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-421">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-422">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-422">
         -PdfFile</span></span><br><span data-ttu-id="e9059-423">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-423">
         - Selection</span></span><br><span data-ttu-id="e9059-424">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-424">
         - Settings</span></span><br><span data-ttu-id="e9059-425">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-425">
         -TableBindings</span></span><br><span data-ttu-id="e9059-426">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-426">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-427">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-427">
         -TextBindings</span></span><br><span data-ttu-id="e9059-428">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-428">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-429">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-429">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-430">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-430">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-431">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-431">- Taskpane</span></span><br><span data-ttu-id="e9059-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-432">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-433">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-436">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-437">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-437">-BindingEvents</span></span><br><span data-ttu-id="e9059-438">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-438">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-439">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-439">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-440">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-440">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-441">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-441">
         - File</span></span><br><span data-ttu-id="e9059-442">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-442">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-443">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-443">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-444">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-444">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-445">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-445">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-446">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-446">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-447">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-447">
         -PdfFile</span></span><br><span data-ttu-id="e9059-448">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-448">
         - Selection</span></span><br><span data-ttu-id="e9059-449">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-449">
         - Settings</span></span><br><span data-ttu-id="e9059-450">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-450">
         -TableBindings</span></span><br><span data-ttu-id="e9059-451">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-451">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-452">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-452">
         -TextBindings</span></span><br><span data-ttu-id="e9059-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-453">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-454">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-454">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-455">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-455">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e9059-456">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-456">- Taskpane</span></span><br><span data-ttu-id="e9059-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-457">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-458">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-461">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-462">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-462">-BindingEvents</span></span><br><span data-ttu-id="e9059-463">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-463">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-464">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-464">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-465">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-465">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-466">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-466">
         - File</span></span><br><span data-ttu-id="e9059-467">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-467">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-468">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-468">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-469">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-469">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-470">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-470">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-471">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-471">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-472">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-472">
         -PdfFile</span></span><br><span data-ttu-id="e9059-473">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-473">
         - Selection</span></span><br><span data-ttu-id="e9059-474">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-474">
         - Settings</span></span><br><span data-ttu-id="e9059-475">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-475">
         -TableBindings</span></span><br><span data-ttu-id="e9059-476">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-476">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-477">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-477">
         -TextBindings</span></span><br><span data-ttu-id="e9059-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-478">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-479">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-479">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-480">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e9059-480">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e9059-481">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-481">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e9059-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-482">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9059-485">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9059-486">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-486">-BindingEvents</span></span><br><span data-ttu-id="e9059-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-487">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-488">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-488">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-489">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-489">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-490">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-490">
         - File</span></span><br><span data-ttu-id="e9059-491">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-491">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-492">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-492">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-493">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-493">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-494">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-494">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-495">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-495">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-496">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-496">
         -PdfFile</span></span><br><span data-ttu-id="e9059-497">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-497">
         - Selection</span></span><br><span data-ttu-id="e9059-498">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-498">
         - Settings</span></span><br><span data-ttu-id="e9059-499">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-499">
         -TableBindings</span></span><br><span data-ttu-id="e9059-500">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-500">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-501">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-501">
         -TextBindings</span></span><br><span data-ttu-id="e9059-502">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-502">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-503">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-503">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-504">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-504">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-505">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-505">- Taskpane</span></span><br><span data-ttu-id="e9059-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-506">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-507">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9059-510">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9059-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-511">-BindingEvents</span></span><br><span data-ttu-id="e9059-512">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-512">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-513">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-513">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-514">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-514">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-515">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-515">
         - File</span></span><br><span data-ttu-id="e9059-516">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-516">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-517">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-517">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-518">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-518">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-519">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-519">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-520">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-520">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-521">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-521">
         -PdfFile</span></span><br><span data-ttu-id="e9059-522">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-522">
         - Selection</span></span><br><span data-ttu-id="e9059-523">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-523">
         - Settings</span></span><br><span data-ttu-id="e9059-524">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-524">
         -TableBindings</span></span><br><span data-ttu-id="e9059-525">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-525">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-526">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-526">
         -TextBindings</span></span><br><span data-ttu-id="e9059-527">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-527">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-528">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-528">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-529">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-529">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-530">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-530">- Taskpane</span></span><br><span data-ttu-id="e9059-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-531">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-532">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e9059-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e9059-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e9059-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e9059-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e9059-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9059-535">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9059-536">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-536">-BindingEvents</span></span><br><span data-ttu-id="e9059-537">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-537">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-538">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e9059-538">
         -CustomXmlParts</span></span><br><span data-ttu-id="e9059-539">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-539">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-540">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-540">
         - File</span></span><br><span data-ttu-id="e9059-541">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-541">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-542">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-542">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-543">
         -MatrixBindings</span></span><br><span data-ttu-id="e9059-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-544">
         -MatrixCoercion</span></span><br><span data-ttu-id="e9059-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-545">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e9059-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-546">
         -PdfFile</span></span><br><span data-ttu-id="e9059-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-547">
         - Selection</span></span><br><span data-ttu-id="e9059-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-548">
         - Settings</span></span><br><span data-ttu-id="e9059-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-549">
         -TableBindings</span></span><br><span data-ttu-id="e9059-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-550">
         -TableCoercion</span></span><br><span data-ttu-id="e9059-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e9059-551">
         -TextBindings</span></span><br><span data-ttu-id="e9059-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-552">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e9059-553">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e9059-554">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e9059-554">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9059-555">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e9059-555">Platform</span></span></th>
    <th><span data-ttu-id="e9059-556">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e9059-556">Extension points</span></span></th>
    <th><span data-ttu-id="e9059-557">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e9059-557">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9059-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e9059-558"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-559">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9059-559">Office Online</span></span></td>
    <td> <span data-ttu-id="e9059-560">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-560">- Content</span></span><br><span data-ttu-id="e9059-561">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-561">
         - Taskpane</span></span><br><span data-ttu-id="e9059-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-563">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-564">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-564">-ActiveView</span></span><br><span data-ttu-id="e9059-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-565">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-566">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-567">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-567">
         - File</span></span><br><span data-ttu-id="e9059-568">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-568">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-569">
         -PdfFile</span></span><br><span data-ttu-id="e9059-570">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-570">
         - Selection</span></span><br><span data-ttu-id="e9059-571">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-571">
         - Settings</span></span><br><span data-ttu-id="e9059-572">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-572">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-573">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-573">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-574">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-574">- Content</span></span><br><span data-ttu-id="e9059-575">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-575">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="e9059-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e9059-576">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e9059-577">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-577">-ActiveView</span></span><br><span data-ttu-id="e9059-578">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-578">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-579">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-580">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-580">
         - File</span></span><br><span data-ttu-id="e9059-581">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-581">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-582">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-582">
         -PdfFile</span></span><br><span data-ttu-id="e9059-583">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-583">
         - Selection</span></span><br><span data-ttu-id="e9059-584">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-584">
         - Settings</span></span><br><span data-ttu-id="e9059-585">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-585">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-586">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-586">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e9059-587">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-587">- Content</span></span><br><span data-ttu-id="e9059-588">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-588">
         - Taskpane</span></span><br><span data-ttu-id="e9059-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-589">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-590">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-591">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-591">-ActiveView</span></span><br><span data-ttu-id="e9059-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-592">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-593">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-594">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-594">
         - File</span></span><br><span data-ttu-id="e9059-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-595">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-596">
         -PdfFile</span></span><br><span data-ttu-id="e9059-597">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-597">
         - Selection</span></span><br><span data-ttu-id="e9059-598">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-598">
         - Settings</span></span><br><span data-ttu-id="e9059-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-599">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-600">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="e9059-600">Office for Windows Desktop.</span></span></td>
    <td> <span data-ttu-id="e9059-601">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-601">- Content</span></span><br><span data-ttu-id="e9059-602">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-602">
         - Taskpane</span></span><br><span data-ttu-id="e9059-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-603">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-604">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-605">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-605">-ActiveView</span></span><br><span data-ttu-id="e9059-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-606">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-607">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-608">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-608">
         - File</span></span><br><span data-ttu-id="e9059-609">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-609">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-610">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-610">
         -PdfFile</span></span><br><span data-ttu-id="e9059-611">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-611">
         - Selection</span></span><br><span data-ttu-id="e9059-612">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-612">
         - Settings</span></span><br><span data-ttu-id="e9059-613">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-613">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-614">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e9059-614">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e9059-615">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-615">- Content</span></span><br><span data-ttu-id="e9059-616">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-616">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="e9059-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-617">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e9059-618">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-618">-ActiveView</span></span><br><span data-ttu-id="e9059-619">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-619">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-620">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-620">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-621">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-621">
         - File</span></span><br><span data-ttu-id="e9059-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-622">
         -PdfFile</span></span><br><span data-ttu-id="e9059-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-623">
         - Selection</span></span><br><span data-ttu-id="e9059-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-624">
         - Settings</span></span><br><span data-ttu-id="e9059-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-625">
         -TextCoercion</span></span><br><span data-ttu-id="e9059-626">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-626">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-627">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-627">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-628">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-628">- Content</span></span><br><span data-ttu-id="e9059-629">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-629">
         - Taskpane</span></span><br><span data-ttu-id="e9059-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-630">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-631">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-632">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-632">-ActiveView</span></span><br><span data-ttu-id="e9059-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-633">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-634">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-635">
         - File</span></span><br><span data-ttu-id="e9059-636">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-636">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-637">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-637">
         -PdfFile</span></span><br><span data-ttu-id="e9059-638">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-638">
         - Selection</span></span><br><span data-ttu-id="e9059-639">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-639">
         - Settings</span></span><br><span data-ttu-id="e9059-640">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-640">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-641">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="e9059-641">Office for Mac</span></span></td>
    <td> <span data-ttu-id="e9059-642">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-642">- Content</span></span><br><span data-ttu-id="e9059-643">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-643">
         - Taskpane</span></span><br><span data-ttu-id="e9059-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-644">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-645">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-646">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e9059-646">-ActiveView</span></span><br><span data-ttu-id="e9059-647">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e9059-647">
         -CompressedFile</span></span><br><span data-ttu-id="e9059-648">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-648">
         -DocumentEvents</span></span><br><span data-ttu-id="e9059-649">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e9059-649">
         - File</span></span><br><span data-ttu-id="e9059-650">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-650">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-651">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e9059-651">
         -PdfFile</span></span><br><span data-ttu-id="e9059-652">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e9059-652">
         - Selection</span></span><br><span data-ttu-id="e9059-653">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-653">
         - Settings</span></span><br><span data-ttu-id="e9059-654">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-654">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e9059-655">OneNote</span><span class="sxs-lookup"><span data-stu-id="e9059-655">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e9059-656">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e9059-656">Platform</span></span></th>
    <th><span data-ttu-id="e9059-657">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e9059-657">Extension points</span></span></th>
    <th><span data-ttu-id="e9059-658">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e9059-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="e9059-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e9059-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e9059-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="e9059-660">Office Online</span></span></td>
    <td> <span data-ttu-id="e9059-661">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e9059-661">- Content</span></span><br><span data-ttu-id="e9059-662">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e9059-662">
         - Taskpane</span></span><br><span data-ttu-id="e9059-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Suplementos de comandos</a></span><span class="sxs-lookup"><span data-stu-id="e9059-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e9059-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e9059-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e9059-665">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e9059-666">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e9059-666">-DocumentEvents</span></span><br><span data-ttu-id="e9059-667">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-667">
         -HtmlCoercion</span></span><br><span data-ttu-id="e9059-668">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-668">
         -ImageCoercion</span></span><br><span data-ttu-id="e9059-669">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e9059-669">
         - Settings</span></span><br><span data-ttu-id="e9059-670">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e9059-670">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e9059-671">Confira também</span><span class="sxs-lookup"><span data-stu-id="e9059-671">See also</span></span>

- [<span data-ttu-id="e9059-672">Visão geral da plataforma de suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e9059-672">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e9059-673">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="e9059-673">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e9059-674">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="e9059-674">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e9059-675">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="e9059-675">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
