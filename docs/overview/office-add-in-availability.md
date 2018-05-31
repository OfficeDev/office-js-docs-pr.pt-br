---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438869"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="d6d7d-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d6d7d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="d6d7d-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="d6d7d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="d6d7d-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="d6d7d-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="d6d7d-106">Se uma célula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="d6d7d-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="d6d7d-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="d6d7d-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="d6d7d-p103">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="d6d7d-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="d6d7d-110">Excel</span><span class="sxs-lookup"><span data-stu-id="d6d7d-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="d6d7d-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d6d7d-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="d6d7d-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d6d7d-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="d6d7d-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="d6d7d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="d6d7d-115">Office Online</span></span></td>
    <td> <span data-ttu-id="d6d7d-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-116">- Taskpane</span></span><br><span data-ttu-id="d6d7d-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-117">
        - Content</span></span><br><span data-ttu-id="d6d7d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="d6d7d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="d6d7d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d6d7d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d6d7d-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-124">
        -BindingEvents</span></span><br><span data-ttu-id="d6d7d-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-125">
        -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-126">
        -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-128">
        -TableBindings</span></span><br><span data-ttu-id="d6d7d-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-129">
        -TableCoercion</span></span><br><span data-ttu-id="d6d7d-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-130">
        -TextBindings</span></span><br><span data-ttu-id="d6d7d-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-131">
        -CompressedFile</span></span><br><span data-ttu-id="d6d7d-132">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-132">
        - Settings</span></span><br><span data-ttu-id="d6d7d-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-134">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="d6d7d-135">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-135">
        - Taskpane</span></span><br><span data-ttu-id="d6d7d-136">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="d6d7d-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d6d7d-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-138">
        -BindingEvents</span></span><br><span data-ttu-id="d6d7d-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-139">
        -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-140">
        -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-142">
        -TableBindings</span></span><br><span data-ttu-id="d6d7d-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-143">
        -TableCoercion</span></span><br><span data-ttu-id="d6d7d-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-144">
        -TextBindings</span></span><br><span data-ttu-id="d6d7d-145">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-145">
        - Settings</span></span><br><span data-ttu-id="d6d7d-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-147">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="d6d7d-148">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-148">- Taskpane</span></span><br><span data-ttu-id="d6d7d-149">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-149">
        - Content</span></span><br><span data-ttu-id="d6d7d-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d6d7d-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="d6d7d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d6d7d-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-156">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-157">
        -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-158">
        -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-160">
        -TableBindings</span></span><br><span data-ttu-id="d6d7d-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-161">
        -TableCoercion</span></span><br><span data-ttu-id="d6d7d-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-162">
        -TextBindings</span></span><br><span data-ttu-id="d6d7d-163">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-163">
        - Settings</span></span><br><span data-ttu-id="d6d7d-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-165">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d6d7d-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="d6d7d-166">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-166">- Taskpane</span></span><br><span data-ttu-id="d6d7d-167">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-167">
        - Content</span></span></td>
    <td><span data-ttu-id="d6d7d-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d6d7d-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-172">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-173">
        -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-174">
        -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-176">
        -TableBindings</span></span><br><span data-ttu-id="d6d7d-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-177">
        -TableCoercion</span></span><br><span data-ttu-id="d6d7d-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-178">
        -TextBindings</span></span><br><span data-ttu-id="d6d7d-179">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-179">
        - Settings</span></span><br><span data-ttu-id="d6d7d-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-181">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d6d7d-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="d6d7d-182">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-182">- Taskpane</span></span><br><span data-ttu-id="d6d7d-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-183">
        - Content</span></span><br><span data-ttu-id="d6d7d-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="d6d7d-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="d6d7d-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-189">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-190">
        -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-191">
        -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-193">
        -TableBindings</span></span><br><span data-ttu-id="d6d7d-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-194">
        -TableCoercion</span></span><br><span data-ttu-id="d6d7d-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-195">
        -TextBindings</span></span><br><span data-ttu-id="d6d7d-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="d6d7d-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="d6d7d-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d6d7d-198">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d6d7d-198">Platform</span></span></th>
    <th><span data-ttu-id="d6d7d-199">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d6d7d-199">Extension points</span></span></th> 
    <th><span data-ttu-id="d6d7d-200">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="d6d7d-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="d6d7d-202">Office Online</span></span></td>
    <td> <span data-ttu-id="d6d7d-203">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-203">- Mail Read</span></span><br><span data-ttu-id="d6d7d-204">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-204">
      - Mail Compose</span></span><br><span data-ttu-id="d6d7d-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d6d7d-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d6d7d-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d6d7d-212">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-213">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-214">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-214">- Mail Read</span></span><br><span data-ttu-id="d6d7d-215">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-215">
      - Mail Compose</span></span><br><span data-ttu-id="d6d7d-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="d6d7d-221">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-222">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-223">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-223">- Mail Read</span></span><br><span data-ttu-id="d6d7d-224">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-224">
      - Mail Compose</span></span><br><span data-ttu-id="d6d7d-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="d6d7d-226">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="d6d7d-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="d6d7d-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d6d7d-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d6d7d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d6d7d-233">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-234">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d6d7d-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d6d7d-235">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-235">- Mail Read</span></span><br><span data-ttu-id="d6d7d-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d6d7d-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="d6d7d-242">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-243">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d6d7d-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d6d7d-244">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-244">- Mail Read</span></span><br><span data-ttu-id="d6d7d-245">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-245">
      - Mail Compose</span></span><br><span data-ttu-id="d6d7d-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d6d7d-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="d6d7d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="d6d7d-253">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-254">Office para Android</span><span class="sxs-lookup"><span data-stu-id="d6d7d-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="d6d7d-255">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="d6d7d-255">- Mail Read</span></span><br><span data-ttu-id="d6d7d-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="d6d7d-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="d6d7d-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="d6d7d-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="d6d7d-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="d6d7d-262">não disponível</span><span class="sxs-lookup"><span data-stu-id="d6d7d-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="d6d7d-263">Word</span><span class="sxs-lookup"><span data-stu-id="d6d7d-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d6d7d-264">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d6d7d-264">Platform</span></span></th>
    <th><span data-ttu-id="d6d7d-265">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d6d7d-265">Extension points</span></span></th> 
    <th><span data-ttu-id="d6d7d-266">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="d6d7d-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="d6d7d-268">Office Online</span></span></td>
    <td> <span data-ttu-id="d6d7d-269">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-269">- Taskpane</span></span><br><span data-ttu-id="d6d7d-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-275">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="d6d7d-276">customXmlParts</span></span><br><span data-ttu-id="d6d7d-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-277">
         -MatrixBindings</span></span><br><span data-ttu-id="d6d7d-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-279">
         -TableBindings</span></span><br><span data-ttu-id="d6d7d-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-280">
         -TableCoercion</span></span><br><span data-ttu-id="d6d7d-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-281">
         -TextBindings</span></span><br><span data-ttu-id="d6d7d-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-282">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-283">
         -TextFile</span></span><br><span data-ttu-id="d6d7d-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-284">
         -ImageCoercion</span></span><br><span data-ttu-id="d6d7d-285">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-285">
         - Settings</span></span><br><span data-ttu-id="d6d7d-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-287">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-288">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="d6d7d-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-290">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-291">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="d6d7d-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="d6d7d-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-293">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-294">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-294">
         - File</span></span><br><span data-ttu-id="d6d7d-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="d6d7d-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-296">
         -ImageCoercion</span></span><br><span data-ttu-id="d6d7d-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d6d7d-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-298">
         -TableBindings</span></span><br><span data-ttu-id="d6d7d-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-299">
         -TableCoercion</span></span><br><span data-ttu-id="d6d7d-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-300">
         -TextBindings</span></span><br><span data-ttu-id="d6d7d-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-301">
         -TextFile</span></span><br><span data-ttu-id="d6d7d-302">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-302">
         - Settings</span></span><br><span data-ttu-id="d6d7d-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-303">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-305">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="d6d7d-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-306">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-307">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-307">- Taskpane</span></span><br><span data-ttu-id="d6d7d-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-313">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-314">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="d6d7d-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="d6d7d-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-316">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-317">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-317">
         - File</span></span><br><span data-ttu-id="d6d7d-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="d6d7d-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-319">
         -ImageCoercion</span></span><br><span data-ttu-id="d6d7d-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d6d7d-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-321">
         -TableBindings</span></span><br><span data-ttu-id="d6d7d-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-322">
         -TableCoercion</span></span><br><span data-ttu-id="d6d7d-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-323">
         -TextBindings</span></span><br><span data-ttu-id="d6d7d-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-324">
         -TextFile</span></span><br><span data-ttu-id="d6d7d-325">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-325">
         - Settings</span></span><br><span data-ttu-id="d6d7d-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-326">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-328">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="d6d7d-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-329">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d6d7d-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d6d7d-330">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="d6d7d-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d6d7d-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d6d7d-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-335">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-336">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="d6d7d-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="d6d7d-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-338">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-339">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-339">
         - File</span></span><br><span data-ttu-id="d6d7d-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="d6d7d-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-341">
         -ImageCoercion</span></span><br><span data-ttu-id="d6d7d-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d6d7d-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-343">
         -TableBindings</span></span><br><span data-ttu-id="d6d7d-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-344">
         -TableCoercion</span></span><br><span data-ttu-id="d6d7d-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-345">
         -TextBindings</span></span><br><span data-ttu-id="d6d7d-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-346">
         -TextFile</span></span><br><span data-ttu-id="d6d7d-347">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-347">
         - Settings</span></span><br><span data-ttu-id="d6d7d-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-348">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-350">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="d6d7d-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-351">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d6d7d-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d6d7d-352">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-352">- Taskpane</span></span><br><span data-ttu-id="d6d7d-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="d6d7d-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="d6d7d-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d6d7d-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d6d7d-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-358">-BindingEvents</span></span><br><span data-ttu-id="d6d7d-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-359">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="d6d7d-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="d6d7d-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-361">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-362">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-362">
         - File</span></span><br><span data-ttu-id="d6d7d-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="d6d7d-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-364">
         -ImageCoercion</span></span><br><span data-ttu-id="d6d7d-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="d6d7d-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-366">
         -TableBindings</span></span><br><span data-ttu-id="d6d7d-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-367">
         -TableCoercion</span></span><br><span data-ttu-id="d6d7d-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="d6d7d-368">
         -TextBindings</span></span><br><span data-ttu-id="d6d7d-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-369">
         -TextFile</span></span><br><span data-ttu-id="d6d7d-370">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-370">
         - Settings</span></span><br><span data-ttu-id="d6d7d-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-371">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="d6d7d-373">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="d6d7d-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="d6d7d-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="d6d7d-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d6d7d-375">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d6d7d-375">Platform</span></span></th>
    <th><span data-ttu-id="d6d7d-376">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d6d7d-376">Extension points</span></span></th> 
    <th><span data-ttu-id="d6d7d-377">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="d6d7d-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="d6d7d-379">Office Online</span></span></td>
    <td> <span data-ttu-id="d6d7d-380">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-380">- Content</span></span><br><span data-ttu-id="d6d7d-381">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-381">
         - Taskpane</span></span><br><span data-ttu-id="d6d7d-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d6d7d-384">-ActiveView</span></span><br><span data-ttu-id="d6d7d-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-385">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-386">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-386">
         - File</span></span><br><span data-ttu-id="d6d7d-387">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d6d7d-387">
         - Selection</span></span><br><span data-ttu-id="d6d7d-388">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-388">
         - Settings</span></span><br><span data-ttu-id="d6d7d-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-389">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-391">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-392">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-392">- Content</span></span><br><span data-ttu-id="d6d7d-393">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="d6d7d-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="d6d7d-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="d6d7d-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d6d7d-395">-ActiveView</span></span><br><span data-ttu-id="d6d7d-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-396">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-397">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-398">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-398">
         - File</span></span><br><span data-ttu-id="d6d7d-399">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d6d7d-399">
         - Selection</span></span><br><span data-ttu-id="d6d7d-400">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-400">
         - Settings</span></span><br><span data-ttu-id="d6d7d-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-402">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="d6d7d-403">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-403">- Content</span></span><br><span data-ttu-id="d6d7d-404">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-404">
         - Taskpane</span></span><br><span data-ttu-id="d6d7d-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d6d7d-407">-ActiveView</span></span><br><span data-ttu-id="d6d7d-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-408">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-409">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-410">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-410">
         - File</span></span><br><span data-ttu-id="d6d7d-411">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d6d7d-411">
         - Selection</span></span><br><span data-ttu-id="d6d7d-412">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-412">
         - Settings</span></span><br><span data-ttu-id="d6d7d-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-413">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-415">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d6d7d-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="d6d7d-416">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-416">- Content</span></span><br><span data-ttu-id="d6d7d-417">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="d6d7d-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="d6d7d-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d6d7d-419">-ActiveView</span></span><br><span data-ttu-id="d6d7d-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-420">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-421">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-422">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-422">
         - File</span></span><br><span data-ttu-id="d6d7d-423">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d6d7d-423">
         - Selection</span></span><br><span data-ttu-id="d6d7d-424">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-424">
         - Settings</span></span><br><span data-ttu-id="d6d7d-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-425">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-427">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d6d7d-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="d6d7d-428">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-428">- Content</span></span><br><span data-ttu-id="d6d7d-429">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-429">
         - Taskpane</span></span><br><span data-ttu-id="d6d7d-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="d6d7d-432">-ActiveView</span></span><br><span data-ttu-id="d6d7d-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="d6d7d-433">
         -CompressedFile</span></span><br><span data-ttu-id="d6d7d-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-434">
         -DocumentEvents</span></span><br><span data-ttu-id="d6d7d-435">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-435">
         - File</span></span><br><span data-ttu-id="d6d7d-436">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="d6d7d-436">
         - Selection</span></span><br><span data-ttu-id="d6d7d-437">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-437">
         - Settings</span></span><br><span data-ttu-id="d6d7d-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-438">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="d6d7d-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="d6d7d-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="d6d7d-441">Plataforma</span><span class="sxs-lookup"><span data-stu-id="d6d7d-441">Platform</span></span></th>
    <th><span data-ttu-id="d6d7d-442">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="d6d7d-442">Extension points</span></span></th> 
    <th><span data-ttu-id="d6d7d-443">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="d6d7d-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="d6d7d-445">Office Online</span></span></td>
    <td> <span data-ttu-id="d6d7d-446">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="d6d7d-446">- Content</span></span><br><span data-ttu-id="d6d7d-447">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="d6d7d-447">
         - Taskpane</span></span><br><span data-ttu-id="d6d7d-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="d6d7d-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="d6d7d-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="d6d7d-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="d6d7d-451">-DocumentEvents</span></span><br><span data-ttu-id="d6d7d-452">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="d6d7d-452">
         - Settings</span></span><br><span data-ttu-id="d6d7d-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-453">
         -TextCoercion</span></span><br><span data-ttu-id="d6d7d-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="d6d7d-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="d6d7d-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-456">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="d6d7d-457">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="d6d7d-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-458">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="d6d7d-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="d6d7d-459">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="d6d7d-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="d6d7d-460">\* = Estamos trabalhando nisso.</span><span class="sxs-lookup"><span data-stu-id="d6d7d-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="d6d7d-461">Veja também</span><span class="sxs-lookup"><span data-stu-id="d6d7d-461">See also</span></span>

- [<span data-ttu-id="d6d7d-462">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="d6d7d-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="d6d7d-463">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="d6d7d-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="d6d7d-464">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="d6d7d-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="d6d7d-465">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="d6d7d-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

