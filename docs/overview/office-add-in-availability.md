---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compat?veis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 03/23/2018
ms.openlocfilehash: f50ab7e5312702eb25fbb2c8a25291c5ff5027a7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7e98d-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7e98d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7e98d-104">Seu suplemento do Office pode depender de um host espec?fico do Office, um conjunto de requisitos, um membro de API ou uma vers?o da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="7e98d-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="7e98d-105">As tabelas a seguir cont?m as plataformas dispon?veis, os pontos de extens?o, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente s?o compat?veis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="7e98d-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="7e98d-106">Se uma c?lula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="7e98d-106">If a table cell contains an asterisk ( \* ), that means we?re working on it.</span></span> <span data-ttu-id="7e98d-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="7e98d-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="7e98d-p103">O n?mero do build do Office 2016 instalado via MSI ? 16.0.4266.1001. Esta vers?o s? cont?m os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="7e98d-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="7e98d-110">Excel</span><span class="sxs-lookup"><span data-stu-id="7e98d-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7e98d-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="7e98d-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7e98d-112">Pontos de extens?o</span><span class="sxs-lookup"><span data-stu-id="7e98d-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="7e98d-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="7e98d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="7e98d-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e98d-115">Office Online</span></span></td>
    <td> <span data-ttu-id="7e98d-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-116">- Taskpane</span></span><br><span data-ttu-id="7e98d-117">
        - Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-117">
        - Content</span></span><br><span data-ttu-id="7e98d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="7e98d-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7e98d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e98d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e98d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e98d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e98d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e98d-124">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-124">
        -BindingEvents</span></span><br><span data-ttu-id="7e98d-125">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-125">
        -DocumentEvents</span></span><br><span data-ttu-id="7e98d-126">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-126">
        -MatrixBindings</span></span><br><span data-ttu-id="7e98d-127">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-127">
        -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-128">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-128">
        -TableBindings</span></span><br><span data-ttu-id="7e98d-129">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-129">
        -TableCoercion</span></span><br><span data-ttu-id="7e98d-130">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-130">
        -TextBindings</span></span><br><span data-ttu-id="7e98d-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-131">
        -CompressedFile</span></span><br><span data-ttu-id="7e98d-132">
        - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-132">
        - Settings</span></span><br><span data-ttu-id="7e98d-133">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-133">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-134">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-134">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="7e98d-135">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-135">
        - Taskpane</span></span><br><span data-ttu-id="7e98d-136">
        - Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-136">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7e98d-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-137">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e98d-138">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-138">
        -BindingEvents</span></span><br><span data-ttu-id="7e98d-139">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-139">
        -DocumentEvents</span></span><br><span data-ttu-id="7e98d-140">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-140">
        -MatrixBindings</span></span><br><span data-ttu-id="7e98d-141">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-141">
        -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-142">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-142">
        -TableBindings</span></span><br><span data-ttu-id="7e98d-143">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-143">
        -TableCoercion</span></span><br><span data-ttu-id="7e98d-144">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-144">
        -TextBindings</span></span><br><span data-ttu-id="7e98d-145">
        - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-145">
        - Settings</span></span><br><span data-ttu-id="7e98d-146">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-146">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-147">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-147">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="7e98d-148">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-148">- Taskpane</span></span><br><span data-ttu-id="7e98d-149">
        - Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-149">
        - Content</span></span><br><span data-ttu-id="7e98d-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-150">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e98d-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-151">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e98d-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-152">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e98d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e98d-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-154">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7e98d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e98d-156">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-156">-BindingEvents</span></span><br><span data-ttu-id="7e98d-157">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-157">
        -DocumentEvents</span></span><br><span data-ttu-id="7e98d-158">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-158">
        -MatrixBindings</span></span><br><span data-ttu-id="7e98d-159">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-159">
        -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-160">
        -TableBindings</span></span><br><span data-ttu-id="7e98d-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-161">
        -TableCoercion</span></span><br><span data-ttu-id="7e98d-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-162">
        -TextBindings</span></span><br><span data-ttu-id="7e98d-163">
        - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-163">
        - Settings</span></span><br><span data-ttu-id="7e98d-164">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-164">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-165">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="7e98d-165">Office for iOS</span></span></td>
    <td><span data-ttu-id="7e98d-166">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-166">- Taskpane</span></span><br><span data-ttu-id="7e98d-167">
        - Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-167">
        - Content</span></span></td>
    <td><span data-ttu-id="7e98d-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-168">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e98d-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-169">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e98d-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-170">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e98d-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-171">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e98d-172">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-172">-BindingEvents</span></span><br><span data-ttu-id="7e98d-173">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-173">
        -DocumentEvents</span></span><br><span data-ttu-id="7e98d-174">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-174">
        -MatrixBindings</span></span><br><span data-ttu-id="7e98d-175">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-175">
        -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-176">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-176">
        -TableBindings</span></span><br><span data-ttu-id="7e98d-177">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-177">
        -TableCoercion</span></span><br><span data-ttu-id="7e98d-178">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-178">
        -TextBindings</span></span><br><span data-ttu-id="7e98d-179">
        - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-179">
        - Settings</span></span><br><span data-ttu-id="7e98d-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-181">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="7e98d-181">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="7e98d-182">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-182">- Taskpane</span></span><br><span data-ttu-id="7e98d-183">
        - Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-183">
        - Content</span></span><br><span data-ttu-id="7e98d-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-184">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7e98d-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-185">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7e98d-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-186">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7e98d-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-187">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7e98d-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-188">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7e98d-189">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-189">-BindingEvents</span></span><br><span data-ttu-id="7e98d-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-190">
        -DocumentEvents</span></span><br><span data-ttu-id="7e98d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-191">
        -MatrixBindings</span></span><br><span data-ttu-id="7e98d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-192">
        -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-193">
        -TableBindings</span></span><br><span data-ttu-id="7e98d-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-194">
        -TableCoercion</span></span><br><span data-ttu-id="7e98d-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-195">
        -TextBindings</span></span><br><span data-ttu-id="7e98d-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-196">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="7e98d-197">Outlook</span><span class="sxs-lookup"><span data-stu-id="7e98d-197">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e98d-198">Plataforma</span><span class="sxs-lookup"><span data-stu-id="7e98d-198">Platform</span></span></th>
    <th><span data-ttu-id="7e98d-199">Pontos de extens?o</span><span class="sxs-lookup"><span data-stu-id="7e98d-199">Extension points</span></span></th> 
    <th><span data-ttu-id="7e98d-200">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-200">API requirement sets</span></span></th> 
    <th><span data-ttu-id="7e98d-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="7e98d-201"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-202">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e98d-202">Office Online</span></span></td>
    <td> <span data-ttu-id="7e98d-203">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-203">- Mail Read</span></span><br><span data-ttu-id="7e98d-204">
      - Composi??o de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-204">
      - Mail Compose</span></span><br><span data-ttu-id="7e98d-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-205">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-206">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-207">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-208">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-209">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e98d-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-210">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e98d-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-211">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e98d-212">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-212">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-213">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-213">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-214">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-214">- Mail Read</span></span><br><span data-ttu-id="7e98d-215">
      - Composi??o de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-215">
      - Mail Compose</span></span><br><span data-ttu-id="7e98d-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-216">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-217">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-218">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-219">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-220">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="7e98d-221">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-221">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-222">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-222">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-223">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-223">- Mail Read</span></span><br><span data-ttu-id="7e98d-224">
      - Composi??o de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-224">
      - Mail Compose</span></span><br><span data-ttu-id="7e98d-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-225">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7e98d-226">
      - M?dulos</span><span class="sxs-lookup"><span data-stu-id="7e98d-226">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7e98d-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-227">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-228">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-229">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-230">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e98d-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-231">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e98d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e98d-233">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-233">not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-234">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="7e98d-234">Office for iOS</span></span></td>
    <td> <span data-ttu-id="7e98d-235">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-235">- Mail Read</span></span><br><span data-ttu-id="7e98d-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-236">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-237">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-238">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-239">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-240">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e98d-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-241">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="7e98d-242">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-242">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-243">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="7e98d-243">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e98d-244">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-244">- Mail Read</span></span><br><span data-ttu-id="7e98d-245">
      - Composi??o de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-245">
      - Mail Compose</span></span><br><span data-ttu-id="7e98d-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-246">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-247">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-248">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-249">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-250">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e98d-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-251">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7e98d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7e98d-253">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-253">not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-254">Office para Android</span><span class="sxs-lookup"><span data-stu-id="7e98d-254">Office for Android</span></span></td>
    <td> <span data-ttu-id="7e98d-255">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="7e98d-255">- Mail Read</span></span><br><span data-ttu-id="7e98d-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-256">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-257">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7e98d-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-258">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7e98d-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-259">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7e98d-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-260">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7e98d-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-261">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7e98d-262">n?o dispon?vel</span><span class="sxs-lookup"><span data-stu-id="7e98d-262">not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="7e98d-263">Word</span><span class="sxs-lookup"><span data-stu-id="7e98d-263">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e98d-264">Plataforma</span><span class="sxs-lookup"><span data-stu-id="7e98d-264">Platform</span></span></th>
    <th><span data-ttu-id="7e98d-265">Pontos de extens?o</span><span class="sxs-lookup"><span data-stu-id="7e98d-265">Extension points</span></span></th> 
    <th><span data-ttu-id="7e98d-266">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-266">API requirement sets</span></span></th> 
    <th><span data-ttu-id="7e98d-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="7e98d-267"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-268">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e98d-268">Office Online</span></span></td>
    <td> <span data-ttu-id="7e98d-269">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-269">- Taskpane</span></span><br><span data-ttu-id="7e98d-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-270">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-271">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e98d-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-272">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e98d-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-273">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e98d-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-274">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-275">-BindingEvents</span></span><br><span data-ttu-id="7e98d-276">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7e98d-276">customXmlParts</span></span><br><span data-ttu-id="7e98d-277">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-277">
         -MatrixBindings</span></span><br><span data-ttu-id="7e98d-278">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-278">
         -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-279">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-279">
         -TableBindings</span></span><br><span data-ttu-id="7e98d-280">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-280">
         -TableCoercion</span></span><br><span data-ttu-id="7e98d-281">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-281">
         -TextBindings</span></span><br><span data-ttu-id="7e98d-282">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-282">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-283">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-283">
         -TextFile</span></span><br><span data-ttu-id="7e98d-284">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-284">
         -ImageCoercion</span></span><br><span data-ttu-id="7e98d-285">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-285">
         - Settings</span></span><br><span data-ttu-id="7e98d-286">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-286">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-287">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-287">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-288">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-288">- Taskpane</span></span></td>
    <td> <span data-ttu-id="7e98d-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-289">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-290">-BindingEvents</span></span><br><span data-ttu-id="7e98d-291">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-291">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-292">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="7e98d-292">
         -CustomXmlPart</span></span><br><span data-ttu-id="7e98d-293">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-293">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-294">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-294">
         - File</span></span><br><span data-ttu-id="7e98d-295">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-295">
         -HtmlCoercion</span></span><br><span data-ttu-id="7e98d-296">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-296">
         -ImageCoercion</span></span><br><span data-ttu-id="7e98d-297">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-297">
         -OoxmlCoercion</span></span><br><span data-ttu-id="7e98d-298">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-298">
         -TableBindings</span></span><br><span data-ttu-id="7e98d-299">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-299">
         -TableCoercion</span></span><br><span data-ttu-id="7e98d-300">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-300">
         -TextBindings</span></span><br><span data-ttu-id="7e98d-301">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-301">
         -TextFile</span></span><br><span data-ttu-id="7e98d-302">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-302">
         - Settings</span></span><br><span data-ttu-id="7e98d-303">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-303">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-304">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-304">
         -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-305">
         - Associa??es de matriz</span><span class="sxs-lookup"><span data-stu-id="7e98d-305">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-306">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-306">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-307">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-307">- Taskpane</span></span><br><span data-ttu-id="7e98d-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-308">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-309">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e98d-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-310">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e98d-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-311">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e98d-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-312">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-313">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-313">-BindingEvents</span></span><br><span data-ttu-id="7e98d-314">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-314">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-315">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="7e98d-315">
         -CustomXmlPart</span></span><br><span data-ttu-id="7e98d-316">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-316">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-317">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-317">
         - File</span></span><br><span data-ttu-id="7e98d-318">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-318">
         -HtmlCoercion</span></span><br><span data-ttu-id="7e98d-319">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-319">
         -ImageCoercion</span></span><br><span data-ttu-id="7e98d-320">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-320">
         -OoxmlCoercion</span></span><br><span data-ttu-id="7e98d-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-321">
         -TableBindings</span></span><br><span data-ttu-id="7e98d-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-322">
         -TableCoercion</span></span><br><span data-ttu-id="7e98d-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-323">
         -TextBindings</span></span><br><span data-ttu-id="7e98d-324">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-324">
         -TextFile</span></span><br><span data-ttu-id="7e98d-325">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-325">
         - Settings</span></span><br><span data-ttu-id="7e98d-326">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-326">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-327">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-327">
         -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-328">
         - Associa??es de matriz</span><span class="sxs-lookup"><span data-stu-id="7e98d-328">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-329">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="7e98d-329">Office for iOS</span></span></td>
    <td> <span data-ttu-id="7e98d-330">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-330">- Taskpane</span></span></td>
    <td> <span data-ttu-id="7e98d-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-331">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e98d-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-332">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e98d-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-333">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e98d-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7e98d-334">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7e98d-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-335">-BindingEvents</span></span><br><span data-ttu-id="7e98d-336">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-336">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-337">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="7e98d-337">
         -CustomXmlPart</span></span><br><span data-ttu-id="7e98d-338">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-338">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-339">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-339">
         - File</span></span><br><span data-ttu-id="7e98d-340">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-340">
         -HtmlCoercion</span></span><br><span data-ttu-id="7e98d-341">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-341">
         -ImageCoercion</span></span><br><span data-ttu-id="7e98d-342">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-342">
         -OoxmlCoercion</span></span><br><span data-ttu-id="7e98d-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-343">
         -TableBindings</span></span><br><span data-ttu-id="7e98d-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-344">
         -TableCoercion</span></span><br><span data-ttu-id="7e98d-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-345">
         -TextBindings</span></span><br><span data-ttu-id="7e98d-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-346">
         -TextFile</span></span><br><span data-ttu-id="7e98d-347">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-347">
         - Settings</span></span><br><span data-ttu-id="7e98d-348">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-348">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-349">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-349">
         -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-350">
         - Associa??es de matriz</span><span class="sxs-lookup"><span data-stu-id="7e98d-350">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-351">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="7e98d-351">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e98d-352">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-352">- Taskpane</span></span><br><span data-ttu-id="7e98d-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-353">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-354">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="7e98d-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-355">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="7e98d-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-356">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="7e98d-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7e98d-357">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7e98d-358">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-358">-BindingEvents</span></span><br><span data-ttu-id="7e98d-359">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-359">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-360">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="7e98d-360">
         -CustomXmlPart</span></span><br><span data-ttu-id="7e98d-361">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-361">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-362">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-362">
         - File</span></span><br><span data-ttu-id="7e98d-363">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-363">
         -HtmlCoercion</span></span><br><span data-ttu-id="7e98d-364">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-364">
         -ImageCoercion</span></span><br><span data-ttu-id="7e98d-365">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-365">
         -OoxmlCoercion</span></span><br><span data-ttu-id="7e98d-366">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-366">
         -TableBindings</span></span><br><span data-ttu-id="7e98d-367">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-367">
         -TableCoercion</span></span><br><span data-ttu-id="7e98d-368">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7e98d-368">
         -TextBindings</span></span><br><span data-ttu-id="7e98d-369">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-369">
         -TextFile</span></span><br><span data-ttu-id="7e98d-370">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-370">
         - Settings</span></span><br><span data-ttu-id="7e98d-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-371">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-372">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-372">
         -MatrixCoercion</span></span><br><span data-ttu-id="7e98d-373">
         - Associa??es de matriz</span><span class="sxs-lookup"><span data-stu-id="7e98d-373">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7e98d-374">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7e98d-374">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e98d-375">Plataforma</span><span class="sxs-lookup"><span data-stu-id="7e98d-375">Platform</span></span></th>
    <th><span data-ttu-id="7e98d-376">Pontos de extens?o</span><span class="sxs-lookup"><span data-stu-id="7e98d-376">Extension points</span></span></th> 
    <th><span data-ttu-id="7e98d-377">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-377">API requirement sets</span></span></th> 
    <th><span data-ttu-id="7e98d-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="7e98d-378"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-379">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e98d-379">Office Online</span></span></td>
    <td> <span data-ttu-id="7e98d-380">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-380">- Content</span></span><br><span data-ttu-id="7e98d-381">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-381">
         - Taskpane</span></span><br><span data-ttu-id="7e98d-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-382">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-383">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-384">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e98d-384">-ActiveView</span></span><br><span data-ttu-id="7e98d-385">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-385">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-386">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-386">
         - File</span></span><br><span data-ttu-id="7e98d-387">
         - Sele??o</span><span class="sxs-lookup"><span data-stu-id="7e98d-387">
         - Selection</span></span><br><span data-ttu-id="7e98d-388">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-388">
         - Settings</span></span><br><span data-ttu-id="7e98d-389">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-389">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-390">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-390">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-391">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-391">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-392">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-392">- Content</span></span><br><span data-ttu-id="7e98d-393">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-393">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="7e98d-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="7e98d-394">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="7e98d-395">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e98d-395">-ActiveView</span></span><br><span data-ttu-id="7e98d-396">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-396">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-397">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-397">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-398">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-398">
         - File</span></span><br><span data-ttu-id="7e98d-399">
         - Sele??o</span><span class="sxs-lookup"><span data-stu-id="7e98d-399">
         - Selection</span></span><br><span data-ttu-id="7e98d-400">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-400">
         - Settings</span></span><br><span data-ttu-id="7e98d-401">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-401">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-402">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-402">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="7e98d-403">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-403">- Content</span></span><br><span data-ttu-id="7e98d-404">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-404">
         - Taskpane</span></span><br><span data-ttu-id="7e98d-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-405">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-406">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-407">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e98d-407">-ActiveView</span></span><br><span data-ttu-id="7e98d-408">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-408">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-409">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-409">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-410">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-410">
         - File</span></span><br><span data-ttu-id="7e98d-411">
         - Sele??o</span><span class="sxs-lookup"><span data-stu-id="7e98d-411">
         - Selection</span></span><br><span data-ttu-id="7e98d-412">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-412">
         - Settings</span></span><br><span data-ttu-id="7e98d-413">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-413">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-414">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-414">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-415">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="7e98d-415">Office for iOS</span></span></td>
    <td> <span data-ttu-id="7e98d-416">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-416">- Content</span></span><br><span data-ttu-id="7e98d-417">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-417">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="7e98d-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-418">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="7e98d-419">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e98d-419">-ActiveView</span></span><br><span data-ttu-id="7e98d-420">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-420">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-421">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-421">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-422">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-422">
         - File</span></span><br><span data-ttu-id="7e98d-423">
         - Sele??o</span><span class="sxs-lookup"><span data-stu-id="7e98d-423">
         - Selection</span></span><br><span data-ttu-id="7e98d-424">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-424">
         - Settings</span></span><br><span data-ttu-id="7e98d-425">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-425">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-426">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-426">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-427">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="7e98d-427">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="7e98d-428">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-428">- Content</span></span><br><span data-ttu-id="7e98d-429">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-429">
         - Taskpane</span></span><br><span data-ttu-id="7e98d-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-430">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-431">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7e98d-432">-ActiveView</span></span><br><span data-ttu-id="7e98d-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7e98d-433">
         -CompressedFile</span></span><br><span data-ttu-id="7e98d-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-434">
         -DocumentEvents</span></span><br><span data-ttu-id="7e98d-435">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="7e98d-435">
         - File</span></span><br><span data-ttu-id="7e98d-436">
         - Sele??o</span><span class="sxs-lookup"><span data-stu-id="7e98d-436">
         - Selection</span></span><br><span data-ttu-id="7e98d-437">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-437">
         - Settings</span></span><br><span data-ttu-id="7e98d-438">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-438">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-439">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-439">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="7e98d-440">OneNote</span><span class="sxs-lookup"><span data-stu-id="7e98d-440">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7e98d-441">Plataforma</span><span class="sxs-lookup"><span data-stu-id="7e98d-441">Platform</span></span></th>
    <th><span data-ttu-id="7e98d-442">Pontos de extens?o</span><span class="sxs-lookup"><span data-stu-id="7e98d-442">Extension points</span></span></th> 
    <th><span data-ttu-id="7e98d-443">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-443">API requirement sets</span></span></th> 
    <th><span data-ttu-id="7e98d-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="7e98d-444"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-445">Office Online</span><span class="sxs-lookup"><span data-stu-id="7e98d-445">Office Online</span></span></td>
    <td> <span data-ttu-id="7e98d-446">- Conte?do</span><span class="sxs-lookup"><span data-stu-id="7e98d-446">- Content</span></span><br><span data-ttu-id="7e98d-447">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="7e98d-447">
         - Taskpane</span></span><br><span data-ttu-id="7e98d-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-448">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7e98d-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-449">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7e98d-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7e98d-450">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7e98d-451">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7e98d-451">-DocumentEvents</span></span><br><span data-ttu-id="7e98d-452">
         - Configura??es</span><span class="sxs-lookup"><span data-stu-id="7e98d-452">
         - Settings</span></span><br><span data-ttu-id="7e98d-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-453">
         -TextCoercion</span></span><br><span data-ttu-id="7e98d-454">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-454">
         -HtmlCoercion</span></span><br><span data-ttu-id="7e98d-455">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="7e98d-455">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-456">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-456">Office 2013 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td><span data-ttu-id="7e98d-457">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="7e98d-457">Office 2016 for Windows</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-458">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="7e98d-458">Office for iOS</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7e98d-459">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="7e98d-459">Office 2016 for Mac</span></span></td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

<span data-ttu-id="7e98d-460">\* = Estamos trabalhando nisso.</span><span class="sxs-lookup"><span data-stu-id="7e98d-460">\* = We're working on it.</span></span> 

## <a name="see-also"></a><span data-ttu-id="7e98d-461">Veja tamb?m</span><span class="sxs-lookup"><span data-stu-id="7e98d-461">See also</span></span>

- [<span data-ttu-id="7e98d-462">Vis?o geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="7e98d-462">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7e98d-463">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="7e98d-463">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="7e98d-464">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="7e98d-464">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="7e98d-465">Refer?ncia da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="7e98d-465">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

