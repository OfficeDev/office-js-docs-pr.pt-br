---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 07/31/2018
ms.openlocfilehash: 084029c0a5b70b73eaa0b3fcc180f4a813fb8b72
ms.sourcegitcommit: bc68b4cf811b45e8b8d1cbd7c8d2867359ab671b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/02/2018
ms.locfileid: "21703907"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="354df-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="354df-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="354df-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="354df-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="354df-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="354df-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span> 

<span data-ttu-id="354df-106">Se uma célula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="354df-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="354df-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="354df-107">For requirement sets for Project or Access, see [Office common requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="354df-p103">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="354df-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="354df-110">Excel</span><span class="sxs-lookup"><span data-stu-id="354df-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="354df-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="354df-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="354df-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="354df-112">Extension points</span></span></th> 
    <th style="width:20%"><span data-ttu-id="354df-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="354df-113">API requirement sets</span></span></th> 
    <th style="width:40%"><span data-ttu-id="354df-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="354df-114"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="354df-115">Office Online</span></span></td>
    <td> <span data-ttu-id="354df-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-116">- Taskpane</span></span><br><span data-ttu-id="354df-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-117">
        - Content</span></span><br><span data-ttu-id="354df-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="354df-118">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="354df-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-119">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="354df-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-120">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="354df-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-121">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="354df-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-122">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="354df-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-123">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="354df-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-124">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="354df-125">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="354df-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="354df-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-126">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="354df-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-127">
        -BindingEvents</span></span><br><span data-ttu-id="354df-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-128">
        -DocumentEvents</span></span><br><span data-ttu-id="354df-129">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-129">
        -MatrixBindings</span></span><br><span data-ttu-id="354df-130">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-130">
        -MatrixCoercion</span></span><br><span data-ttu-id="354df-131">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-131">
        -TableBindings</span></span><br><span data-ttu-id="354df-132">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-132">
        -TableCoercion</span></span><br><span data-ttu-id="354df-133">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-133">
        -TextBindings</span></span><br><span data-ttu-id="354df-134">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-134">
        -CompressedFile</span></span><br><span data-ttu-id="354df-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-135">
        - Settings</span></span><br><span data-ttu-id="354df-136">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-136">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-137">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-137">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="354df-138">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-138">
        - Taskpane</span></span><br><span data-ttu-id="354df-139">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-139">
        - Content</span></span></td>
    <td>  <span data-ttu-id="354df-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-140">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="354df-141">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-141">
        -BindingEvents</span></span><br><span data-ttu-id="354df-142">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-142">
        -DocumentEvents</span></span><br><span data-ttu-id="354df-143">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-143">
        -MatrixBindings</span></span><br><span data-ttu-id="354df-144">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-144">
        -MatrixCoercion</span></span><br><span data-ttu-id="354df-145">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-145">
        -TableBindings</span></span><br><span data-ttu-id="354df-146">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-146">
        -TableCoercion</span></span><br><span data-ttu-id="354df-147">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-147">
        -TextBindings</span></span><br><span data-ttu-id="354df-148">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-148">
        - Settings</span></span><br><span data-ttu-id="354df-149">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-149">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-150">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-150">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="354df-151">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-151">- Taskpane</span></span><br><span data-ttu-id="354df-152">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-152">
        - Content</span></span><br><span data-ttu-id="354df-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-153">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="354df-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-154">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="354df-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-155">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="354df-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-156">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="354df-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-157">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="354df-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-158">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="354df-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-159">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="354df-160">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="354df-160">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="354df-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-161">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="354df-162">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-162">-BindingEvents</span></span><br><span data-ttu-id="354df-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-163">
        -DocumentEvents</span></span><br><span data-ttu-id="354df-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-164">
        -MatrixBindings</span></span><br><span data-ttu-id="354df-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-165">
        -MatrixCoercion</span></span><br><span data-ttu-id="354df-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-166">
        -TableBindings</span></span><br><span data-ttu-id="354df-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-167">
        -TableCoercion</span></span><br><span data-ttu-id="354df-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-168">
        -TextBindings</span></span><br><span data-ttu-id="354df-169">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-169">
        - Settings</span></span><br><span data-ttu-id="354df-170">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-170">
        -TextCoercion</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-171">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="354df-171">Office for iOS</span></span></td>
    <td><span data-ttu-id="354df-172">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-172">- Taskpane</span></span><br><span data-ttu-id="354df-173">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-173">
        - Content</span></span></td>
    <td><span data-ttu-id="354df-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-174">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="354df-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-175">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="354df-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-176">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="354df-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-177">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="354df-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-178">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="354df-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-179">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="354df-180">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="354df-180">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="354df-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-181">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="354df-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-182">-BindingEvents</span></span><br><span data-ttu-id="354df-183">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-183">
        -DocumentEvents</span></span><br><span data-ttu-id="354df-184">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-184">
        -MatrixBindings</span></span><br><span data-ttu-id="354df-185">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-185">
        -MatrixCoercion</span></span><br><span data-ttu-id="354df-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-186">
        -TableBindings</span></span><br><span data-ttu-id="354df-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-187">
        -TableCoercion</span></span><br><span data-ttu-id="354df-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-188">
        -TextBindings</span></span><br><span data-ttu-id="354df-189">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-189">
        - Settings</span></span><br><span data-ttu-id="354df-190">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-190">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-191">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="354df-191">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="354df-192">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-192">- Taskpane</span></span><br><span data-ttu-id="354df-193">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-193">
        - Content</span></span><br><span data-ttu-id="354df-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-194">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="354df-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-195">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="354df-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-196">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="354df-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-197">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="354df-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-198">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="354df-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-199">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="354df-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-200">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="354df-201">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="354df-201">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="354df-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-202">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="354df-203">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-203">-BindingEvents</span></span><br><span data-ttu-id="354df-204">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-204">
        -DocumentEvents</span></span><br><span data-ttu-id="354df-205">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-205">
        -MatrixBindings</span></span><br><span data-ttu-id="354df-206">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-206">
        -MatrixCoercion</span></span><br><span data-ttu-id="354df-207">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-207">
        -TableBindings</span></span><br><span data-ttu-id="354df-208">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-208">
        -TableCoercion</span></span><br><span data-ttu-id="354df-209">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-209">
        -TextBindings</span></span><br><span data-ttu-id="354df-210">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-210">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="354df-211">Outlook</span><span class="sxs-lookup"><span data-stu-id="354df-211">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="354df-212">Plataforma</span><span class="sxs-lookup"><span data-stu-id="354df-212">Platform</span></span></th>
    <th><span data-ttu-id="354df-213">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="354df-213">Extension points</span></span></th> 
    <th><span data-ttu-id="354df-214">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="354df-214">API requirement sets</span></span></th> 
    <th><span data-ttu-id="354df-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="354df-215"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-216">Office Online</span><span class="sxs-lookup"><span data-stu-id="354df-216">Office Online</span></span></td>
    <td> <span data-ttu-id="354df-217">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-217">- Mail Read</span></span><br><span data-ttu-id="354df-218">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="354df-218">
      - Mail Compose</span></span><br><span data-ttu-id="354df-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-219">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-220">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-221">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-222">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-223">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="354df-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-224">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="354df-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-225">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="354df-226">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-226">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-227">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-227">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-228">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-228">- Mail Read</span></span><br><span data-ttu-id="354df-229">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="354df-229">
      - Mail Compose</span></span><br><span data-ttu-id="354df-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-230">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-231">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-232">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-233">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-234">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="354df-235">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-235">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-236">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-236">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-237">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-237">- Mail Read</span></span><br><span data-ttu-id="354df-238">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="354df-238">
      - Mail Compose</span></span><br><span data-ttu-id="354df-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-239">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="354df-240">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="354df-240">
      - Modules</span></span></td>
    <td> <span data-ttu-id="354df-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-241">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-242">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-243">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-244">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="354df-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-245">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="354df-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-246">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="354df-247">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-247">Not available</span></span></td> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-248">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="354df-248">Office for iOS</span></span></td>
    <td> <span data-ttu-id="354df-249">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-249">- Mail Read</span></span><br><span data-ttu-id="354df-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-250">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-251">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-252">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-253">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-254">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="354df-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-255">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>    
    <td><span data-ttu-id="354df-256">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-256">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-257">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="354df-257">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="354df-258">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-258">- Mail Read</span></span><br><span data-ttu-id="354df-259">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="354df-259">
      - Mail Compose</span></span><br><span data-ttu-id="354df-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-260">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-261">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-262">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-263">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-264">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="354df-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-265">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="354df-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="354df-266">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="354df-267">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-268">Office para Android</span><span class="sxs-lookup"><span data-stu-id="354df-268">Office for Android</span></span></td>
    <td> <span data-ttu-id="354df-269">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="354df-269">- Mail Read</span></span><br><span data-ttu-id="354df-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-270">
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-271">- <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="354df-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-272">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="354df-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-273">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="354df-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="354df-274">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="354df-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="354df-275">
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="354df-276">Não disponível</span><span class="sxs-lookup"><span data-stu-id="354df-276">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="354df-277">Word</span><span class="sxs-lookup"><span data-stu-id="354df-277">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="354df-278">Plataforma</span><span class="sxs-lookup"><span data-stu-id="354df-278">Platform</span></span></th>
    <th><span data-ttu-id="354df-279">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="354df-279">Extension points</span></span></th> 
    <th><span data-ttu-id="354df-280">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="354df-280">API requirement sets</span></span></th> 
    <th><span data-ttu-id="354df-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="354df-281"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-282">Office Online</span><span class="sxs-lookup"><span data-stu-id="354df-282">Office Online</span></span></td>
    <td> <span data-ttu-id="354df-283">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-283">- Taskpane</span></span><br><span data-ttu-id="354df-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-284">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-285">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="354df-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-286">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="354df-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-287">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="354df-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-288">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-289">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-289">-BindingEvents</span></span><br><span data-ttu-id="354df-290">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="354df-290">customXmlParts</span></span><br><span data-ttu-id="354df-291">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="354df-291">
         -MatrixBindings</span></span><br><span data-ttu-id="354df-292">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-292">
         -MatrixCoercion</span></span><br><span data-ttu-id="354df-293">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-293">
         -TableBindings</span></span><br><span data-ttu-id="354df-294">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-294">
         -TableCoercion</span></span><br><span data-ttu-id="354df-295">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-295">
         -TextBindings</span></span><br><span data-ttu-id="354df-296">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-296">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-297">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="354df-297">
         -TextFile</span></span><br><span data-ttu-id="354df-298">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-298">
         -ImageCoercion</span></span><br><span data-ttu-id="354df-299">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-299">
         - Settings</span></span><br><span data-ttu-id="354df-300">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-300">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-301">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-301">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-302">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-302">- Taskpane</span></span></td>
    <td> <span data-ttu-id="354df-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-303">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-304">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-304">-BindingEvents</span></span><br><span data-ttu-id="354df-305">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-305">
         -CompressedFile</span></span><br><span data-ttu-id="354df-306">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="354df-306">
         -CustomXmlPart</span></span><br><span data-ttu-id="354df-307">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-307">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-308">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-308">
         - File</span></span><br><span data-ttu-id="354df-309">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-309">
         -HtmlCoercion</span></span><br><span data-ttu-id="354df-310">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-310">
         -ImageCoercion</span></span><br><span data-ttu-id="354df-311">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-311">
         -OoxmlCoercion</span></span><br><span data-ttu-id="354df-312">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-312">
         -TableBindings</span></span><br><span data-ttu-id="354df-313">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-313">
         -TableCoercion</span></span><br><span data-ttu-id="354df-314">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-314">
         -TextBindings</span></span><br><span data-ttu-id="354df-315">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="354df-315">
         -TextFile</span></span><br><span data-ttu-id="354df-316">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-316">
         - Settings</span></span><br><span data-ttu-id="354df-317">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-317">
         -TextCoercion</span></span><br><span data-ttu-id="354df-318">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-318">
         -MatrixCoercion</span></span><br><span data-ttu-id="354df-319">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="354df-319">
         - Matrix Bindings</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-320">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-320">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-321">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-321">- Taskpane</span></span><br><span data-ttu-id="354df-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-322">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-323">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="354df-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-324">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="354df-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-325">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="354df-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-326">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-327">-BindingEvents</span></span><br><span data-ttu-id="354df-328">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-328">
         -CompressedFile</span></span><br><span data-ttu-id="354df-329">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="354df-329">
         -CustomXmlPart</span></span><br><span data-ttu-id="354df-330">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-330">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-331">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-331">
         - File</span></span><br><span data-ttu-id="354df-332">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-332">
         -HtmlCoercion</span></span><br><span data-ttu-id="354df-333">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-333">
         -ImageCoercion</span></span><br><span data-ttu-id="354df-334">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-334">
         -OoxmlCoercion</span></span><br><span data-ttu-id="354df-335">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-335">
         -TableBindings</span></span><br><span data-ttu-id="354df-336">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-336">
         -TableCoercion</span></span><br><span data-ttu-id="354df-337">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-337">
         -TextBindings</span></span><br><span data-ttu-id="354df-338">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="354df-338">
         -TextFile</span></span><br><span data-ttu-id="354df-339">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-339">
         - Settings</span></span><br><span data-ttu-id="354df-340">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-340">
         -TextCoercion</span></span><br><span data-ttu-id="354df-341">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-341">
         -MatrixCoercion</span></span><br><span data-ttu-id="354df-342">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="354df-342">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-343">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="354df-343">Office for iOS</span></span></td>
    <td> <span data-ttu-id="354df-344">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-344">- Taskpane</span></span></td>
    <td> <span data-ttu-id="354df-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-345">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="354df-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-346">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="354df-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-347">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="354df-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="354df-348">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="354df-349">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-349">-BindingEvents</span></span><br><span data-ttu-id="354df-350">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-350">
         -CompressedFile</span></span><br><span data-ttu-id="354df-351">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="354df-351">
         -CustomXmlPart</span></span><br><span data-ttu-id="354df-352">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-352">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-353">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-353">
         - File</span></span><br><span data-ttu-id="354df-354">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-354">
         -HtmlCoercion</span></span><br><span data-ttu-id="354df-355">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-355">
         -ImageCoercion</span></span><br><span data-ttu-id="354df-356">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-356">
         -OoxmlCoercion</span></span><br><span data-ttu-id="354df-357">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-357">
         -TableBindings</span></span><br><span data-ttu-id="354df-358">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-358">
         -TableCoercion</span></span><br><span data-ttu-id="354df-359">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-359">
         -TextBindings</span></span><br><span data-ttu-id="354df-360">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="354df-360">
         -TextFile</span></span><br><span data-ttu-id="354df-361">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-361">
         - Settings</span></span><br><span data-ttu-id="354df-362">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-362">
         -TextCoercion</span></span><br><span data-ttu-id="354df-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="354df-364">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="354df-364">
         - Matrix Bindings</span></span> </td> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-365">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="354df-365">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="354df-366">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-366">- Taskpane</span></span><br><span data-ttu-id="354df-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-367">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-368">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="354df-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="354df-369">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="354df-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="354df-370">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="354df-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="354df-371">
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="354df-372">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="354df-372">-BindingEvents</span></span><br><span data-ttu-id="354df-373">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-373">
         -CompressedFile</span></span><br><span data-ttu-id="354df-374">
         - CustomXmlPart</span><span class="sxs-lookup"><span data-stu-id="354df-374">
         -CustomXmlPart</span></span><br><span data-ttu-id="354df-375">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-375">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-376">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-376">
         - File</span></span><br><span data-ttu-id="354df-377">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-377">
         -HtmlCoercion</span></span><br><span data-ttu-id="354df-378">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-378">
         -ImageCoercion</span></span><br><span data-ttu-id="354df-379">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-379">
         -OoxmlCoercion</span></span><br><span data-ttu-id="354df-380">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="354df-380">
         -TableBindings</span></span><br><span data-ttu-id="354df-381">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-381">
         -TableCoercion</span></span><br><span data-ttu-id="354df-382">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="354df-382">
         -TextBindings</span></span><br><span data-ttu-id="354df-383">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="354df-383">
         -TextFile</span></span><br><span data-ttu-id="354df-384">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-384">
         - Settings</span></span><br><span data-ttu-id="354df-385">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-385">
         -TextCoercion</span></span><br><span data-ttu-id="354df-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="354df-387">
         - Associações de matriz</span><span class="sxs-lookup"><span data-stu-id="354df-387">
         - Matrix Bindings</span></span> </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="354df-388">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="354df-388">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="354df-389">Plataforma</span><span class="sxs-lookup"><span data-stu-id="354df-389">Platform</span></span></th>
    <th><span data-ttu-id="354df-390">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="354df-390">Extension points</span></span></th> 
    <th><span data-ttu-id="354df-391">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="354df-391">API requirement sets</span></span></th> 
    <th><span data-ttu-id="354df-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="354df-392"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-393">Office Online</span><span class="sxs-lookup"><span data-stu-id="354df-393">Office Online</span></span></td>
    <td> <span data-ttu-id="354df-394">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-394">- Content</span></span><br><span data-ttu-id="354df-395">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-395">
         - Taskpane</span></span><br><span data-ttu-id="354df-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-396">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-397">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-398">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="354df-398">-ActiveView</span></span><br><span data-ttu-id="354df-399">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-399">
         -CompressedFile</span></span><br><span data-ttu-id="354df-400">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-400">
         - File</span></span><br><span data-ttu-id="354df-401">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="354df-401">
         - Selection</span></span><br><span data-ttu-id="354df-402">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-402">
         - Settings</span></span><br><span data-ttu-id="354df-403">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-403">
         -TextCoercion</span></span><br><span data-ttu-id="354df-404">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-404">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-405">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-405">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-406">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-406">- Content</span></span><br><span data-ttu-id="354df-407">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-407">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="354df-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="354df-408">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="354df-409">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="354df-409">-ActiveView</span></span><br><span data-ttu-id="354df-410">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-410">
         -CompressedFile</span></span><br><span data-ttu-id="354df-411">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-411">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-412">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-412">
         - File</span></span><br><span data-ttu-id="354df-413">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="354df-413">
         - Selection</span></span><br><span data-ttu-id="354df-414">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-414">
         - Settings</span></span><br><span data-ttu-id="354df-415">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-415">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-416">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="354df-416">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="354df-417">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-417">- Content</span></span><br><span data-ttu-id="354df-418">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-418">
         - Taskpane</span></span><br><span data-ttu-id="354df-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-419">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-420">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-421">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="354df-421">-ActiveView</span></span><br><span data-ttu-id="354df-422">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-422">
         -CompressedFile</span></span><br><span data-ttu-id="354df-423">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-423">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-424">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-424">
         - File</span></span><br><span data-ttu-id="354df-425">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="354df-425">
         - Selection</span></span><br><span data-ttu-id="354df-426">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-426">
         - Settings</span></span><br><span data-ttu-id="354df-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-427">
         -TextCoercion</span></span><br><span data-ttu-id="354df-428">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-428">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-429">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="354df-429">Office for iOS</span></span></td>
    <td> <span data-ttu-id="354df-430">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-430">- Content</span></span><br><span data-ttu-id="354df-431">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-431">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="354df-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-432">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="354df-433">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="354df-433">-ActiveView</span></span><br><span data-ttu-id="354df-434">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-434">
         -CompressedFile</span></span><br><span data-ttu-id="354df-435">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-435">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-436">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-436">
         - File</span></span><br><span data-ttu-id="354df-437">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="354df-437">
         - Selection</span></span><br><span data-ttu-id="354df-438">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-438">
         - Settings</span></span><br><span data-ttu-id="354df-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-439">
         -TextCoercion</span></span><br><span data-ttu-id="354df-440">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-440">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="354df-441">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="354df-441">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="354df-442">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-442">- Content</span></span><br><span data-ttu-id="354df-443">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-443">
         - Taskpane</span></span><br><span data-ttu-id="354df-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-444">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-445">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-446">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="354df-446">-ActiveView</span></span><br><span data-ttu-id="354df-447">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="354df-447">
         -CompressedFile</span></span><br><span data-ttu-id="354df-448">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-448">
         -DocumentEvents</span></span><br><span data-ttu-id="354df-449">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="354df-449">
         - File</span></span><br><span data-ttu-id="354df-450">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="354df-450">
         - Selection</span></span><br><span data-ttu-id="354df-451">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-451">
         - Settings</span></span><br><span data-ttu-id="354df-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-452">
         -TextCoercion</span></span><br><span data-ttu-id="354df-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-453">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="354df-454">OneNote</span><span class="sxs-lookup"><span data-stu-id="354df-454">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="354df-455">Plataforma</span><span class="sxs-lookup"><span data-stu-id="354df-455">Platform</span></span></th>
    <th><span data-ttu-id="354df-456">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="354df-456">Extension points</span></span></th> 
    <th><span data-ttu-id="354df-457">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="354df-457">API requirement sets</span></span></th> 
    <th><span data-ttu-id="354df-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="354df-458"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th> 
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="354df-459">Office Online</span><span class="sxs-lookup"><span data-stu-id="354df-459">Office Online</span></span></td>
    <td> <span data-ttu-id="354df-460">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="354df-460">- Content</span></span><br><span data-ttu-id="354df-461">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="354df-461">
         - Taskpane</span></span><br><span data-ttu-id="354df-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="354df-462">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="354df-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-463">- <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="354df-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="354df-464">
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="354df-465">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="354df-465">-DocumentEvents</span></span><br><span data-ttu-id="354df-466">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="354df-466">
         - Settings</span></span><br><span data-ttu-id="354df-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-467">
         -TextCoercion</span></span><br><span data-ttu-id="354df-468">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-468">
         -HtmlCoercion</span></span><br><span data-ttu-id="354df-469">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="354df-469">
         -ImageCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="354df-470">Veja também</span><span class="sxs-lookup"><span data-stu-id="354df-470">See also</span></span>

- [<span data-ttu-id="354df-471">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="354df-471">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="354df-472">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="354df-472">Common API requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="354df-473">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="354df-473">Add-in Commands requirement sets</span></span>](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="354df-474">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="354df-474">JavaScript API for Office reference</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)

