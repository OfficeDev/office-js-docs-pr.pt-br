---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 09/19/2018
ms.openlocfilehash: 09fb72c88bd0496c413f94b7ba4149192380d664
ms.sourcegitcommit: e7e4d08569a01c69168bb005188e9a1e628304b9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/22/2018
ms.locfileid: "24967701"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b8f55-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8f55-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b8f55-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="b8f55-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b8f55-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="b8f55-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="b8f55-106">Se uma célula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="b8f55-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="b8f55-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="b8f55-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="b8f55-p103">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="b8f55-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="b8f55-110">Excel</span><span class="sxs-lookup"><span data-stu-id="b8f55-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b8f55-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8f55-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b8f55-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8f55-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b8f55-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b8f55-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8f55-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="b8f55-115">Office Online</span></span></td>
    <td> <span data-ttu-id="b8f55-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-116">- Taskpane</span></span><br><span data-ttu-id="b8f55-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-117">
        - Content</span></span><br><span data-ttu-id="b8f55-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b8f55-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b8f55-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8f55-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8f55-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8f55-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8f55-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8f55-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8f55-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b8f55-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8f55-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-127">
        -BindingEvents</span></span><br><span data-ttu-id="b8f55-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-128">
        -CompressedFile</span></span><br><span data-ttu-id="b8f55-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-129">
        -DocumentEvents</span></span><br><span data-ttu-id="b8f55-130">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-130">
        - File</span></span><br><span data-ttu-id="b8f55-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-131">
        -MatrixBindings</span></span><br><span data-ttu-id="b8f55-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-133">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-133">
        - Selection</span></span><br><span data-ttu-id="b8f55-134">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-134">
        - Settings</span></span><br><span data-ttu-id="b8f55-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-135">
        -TableBindings</span></span><br><span data-ttu-id="b8f55-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-136">
        -TableCoercion</span></span><br><span data-ttu-id="b8f55-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-137">
        -TextBindings</span></span><br><span data-ttu-id="b8f55-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-139">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="b8f55-140">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-140">
        - Taskpane</span></span><br><span data-ttu-id="b8f55-141">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b8f55-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8f55-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-143">
        -BindingEvents</span></span><br><span data-ttu-id="b8f55-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-144">
        -CompressedFile</span></span><br><span data-ttu-id="b8f55-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-145">
        -DocumentEvents</span></span><br><span data-ttu-id="b8f55-146">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-146">
        - File</span></span><br><span data-ttu-id="b8f55-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-147">
        -ImageCoercion</span></span><br><span data-ttu-id="b8f55-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-148">
        -MatrixBindings</span></span><br><span data-ttu-id="b8f55-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-150">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-150">
        - Selection</span></span><br><span data-ttu-id="b8f55-151">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-151">
        - Settings</span></span><br><span data-ttu-id="b8f55-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-152">
        -TableBindings</span></span><br><span data-ttu-id="b8f55-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-153">
        -TableCoercion</span></span><br><span data-ttu-id="b8f55-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-154">
        -TextBindings</span></span><br><span data-ttu-id="b8f55-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-156">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="b8f55-157">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-157">- Taskpane</span></span><br><span data-ttu-id="b8f55-158">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-158">
        - Content</span></span><br><span data-ttu-id="b8f55-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8f55-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8f55-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8f55-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8f55-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8f55-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8f55-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8f55-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b8f55-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8f55-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-168">-BindingEvents</span></span><br><span data-ttu-id="b8f55-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-169">
        -CompressedFile</span></span><br><span data-ttu-id="b8f55-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-170">
        -DocumentEvents</span></span><br><span data-ttu-id="b8f55-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-171">
        - File</span></span><br><span data-ttu-id="b8f55-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-172">
        -ImageCoercion</span></span><br><span data-ttu-id="b8f55-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-173">
        -MatrixBindings</span></span><br><span data-ttu-id="b8f55-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-175">
        - Selection</span></span><br><span data-ttu-id="b8f55-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-176">
        - Settings</span></span><br><span data-ttu-id="b8f55-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-177">
        -TableBindings</span></span><br><span data-ttu-id="b8f55-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-178">
        -TableCoercion</span></span><br><span data-ttu-id="b8f55-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-179">
        -TextBindings</span></span><br><span data-ttu-id="b8f55-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-181">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="b8f55-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="b8f55-182">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-182">- Taskpane</span></span><br><span data-ttu-id="b8f55-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-183">
        - Content</span></span></td>
    <td><span data-ttu-id="b8f55-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8f55-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8f55-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8f55-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8f55-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8f55-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8f55-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b8f55-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8f55-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-192">-BindingEvents</span></span><br><span data-ttu-id="b8f55-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-193">
        -CompressedFile</span></span><br><span data-ttu-id="b8f55-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-194">
        -DocumentEvents</span></span><br><span data-ttu-id="b8f55-195">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-195">
        - File</span></span><br><span data-ttu-id="b8f55-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-196">
        -ImageCoercion</span></span><br><span data-ttu-id="b8f55-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-197">
        -MatrixBindings</span></span><br><span data-ttu-id="b8f55-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-199">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-199">
        - Selection</span></span><br><span data-ttu-id="b8f55-200">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-200">
        - Settings</span></span><br><span data-ttu-id="b8f55-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-201">
        -TableBindings</span></span><br><span data-ttu-id="b8f55-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-202">
        -TableCoercion</span></span><br><span data-ttu-id="b8f55-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-203">
        -TextBindings</span></span><br><span data-ttu-id="b8f55-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-205">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="b8f55-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="b8f55-206">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-206">- Taskpane</span></span><br><span data-ttu-id="b8f55-207">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-207">
        - Content</span></span><br><span data-ttu-id="b8f55-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8f55-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8f55-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8f55-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8f55-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8f55-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8f55-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8f55-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="b8f55-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8f55-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-217">-BindingEvents</span></span><br><span data-ttu-id="b8f55-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-218">
        -CompressedFile</span></span><br><span data-ttu-id="b8f55-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-219">
        -DocumentEvents</span></span><br><span data-ttu-id="b8f55-220">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-220">
        - File</span></span><br><span data-ttu-id="b8f55-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-221">
        -ImageCoercion</span></span><br><span data-ttu-id="b8f55-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-222">
        -MatrixBindings</span></span><br><span data-ttu-id="b8f55-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-224">
        -PdfFile</span></span><br><span data-ttu-id="b8f55-225">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-225">
        - Selection</span></span><br><span data-ttu-id="b8f55-226">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-226">
        - Settings</span></span><br><span data-ttu-id="b8f55-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-227">
        -TableBindings</span></span><br><span data-ttu-id="b8f55-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-228">
        -TableCoercion</span></span><br><span data-ttu-id="b8f55-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-229">
        -TextBindings</span></span><br><span data-ttu-id="b8f55-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="b8f55-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="b8f55-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8f55-232">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8f55-232">Platform</span></span></th>
    <th><span data-ttu-id="b8f55-233">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8f55-233">Extension points</span></span></th>
    <th><span data-ttu-id="b8f55-234">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8f55-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8f55-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="b8f55-236">Office Online</span></span></td>
    <td> <span data-ttu-id="b8f55-237">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-237">- Mail Read</span></span><br><span data-ttu-id="b8f55-238">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-238">
      - Mail Compose</span></span><br><span data-ttu-id="b8f55-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8f55-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8f55-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8f55-246">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-247">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-248">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-248">- Mail Read</span></span><br><span data-ttu-id="b8f55-249">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-249">
      - Mail Compose</span></span><br><span data-ttu-id="b8f55-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="b8f55-255">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-256">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-257">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-257">- Mail Read</span></span><br><span data-ttu-id="b8f55-258">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-258">
      - Mail Compose</span></span><br><span data-ttu-id="b8f55-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8f55-260">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b8f55-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b8f55-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8f55-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8f55-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8f55-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-267">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b8f55-268">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-268">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-269">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="b8f55-269">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b8f55-270">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-270">- Mail Read</span></span><br><span data-ttu-id="b8f55-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-271">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-272">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8f55-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-276">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8f55-277">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-277">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-278">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="b8f55-278">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b8f55-279">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-279">- Mail Read</span></span><br><span data-ttu-id="b8f55-280">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-280">
      - Mail Compose</span></span><br><span data-ttu-id="b8f55-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-281">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-282">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8f55-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8f55-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-287">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8f55-288">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-288">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-289">Office para Android</span><span class="sxs-lookup"><span data-stu-id="b8f55-289">Office for Android</span></span></td>
    <td> <span data-ttu-id="b8f55-290">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8f55-290">- Mail Read</span></span><br><span data-ttu-id="b8f55-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-291">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-292">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8f55-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8f55-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8f55-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8f55-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8f55-297">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8f55-297">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="b8f55-298">Word</span><span class="sxs-lookup"><span data-stu-id="b8f55-298">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8f55-299">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8f55-299">Platform</span></span></th>
    <th><span data-ttu-id="b8f55-300">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8f55-300">Extension points</span></span></th>
    <th><span data-ttu-id="b8f55-301">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-301">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8f55-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8f55-302"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-303">Office Online</span><span class="sxs-lookup"><span data-stu-id="b8f55-303">Office Online</span></span></td>
    <td> <span data-ttu-id="b8f55-304">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-304">- Taskpane</span></span><br><span data-ttu-id="b8f55-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-305">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-306">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b8f55-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b8f55-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b8f55-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-309">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-310">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-310">-BindingEvents</span></span><br><span data-ttu-id="b8f55-311">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8f55-311">customXmlParts</span></span><br><span data-ttu-id="b8f55-312">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-312">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-313">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-313">
         - File</span></span><br><span data-ttu-id="b8f55-314">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-314">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-315">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-315">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-316">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-316">
         -MatrixBindings</span></span><br><span data-ttu-id="b8f55-317">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-317">
         -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-318">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-318">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b8f55-319">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-319">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-320">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-320">
         - Selection</span></span><br><span data-ttu-id="b8f55-321">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-321">
         - Settings</span></span><br><span data-ttu-id="b8f55-322">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-322">
         -TableBindings</span></span><br><span data-ttu-id="b8f55-323">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-323">
         -TableCoercion</span></span><br><span data-ttu-id="b8f55-324">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-324">
         -TextBindings</span></span><br><span data-ttu-id="b8f55-325">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-325">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-326">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-326">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-327">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-327">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-328">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-328">- Taskpane</span></span></td>
    <td> <span data-ttu-id="b8f55-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-329">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-330">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-330">-BindingEvents</span></span><br><span data-ttu-id="b8f55-331">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-331">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-332">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8f55-332">customXmlParts</span></span><br><span data-ttu-id="b8f55-333">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-333">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-334">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-334">
         - File</span></span><br><span data-ttu-id="b8f55-335">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-335">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-336">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-336">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-337">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-337">
         -MatrixBindings</span></span><br><span data-ttu-id="b8f55-338">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-338">
         -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-339">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-339">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b8f55-340">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-340">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-341">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-341">
         - Selection</span></span><br><span data-ttu-id="b8f55-342">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-342">
         - Settings</span></span><br><span data-ttu-id="b8f55-343">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-343">
         -TableBindings</span></span><br><span data-ttu-id="b8f55-344">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-344">
         -TableCoercion</span></span><br><span data-ttu-id="b8f55-345">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-345">
         -TextBindings</span></span><br><span data-ttu-id="b8f55-346">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-346">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-347">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-347">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-348">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-348">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-349">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-349">- Taskpane</span></span><br><span data-ttu-id="b8f55-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-350">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-351">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b8f55-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b8f55-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b8f55-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-354">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-355">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-355">-BindingEvents</span></span><br><span data-ttu-id="b8f55-356">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-356">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-357">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8f55-357">customXmlParts</span></span><br><span data-ttu-id="b8f55-358">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-358">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-359">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-359">
         - File</span></span><br><span data-ttu-id="b8f55-360">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-360">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-361">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-361">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-362">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-362">
         -MatrixBindings</span></span><br><span data-ttu-id="b8f55-363">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-363">
         -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-364">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-364">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b8f55-365">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-365">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-366">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-366">
         - Selection</span></span><br><span data-ttu-id="b8f55-367">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-367">
         - Settings</span></span><br><span data-ttu-id="b8f55-368">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-368">
         -TableBindings</span></span><br><span data-ttu-id="b8f55-369">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-369">
         -TableCoercion</span></span><br><span data-ttu-id="b8f55-370">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-370">
         -TextBindings</span></span><br><span data-ttu-id="b8f55-371">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-371">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-372">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-372">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-373">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="b8f55-373">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b8f55-374">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-374">- Taskpane</span></span></td>
    <td> <span data-ttu-id="b8f55-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-375">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b8f55-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b8f55-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b8f55-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b8f55-378">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b8f55-379">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-379">-BindingEvents</span></span><br><span data-ttu-id="b8f55-380">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-380">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-381">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8f55-381">customXmlParts</span></span><br><span data-ttu-id="b8f55-382">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-382">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-383">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-383">
         - File</span></span><br><span data-ttu-id="b8f55-384">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-384">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-385">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-385">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-386">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-386">
         -MatrixBindings</span></span><br><span data-ttu-id="b8f55-387">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-387">
         -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-388">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-388">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b8f55-389">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-389">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-390">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-390">
         - Selection</span></span><br><span data-ttu-id="b8f55-391">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-391">
         - Settings</span></span><br><span data-ttu-id="b8f55-392">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-392">
         -TableBindings</span></span><br><span data-ttu-id="b8f55-393">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-393">
         -TableCoercion</span></span><br><span data-ttu-id="b8f55-394">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-394">
         -TextBindings</span></span><br><span data-ttu-id="b8f55-395">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-395">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-396">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-396">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-397">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="b8f55-397">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b8f55-398">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-398">- Taskpane</span></span><br><span data-ttu-id="b8f55-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-399">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-400">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="b8f55-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="b8f55-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="b8f55-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b8f55-403">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b8f55-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-404">-BindingEvents</span></span><br><span data-ttu-id="b8f55-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-405">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8f55-406">customXmlParts</span></span><br><span data-ttu-id="b8f55-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-407">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-408">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-408">
         - File</span></span><br><span data-ttu-id="b8f55-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-410">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-411">
         -MatrixBindings</span></span><br><span data-ttu-id="b8f55-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="b8f55-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="b8f55-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-414">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-415">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-415">
         - Selection</span></span><br><span data-ttu-id="b8f55-416">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-416">
         - Settings</span></span><br><span data-ttu-id="b8f55-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-417">
         -TableBindings</span></span><br><span data-ttu-id="b8f55-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-418">
         -TableCoercion</span></span><br><span data-ttu-id="b8f55-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8f55-419">
         -TextBindings</span></span><br><span data-ttu-id="b8f55-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-420">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-421">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b8f55-422">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b8f55-422">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8f55-423">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8f55-423">Platform</span></span></th>
    <th><span data-ttu-id="b8f55-424">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8f55-424">Extension points</span></span></th>
    <th><span data-ttu-id="b8f55-425">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-425">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8f55-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8f55-426"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-427">Office Online</span><span class="sxs-lookup"><span data-stu-id="b8f55-427">Office Online</span></span></td>
    <td> <span data-ttu-id="b8f55-428">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-428">- Content</span></span><br><span data-ttu-id="b8f55-429">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-429">
         - Taskpane</span></span><br><span data-ttu-id="b8f55-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-430">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-431">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-432">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8f55-432">-ActiveView</span></span><br><span data-ttu-id="b8f55-433">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-433">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-434">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-434">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-435">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-435">
         - File</span></span><br><span data-ttu-id="b8f55-436">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-436">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-437">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-437">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-438">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-438">
         - Selection</span></span><br><span data-ttu-id="b8f55-439">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-439">
         - Settings</span></span><br><span data-ttu-id="b8f55-440">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-440">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-441">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-441">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-442">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-442">- Content</span></span><br><span data-ttu-id="b8f55-443">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-443">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="b8f55-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="b8f55-444">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="b8f55-445">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8f55-445">-ActiveView</span></span><br><span data-ttu-id="b8f55-446">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-446">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-447">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-447">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-448">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-448">
         - File</span></span><br><span data-ttu-id="b8f55-449">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-449">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-450">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-450">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-451">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-451">
         - Selection</span></span><br><span data-ttu-id="b8f55-452">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-452">
         - Settings</span></span><br><span data-ttu-id="b8f55-453">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-453">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-454">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="b8f55-454">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="b8f55-455">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-455">- Content</span></span><br><span data-ttu-id="b8f55-456">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-456">
         - Taskpane</span></span><br><span data-ttu-id="b8f55-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-457">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-458">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-459">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8f55-459">-ActiveView</span></span><br><span data-ttu-id="b8f55-460">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-460">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-461">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-461">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-462">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-462">
         - File</span></span><br><span data-ttu-id="b8f55-463">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-463">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-464">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-465">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-465">
         - Selection</span></span><br><span data-ttu-id="b8f55-466">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-466">
         - Settings</span></span><br><span data-ttu-id="b8f55-467">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-467">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-468">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="b8f55-468">Office for iOS</span></span></td>
    <td> <span data-ttu-id="b8f55-469">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-469">- Content</span></span><br><span data-ttu-id="b8f55-470">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-470">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="b8f55-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-471">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="b8f55-472">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8f55-472">-ActiveView</span></span><br><span data-ttu-id="b8f55-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-473">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-474">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-474">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-475">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-475">
         - File</span></span><br><span data-ttu-id="b8f55-476">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-476">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-477">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-477">
         - Selection</span></span><br><span data-ttu-id="b8f55-478">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-478">
         - Settings</span></span><br><span data-ttu-id="b8f55-479">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-479">
         -TextCoercion</span></span><br><span data-ttu-id="b8f55-480">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-480">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-481">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="b8f55-481">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="b8f55-482">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-482">- Content</span></span><br><span data-ttu-id="b8f55-483">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-483">
         - Taskpane</span></span><br><span data-ttu-id="b8f55-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-484">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-485">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-486">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8f55-486">-ActiveView</span></span><br><span data-ttu-id="b8f55-487">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-487">
         -CompressedFile</span></span><br><span data-ttu-id="b8f55-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-488">
         -DocumentEvents</span></span><br><span data-ttu-id="b8f55-489">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8f55-489">
         - File</span></span><br><span data-ttu-id="b8f55-490">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-490">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-491">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8f55-491">
         -PdfFile</span></span><br><span data-ttu-id="b8f55-492">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8f55-492">
         - Selection</span></span><br><span data-ttu-id="b8f55-493">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-493">
         - Settings</span></span><br><span data-ttu-id="b8f55-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-494">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="b8f55-495">OneNote</span><span class="sxs-lookup"><span data-stu-id="b8f55-495">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8f55-496">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8f55-496">Platform</span></span></th>
    <th><span data-ttu-id="b8f55-497">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8f55-497">Extension points</span></span></th>
    <th><span data-ttu-id="b8f55-498">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8f55-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8f55-499"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="b8f55-500">Office Online</span><span class="sxs-lookup"><span data-stu-id="b8f55-500">Office Online</span></span></td>
    <td> <span data-ttu-id="b8f55-501">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8f55-501">- Content</span></span><br><span data-ttu-id="b8f55-502">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="b8f55-502">
         - Taskpane</span></span><br><span data-ttu-id="b8f55-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-503">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8f55-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-504">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b8f55-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8f55-505">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8f55-506">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8f55-506">-DocumentEvents</span></span><br><span data-ttu-id="b8f55-507">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-507">
         -HtmlCoercion</span></span><br><span data-ttu-id="b8f55-508">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-508">
         -ImageCoercion</span></span><br><span data-ttu-id="b8f55-509">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8f55-509">
         - Settings</span></span><br><span data-ttu-id="b8f55-510">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8f55-510">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b8f55-511">Veja também</span><span class="sxs-lookup"><span data-stu-id="b8f55-511">See also</span></span>

- [<span data-ttu-id="b8f55-512">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8f55-512">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b8f55-513">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="b8f55-513">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="b8f55-514">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="b8f55-514">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="b8f55-515">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="b8f55-515">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
