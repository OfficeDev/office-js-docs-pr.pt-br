---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 09/24/2018
ms.openlocfilehash: b06602e35ec906866ad16d667036a4cbaff2d89e
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985820"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4ffa6-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4ffa6-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4ffa6-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="4ffa6-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="4ffa6-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="4ffa6-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="4ffa6-106">Se uma célula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="4ffa6-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="4ffa6-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="4ffa6-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="4ffa6-p103">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="4ffa6-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="4ffa6-110">Excel</span><span class="sxs-lookup"><span data-stu-id="4ffa6-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4ffa6-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4ffa6-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4ffa6-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4ffa6-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4ffa6-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4ffa6-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="4ffa6-115">Office Online</span></span></td>
    <td> <span data-ttu-id="4ffa6-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-116">- Taskpane</span></span><br><span data-ttu-id="4ffa6-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-117">
        - Content</span></span><br><span data-ttu-id="4ffa6-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="4ffa6-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4ffa6-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-127">
        -BindingEvents</span></span><br><span data-ttu-id="4ffa6-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-128">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-129">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-130">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-130">
        - File</span></span><br><span data-ttu-id="4ffa6-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-131">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-133">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-133">
        - Selection</span></span><br><span data-ttu-id="4ffa6-134">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-134">
        - Settings</span></span><br><span data-ttu-id="4ffa6-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-135">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-136">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-137">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-139">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4ffa6-140">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-140">
        - Taskpane</span></span><br><span data-ttu-id="4ffa6-141">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4ffa6-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-143">
        -BindingEvents</span></span><br><span data-ttu-id="4ffa6-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-144">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-145">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-146">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-146">
        - File</span></span><br><span data-ttu-id="4ffa6-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-147">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-148">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-150">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-150">
        - Selection</span></span><br><span data-ttu-id="4ffa6-151">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-151">
        - Settings</span></span><br><span data-ttu-id="4ffa6-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-152">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-153">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-154">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-156">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4ffa6-157">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-157">- Taskpane</span></span><br><span data-ttu-id="4ffa6-158">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-158">
        - Content</span></span><br><span data-ttu-id="4ffa6-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ffa6-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-168">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-169">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-170">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-171">
        - File</span></span><br><span data-ttu-id="4ffa6-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-172">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-173">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-175">
        - Selection</span></span><br><span data-ttu-id="4ffa6-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-176">
        - Settings</span></span><br><span data-ttu-id="4ffa6-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-177">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-178">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-179">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-181">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-181">Office for Windows</span></span></td>
    <td><span data-ttu-id="4ffa6-182">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-182">- Taskpane</span></span><br><span data-ttu-id="4ffa6-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-183">
        - Content</span></span><br><span data-ttu-id="4ffa6-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-184">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ffa6-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-185">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-191">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-192">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-193">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-193">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-194">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-194">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-195">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-195">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-196">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-196">
        - File</span></span><br><span data-ttu-id="4ffa6-197">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-197">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-198">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-198">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-199">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-199">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-200">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-200">
        - Selection</span></span><br><span data-ttu-id="4ffa6-201">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-201">
        - Settings</span></span><br><span data-ttu-id="4ffa6-202">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-202">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-203">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-203">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-204">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-204">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-205">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-205">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-206">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="4ffa6-206">Office for iOS</span></span></td>
    <td><span data-ttu-id="4ffa6-207">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-207">- Taskpane</span></span><br><span data-ttu-id="4ffa6-208">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-208">
        - Content</span></span></td>
    <td><span data-ttu-id="4ffa6-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-217">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-218">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-219">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-220">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-220">
        - File</span></span><br><span data-ttu-id="4ffa6-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-221">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-222">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-224">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-224">
        - Selection</span></span><br><span data-ttu-id="4ffa6-225">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-225">
        - Settings</span></span><br><span data-ttu-id="4ffa6-226">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-226">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-227">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-227">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-228">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-228">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-229">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-229">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-230">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-230">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4ffa6-231">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-231">- Taskpane</span></span><br><span data-ttu-id="4ffa6-232">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-232">
        - Content</span></span><br><span data-ttu-id="4ffa6-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-233">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ffa6-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-234">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-235">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-236">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-237">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-238">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-239">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-240">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-240">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-241">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-242">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-242">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-243">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-243">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-244">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-244">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-245">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-245">
        - File</span></span><br><span data-ttu-id="4ffa6-246">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-246">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-247">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-247">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-248">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-248">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-249">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-249">
        -PdfFile</span></span><br><span data-ttu-id="4ffa6-250">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-250">
        - Selection</span></span><br><span data-ttu-id="4ffa6-251">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-251">
        - Settings</span></span><br><span data-ttu-id="4ffa6-252">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-252">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-253">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-253">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-254">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-254">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-255">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-255">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-256">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-256">Office for Mac</span></span></td>
    <td><span data-ttu-id="4ffa6-257">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-257">- Taskpane</span></span><br><span data-ttu-id="4ffa6-258">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-258">
        - Content</span></span><br><span data-ttu-id="4ffa6-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-259">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4ffa6-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-260">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-261">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-262">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-263">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4ffa6-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-264">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4ffa6-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-265">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4ffa6-266">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-266">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="4ffa6-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-267">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4ffa6-268">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-268">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-269">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-269">
        -CompressedFile</span></span><br><span data-ttu-id="4ffa6-270">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-270">
        -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-271">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-271">
        - File</span></span><br><span data-ttu-id="4ffa6-272">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-272">
        -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-273">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-273">
        -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-274">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-274">
        -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-275">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-275">
        -PdfFile</span></span><br><span data-ttu-id="4ffa6-276">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-276">
        - Selection</span></span><br><span data-ttu-id="4ffa6-277">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-277">
        - Settings</span></span><br><span data-ttu-id="4ffa6-278">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-278">
        -TableBindings</span></span><br><span data-ttu-id="4ffa6-279">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-279">
        -TableCoercion</span></span><br><span data-ttu-id="4ffa6-280">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-280">
        -TextBindings</span></span><br><span data-ttu-id="4ffa6-281">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-281">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="4ffa6-282">Outlook</span><span class="sxs-lookup"><span data-stu-id="4ffa6-282">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ffa6-283">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4ffa6-283">Platform</span></span></th>
    <th><span data-ttu-id="4ffa6-284">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4ffa6-284">Extension points</span></span></th>
    <th><span data-ttu-id="4ffa6-285">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-285">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ffa6-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-286"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-287">Office Online</span><span class="sxs-lookup"><span data-stu-id="4ffa6-287">Office Online</span></span></td>
    <td> <span data-ttu-id="4ffa6-288">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-288">- Mail Read</span></span><br><span data-ttu-id="4ffa6-289">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-289">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ffa6-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-296">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ffa6-297">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-297">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-298">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-298">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-299">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-299">- Mail Read</span></span><br><span data-ttu-id="4ffa6-300">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-300">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-301">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-302">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-303">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-304">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-305">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4ffa6-306">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-306">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-307">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-307">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-308">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-308">- Mail Read</span></span><br><span data-ttu-id="4ffa6-309">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-309">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-310">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4ffa6-311">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="4ffa6-311">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4ffa6-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-312">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-313">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-314">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-315">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-316">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ffa6-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-317">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4ffa6-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-318">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4ffa6-319">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-319">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-320">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-320">Office for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-321">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-321">- Mail Read</span></span><br><span data-ttu-id="4ffa6-322">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-322">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-323">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4ffa6-324">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="4ffa6-324">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4ffa6-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-325">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-326">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-327">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-328">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-329">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ffa6-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-330">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ffa6-331">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-331">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-332">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="4ffa6-332">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4ffa6-333">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-333">- Mail Read</span></span><br><span data-ttu-id="4ffa6-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-334">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-335">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-336">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-337">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-338">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-339">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4ffa6-340">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-341">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-341">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-342">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-342">- Mail Read</span></span><br><span data-ttu-id="4ffa6-343">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-343">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-344">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-345">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-346">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-347">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-348">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-349">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ffa6-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-350">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ffa6-351">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-351">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-352">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-352">Office for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-353">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-353">- Mail Read</span></span><br><span data-ttu-id="4ffa6-354">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-354">
      - Mail Compose</span></span><br><span data-ttu-id="4ffa6-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-355">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-356">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-357">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-358">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-359">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-360">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4ffa6-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-361">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4ffa6-362">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-362">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-363">Office para Android</span><span class="sxs-lookup"><span data-stu-id="4ffa6-363">Office for Android</span></span></td>
    <td> <span data-ttu-id="4ffa6-364">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="4ffa6-364">- Mail Read</span></span><br><span data-ttu-id="4ffa6-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-365">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-366">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4ffa6-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-367">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4ffa6-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-368">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4ffa6-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-369">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4ffa6-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-370">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4ffa6-371">Não disponível</span><span class="sxs-lookup"><span data-stu-id="4ffa6-371">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4ffa6-372">Word</span><span class="sxs-lookup"><span data-stu-id="4ffa6-372">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ffa6-373">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4ffa6-373">Platform</span></span></th>
    <th><span data-ttu-id="4ffa6-374">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4ffa6-374">Extension points</span></span></th>
    <th><span data-ttu-id="4ffa6-375">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-375">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ffa6-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-376"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-377">Office Online</span><span class="sxs-lookup"><span data-stu-id="4ffa6-377">Office Online</span></span></td>
    <td> <span data-ttu-id="4ffa6-378">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-378">- Taskpane</span></span><br><span data-ttu-id="4ffa6-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-379">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-380">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-381">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-382">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-383">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-384">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-384">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-385">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-385">customXmlParts</span></span><br><span data-ttu-id="4ffa6-386">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-386">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-387">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-387">
         - File</span></span><br><span data-ttu-id="4ffa6-388">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-388">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-389">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-389">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-390">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-390">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-391">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-391">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-392">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-392">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-393">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-393">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-394">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-394">
         - Selection</span></span><br><span data-ttu-id="4ffa6-395">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-395">
         - Settings</span></span><br><span data-ttu-id="4ffa6-396">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-396">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-397">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-397">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-398">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-398">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-399">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-399">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-400">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-400">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-401">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-401">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-402">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-402">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4ffa6-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-403">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-404">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-404">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-405">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-405">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-406">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-406">customXmlParts</span></span><br><span data-ttu-id="4ffa6-407">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-407">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-408">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-408">
         - File</span></span><br><span data-ttu-id="4ffa6-409">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-409">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-410">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-410">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-411">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-411">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-412">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-412">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-413">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-413">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-414">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-414">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-415">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-415">
         - Selection</span></span><br><span data-ttu-id="4ffa6-416">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-416">
         - Settings</span></span><br><span data-ttu-id="4ffa6-417">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-417">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-418">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-418">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-419">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-419">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-420">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-420">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-421">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-421">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-422">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-422">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-423">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-423">- Taskpane</span></span><br><span data-ttu-id="4ffa6-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-424">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-425">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-426">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-427">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-428">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-429">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-429">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-430">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-430">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-431">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-431">customXmlParts</span></span><br><span data-ttu-id="4ffa6-432">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-432">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-433">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-433">
         - File</span></span><br><span data-ttu-id="4ffa6-434">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-434">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-435">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-436">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-436">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-437">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-437">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-438">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-438">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-439">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-439">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-440">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-440">
         - Selection</span></span><br><span data-ttu-id="4ffa6-441">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-441">
         - Settings</span></span><br><span data-ttu-id="4ffa6-442">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-442">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-443">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-443">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-444">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-444">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-445">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-445">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-446">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-446">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-447">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-447">Office for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-448">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-448">- Taskpane</span></span><br><span data-ttu-id="4ffa6-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-449">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-450">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-451">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-452">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-453">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-454">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-454">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-455">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-455">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-456">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-456">customXmlParts</span></span><br><span data-ttu-id="4ffa6-457">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-457">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-458">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-458">
         - File</span></span><br><span data-ttu-id="4ffa6-459">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-459">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-460">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-460">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-461">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-461">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-462">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-462">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-463">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-463">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-464">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-464">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-465">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-465">
         - Selection</span></span><br><span data-ttu-id="4ffa6-466">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-466">
         - Settings</span></span><br><span data-ttu-id="4ffa6-467">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-467">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-468">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-468">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-469">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-469">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-470">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-470">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-471">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-471">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-472">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="4ffa6-472">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4ffa6-473">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-473">- Taskpane</span></span></td>
    <td> <span data-ttu-id="4ffa6-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-474">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-475">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-476">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4ffa6-477">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4ffa6-478">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-478">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-479">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-479">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-480">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-480">customXmlParts</span></span><br><span data-ttu-id="4ffa6-481">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-481">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-482">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-482">
         - File</span></span><br><span data-ttu-id="4ffa6-483">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-483">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-484">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-484">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-485">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-485">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-486">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-486">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-487">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-487">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-488">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-488">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-489">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-489">
         - Selection</span></span><br><span data-ttu-id="4ffa6-490">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-490">
         - Settings</span></span><br><span data-ttu-id="4ffa6-491">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-491">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-492">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-492">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-493">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-493">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-494">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-494">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-495">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-495">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-496">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-496">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-497">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-497">- Taskpane</span></span><br><span data-ttu-id="4ffa6-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-498">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-499">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-500">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-501">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4ffa6-502">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4ffa6-503">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-503">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-504">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-504">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-505">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-505">customXmlParts</span></span><br><span data-ttu-id="4ffa6-506">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-506">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-507">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-507">
         - File</span></span><br><span data-ttu-id="4ffa6-508">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-508">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-509">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-509">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-510">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-510">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-511">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-511">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-512">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-512">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-513">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-513">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-514">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-514">
         - Selection</span></span><br><span data-ttu-id="4ffa6-515">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-515">
         - Settings</span></span><br><span data-ttu-id="4ffa6-516">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-516">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-517">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-517">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-518">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-518">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-519">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-519">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-520">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-520">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-521">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-521">Office for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-522">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-522">- Taskpane</span></span><br><span data-ttu-id="4ffa6-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-523">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-524">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-525">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4ffa6-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-526">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4ffa6-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4ffa6-527">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4ffa6-528">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-528">-BindingEvents</span></span><br><span data-ttu-id="4ffa6-529">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-529">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-530">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4ffa6-530">customXmlParts</span></span><br><span data-ttu-id="4ffa6-531">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-531">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-532">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-532">
         - File</span></span><br><span data-ttu-id="4ffa6-533">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-533">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-534">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-534">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-535">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-535">
         -MatrixBindings</span></span><br><span data-ttu-id="4ffa6-536">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-536">
         -MatrixCoercion</span></span><br><span data-ttu-id="4ffa6-537">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-537">
         -OoxmlCoercion</span></span><br><span data-ttu-id="4ffa6-538">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-538">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-539">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-539">
         - Selection</span></span><br><span data-ttu-id="4ffa6-540">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-540">
         - Settings</span></span><br><span data-ttu-id="4ffa6-541">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-541">
         -TableBindings</span></span><br><span data-ttu-id="4ffa6-542">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-542">
         -TableCoercion</span></span><br><span data-ttu-id="4ffa6-543">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4ffa6-543">
         -TextBindings</span></span><br><span data-ttu-id="4ffa6-544">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-544">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-545">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-545">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4ffa6-546">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4ffa6-546">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ffa6-547">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4ffa6-547">Platform</span></span></th>
    <th><span data-ttu-id="4ffa6-548">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4ffa6-548">Extension points</span></span></th>
    <th><span data-ttu-id="4ffa6-549">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-549">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ffa6-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-550"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-551">Office Online</span><span class="sxs-lookup"><span data-stu-id="4ffa6-551">Office Online</span></span></td>
    <td> <span data-ttu-id="4ffa6-552">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-552">- Content</span></span><br><span data-ttu-id="4ffa6-553">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-553">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-554">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-555">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-556">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-556">-ActiveView</span></span><br><span data-ttu-id="4ffa6-557">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-557">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-558">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-558">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-559">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-559">
         - File</span></span><br><span data-ttu-id="4ffa6-560">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-560">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-561">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-561">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-562">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-562">
         - Selection</span></span><br><span data-ttu-id="4ffa6-563">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-563">
         - Settings</span></span><br><span data-ttu-id="4ffa6-564">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-564">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-565">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-565">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-566">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-566">- Content</span></span><br><span data-ttu-id="4ffa6-567">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-567">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="4ffa6-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4ffa6-568">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4ffa6-569">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-569">-ActiveView</span></span><br><span data-ttu-id="4ffa6-570">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-570">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-571">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-572">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-572">
         - File</span></span><br><span data-ttu-id="4ffa6-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-573">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-574">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-575">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-575">
         - Selection</span></span><br><span data-ttu-id="4ffa6-576">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-576">
         - Settings</span></span><br><span data-ttu-id="4ffa6-577">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-577">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-578">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-578">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-579">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-579">- Content</span></span><br><span data-ttu-id="4ffa6-580">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-580">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-581">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-582">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-583">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-583">-ActiveView</span></span><br><span data-ttu-id="4ffa6-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-584">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-585">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-586">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-586">
         - File</span></span><br><span data-ttu-id="4ffa6-587">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-587">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-588">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-588">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-589">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-589">
         - Selection</span></span><br><span data-ttu-id="4ffa6-590">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-590">
         - Settings</span></span><br><span data-ttu-id="4ffa6-591">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-591">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-592">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="4ffa6-592">Office for Windows</span></span></td>
    <td> <span data-ttu-id="4ffa6-593">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-593">- Content</span></span><br><span data-ttu-id="4ffa6-594">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-594">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-595">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-596">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-597">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-597">-ActiveView</span></span><br><span data-ttu-id="4ffa6-598">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-598">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-599">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-599">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-600">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-600">
         - File</span></span><br><span data-ttu-id="4ffa6-601">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-601">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-602">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-602">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-603">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-603">
         - Selection</span></span><br><span data-ttu-id="4ffa6-604">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-604">
         - Settings</span></span><br><span data-ttu-id="4ffa6-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-605">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-606">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="4ffa6-606">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4ffa6-607">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-607">- Content</span></span><br><span data-ttu-id="4ffa6-608">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-608">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="4ffa6-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-609">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4ffa6-610">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-610">-ActiveView</span></span><br><span data-ttu-id="4ffa6-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-611">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-612">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-613">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-613">
         - File</span></span><br><span data-ttu-id="4ffa6-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-614">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-615">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-615">
         - Selection</span></span><br><span data-ttu-id="4ffa6-616">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-616">
         - Settings</span></span><br><span data-ttu-id="4ffa6-617">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-617">
         -TextCoercion</span></span><br><span data-ttu-id="4ffa6-618">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-618">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-619">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-619">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-620">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-620">- Content</span></span><br><span data-ttu-id="4ffa6-621">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-621">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-622">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-623">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-624">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-624">-ActiveView</span></span><br><span data-ttu-id="4ffa6-625">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-625">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-626">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-627">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-627">
         - File</span></span><br><span data-ttu-id="4ffa6-628">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-628">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-629">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-629">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-630">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-630">
         - Selection</span></span><br><span data-ttu-id="4ffa6-631">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-631">
         - Settings</span></span><br><span data-ttu-id="4ffa6-632">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-632">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-633">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="4ffa6-633">Office for Mac</span></span></td>
    <td> <span data-ttu-id="4ffa6-634">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-634">- Content</span></span><br><span data-ttu-id="4ffa6-635">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-635">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-636">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-637">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-638">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4ffa6-638">-ActiveView</span></span><br><span data-ttu-id="4ffa6-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-639">
         -CompressedFile</span></span><br><span data-ttu-id="4ffa6-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-640">
         -DocumentEvents</span></span><br><span data-ttu-id="4ffa6-641">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-641">
         - File</span></span><br><span data-ttu-id="4ffa6-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-642">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-643">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4ffa6-643">
         -PdfFile</span></span><br><span data-ttu-id="4ffa6-644">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="4ffa6-644">
         - Selection</span></span><br><span data-ttu-id="4ffa6-645">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-645">
         - Settings</span></span><br><span data-ttu-id="4ffa6-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-646">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="4ffa6-647">OneNote</span><span class="sxs-lookup"><span data-stu-id="4ffa6-647">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4ffa6-648">Plataforma</span><span class="sxs-lookup"><span data-stu-id="4ffa6-648">Platform</span></span></th>
    <th><span data-ttu-id="4ffa6-649">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="4ffa6-649">Extension points</span></span></th>
    <th><span data-ttu-id="4ffa6-650">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-650">API requirement sets</span></span></th>
    <th><span data-ttu-id="4ffa6-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-651"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="4ffa6-652">Office Online</span><span class="sxs-lookup"><span data-stu-id="4ffa6-652">Office Online</span></span></td>
    <td> <span data-ttu-id="4ffa6-653">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="4ffa6-653">- Content</span></span><br><span data-ttu-id="4ffa6-654">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="4ffa6-654">
         - Taskpane</span></span><br><span data-ttu-id="4ffa6-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-655">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-656">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4ffa6-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4ffa6-657">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4ffa6-658">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4ffa6-658">-DocumentEvents</span></span><br><span data-ttu-id="4ffa6-659">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-659">
         -HtmlCoercion</span></span><br><span data-ttu-id="4ffa6-660">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-660">
         -ImageCoercion</span></span><br><span data-ttu-id="4ffa6-661">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="4ffa6-661">
         - Settings</span></span><br><span data-ttu-id="4ffa6-662">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4ffa6-662">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4ffa6-663">Veja também</span><span class="sxs-lookup"><span data-stu-id="4ffa6-663">See also</span></span>

- [<span data-ttu-id="4ffa6-664">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="4ffa6-664">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4ffa6-665">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="4ffa6-665">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4ffa6-666">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="4ffa6-666">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4ffa6-667">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="4ffa6-667">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
