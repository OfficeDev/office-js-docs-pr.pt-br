---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.
ms.date: 08/30/2018
ms.openlocfilehash: 06fb073693bd8adca7d196f4361699ac3f54cee1
ms.sourcegitcommit: 78b28ae88d53bfef3134c09cc4336a5a8722c70b
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/01/2018
ms.locfileid: "23797298"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e0608-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e0608-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e0608-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="e0608-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="e0608-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="e0608-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

<span data-ttu-id="e0608-106">Se uma célula de tabela apresenta um asterisco (\*), significa que estamos trabalhando no assunto.</span><span class="sxs-lookup"><span data-stu-id="e0608-106">If a table cell contains an asterisk ( \* ), that means we’re working on it.</span></span> <span data-ttu-id="e0608-107">Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span><span class="sxs-lookup"><span data-stu-id="e0608-107">For requirement sets for Project or Access, see [Office common requirement sets](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets).</span></span>  

> [!NOTE]
> <span data-ttu-id="e0608-p103">O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="e0608-p103">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e0608-110">Excel</span><span class="sxs-lookup"><span data-stu-id="e0608-110">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e0608-111">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e0608-111">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e0608-112">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e0608-112">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e0608-113">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e0608-113">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e0608-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e0608-114"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-115">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0608-115">Office Online</span></span></td>
    <td> <span data-ttu-id="e0608-116">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-116">- Taskpane</span></span><br><span data-ttu-id="e0608-117">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-117">
        - Content</span></span><br><span data-ttu-id="e0608-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="e0608-118">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e0608-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-119">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0608-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-120">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0608-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-121">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0608-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-122">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0608-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-123">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0608-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-124">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0608-125">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0608-125">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0608-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-126">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0608-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-127">
        -BindingEvents</span></span><br><span data-ttu-id="e0608-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-128">
        -CompressedFile</span></span><br><span data-ttu-id="e0608-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-129">
        -DocumentEvents</span></span><br><span data-ttu-id="e0608-130">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-130">
        - File</span></span><br><span data-ttu-id="e0608-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-131">
        -MatrixBindings</span></span><br><span data-ttu-id="e0608-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-132">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0608-133">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-133">
        - Selection</span></span><br><span data-ttu-id="e0608-134">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-134">
        - Settings</span></span><br><span data-ttu-id="e0608-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-135">
        -TableBindings</span></span><br><span data-ttu-id="e0608-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-136">
        -TableCoercion</span></span><br><span data-ttu-id="e0608-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-137">
        -TextBindings</span></span><br><span data-ttu-id="e0608-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-138">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-139">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-139">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e0608-140">
        - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-140">
        - Taskpane</span></span><br><span data-ttu-id="e0608-141">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-141">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e0608-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-142">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0608-143">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-143">
        -BindingEvents</span></span><br><span data-ttu-id="e0608-144">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-144">
        -CompressedFile</span></span><br><span data-ttu-id="e0608-145">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-145">
        -DocumentEvents</span></span><br><span data-ttu-id="e0608-146">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-146">
        - File</span></span><br><span data-ttu-id="e0608-147">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-147">
        -ImageCoercion</span></span><br><span data-ttu-id="e0608-148">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-148">
        -MatrixBindings</span></span><br><span data-ttu-id="e0608-149">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-149">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0608-150">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-150">
        - Selection</span></span><br><span data-ttu-id="e0608-151">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-151">
        - Settings</span></span><br><span data-ttu-id="e0608-152">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-152">
        -TableBindings</span></span><br><span data-ttu-id="e0608-153">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-153">
        -TableCoercion</span></span><br><span data-ttu-id="e0608-154">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-154">
        -TextBindings</span></span><br><span data-ttu-id="e0608-155">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-155">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-156">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-156">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e0608-157">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-157">- Taskpane</span></span><br><span data-ttu-id="e0608-158">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-158">
        - Content</span></span><br><span data-ttu-id="e0608-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-159">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e0608-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-160">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0608-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-161">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0608-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-162">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0608-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-163">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0608-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-164">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0608-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-165">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0608-166">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0608-166">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0608-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-167">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0608-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-168">-BindingEvents</span></span><br><span data-ttu-id="e0608-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-169">
        -CompressedFile</span></span><br><span data-ttu-id="e0608-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-170">
        -DocumentEvents</span></span><br><span data-ttu-id="e0608-171">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-171">
        - File</span></span><br><span data-ttu-id="e0608-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-172">
        -ImageCoercion</span></span><br><span data-ttu-id="e0608-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-173">
        -MatrixBindings</span></span><br><span data-ttu-id="e0608-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-174">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0608-175">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-175">
        - Selection</span></span><br><span data-ttu-id="e0608-176">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-176">
        - Settings</span></span><br><span data-ttu-id="e0608-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-177">
        -TableBindings</span></span><br><span data-ttu-id="e0608-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-178">
        -TableCoercion</span></span><br><span data-ttu-id="e0608-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-179">
        -TextBindings</span></span><br><span data-ttu-id="e0608-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-180">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-181">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e0608-181">Office for iOS</span></span></td>
    <td><span data-ttu-id="e0608-182">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-182">- Taskpane</span></span><br><span data-ttu-id="e0608-183">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-183">
        - Content</span></span></td>
    <td><span data-ttu-id="e0608-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-184">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0608-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-185">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0608-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-186">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0608-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-187">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0608-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-188">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0608-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-189">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0608-190">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0608-190">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0608-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-191">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0608-192">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-192">-BindingEvents</span></span><br><span data-ttu-id="e0608-193">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-193">
        -CompressedFile</span></span><br><span data-ttu-id="e0608-194">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-194">
        -DocumentEvents</span></span><br><span data-ttu-id="e0608-195">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-195">
        - File</span></span><br><span data-ttu-id="e0608-196">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-196">
        -ImageCoercion</span></span><br><span data-ttu-id="e0608-197">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-197">
        -MatrixBindings</span></span><br><span data-ttu-id="e0608-198">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-198">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0608-199">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-199">
        - Selection</span></span><br><span data-ttu-id="e0608-200">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-200">
        - Settings</span></span><br><span data-ttu-id="e0608-201">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-201">
        -TableBindings</span></span><br><span data-ttu-id="e0608-202">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-202">
        -TableCoercion</span></span><br><span data-ttu-id="e0608-203">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-203">
        -TextBindings</span></span><br><span data-ttu-id="e0608-204">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-204">
        -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-205">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e0608-205">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e0608-206">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-206">- Taskpane</span></span><br><span data-ttu-id="e0608-207">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-207">
        - Content</span></span><br><span data-ttu-id="e0608-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-208">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e0608-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-209">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e0608-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-210">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e0608-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-211">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e0608-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-212">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e0608-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-213">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e0608-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-214">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e0608-215">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e0608-215">ExcelApi 1.7 Beta</span></span><br><span data-ttu-id="e0608-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-216">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e0608-217">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-217">-BindingEvents</span></span><br><span data-ttu-id="e0608-218">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-218">
        -CompressedFile</span></span><br><span data-ttu-id="e0608-219">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-219">
        -DocumentEvents</span></span><br><span data-ttu-id="e0608-220">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-220">
        - File</span></span><br><span data-ttu-id="e0608-221">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-221">
        -ImageCoercion</span></span><br><span data-ttu-id="e0608-222">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-222">
        -MatrixBindings</span></span><br><span data-ttu-id="e0608-223">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-223">
        -MatrixCoercion</span></span><br><span data-ttu-id="e0608-224">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-224">
        -PdfFile</span></span><br><span data-ttu-id="e0608-225">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-225">
        - Selection</span></span><br><span data-ttu-id="e0608-226">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-226">
        - Settings</span></span><br><span data-ttu-id="e0608-227">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-227">
        -TableBindings</span></span><br><span data-ttu-id="e0608-228">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-228">
        -TableCoercion</span></span><br><span data-ttu-id="e0608-229">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-229">
        -TextBindings</span></span><br><span data-ttu-id="e0608-230">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-230">
        -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e0608-231">Outlook</span><span class="sxs-lookup"><span data-stu-id="e0608-231">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0608-232">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e0608-232">Platform</span></span></th>
    <th><span data-ttu-id="e0608-233">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e0608-233">Extension points</span></span></th>
    <th><span data-ttu-id="e0608-234">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e0608-234">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0608-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e0608-235"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-236">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0608-236">Office Online</span></span></td>
    <td> <span data-ttu-id="e0608-237">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-237">- Mail Read</span></span><br><span data-ttu-id="e0608-238">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e0608-238">
      - Mail Compose</span></span><br><span data-ttu-id="e0608-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-239">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-240">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-241">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-242">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-243">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0608-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-244">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0608-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-245">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0608-246">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-246">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-247">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-247">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-248">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-248">- Mail Read</span></span><br><span data-ttu-id="e0608-249">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e0608-249">
      - Mail Compose</span></span><br><span data-ttu-id="e0608-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-250">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-251">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-252">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-253">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-254">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e0608-255">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-255">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-256">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-256">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-257">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-257">- Mail Read</span></span><br><span data-ttu-id="e0608-258">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e0608-258">
      - Mail Compose</span></span><br><span data-ttu-id="e0608-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-259">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e0608-260">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="e0608-260">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e0608-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-261">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-262">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-263">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-264">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0608-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-265">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0608-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-266">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0608-267">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-267">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-268">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e0608-268">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0608-269">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-269">- Mail Read</span></span><br><span data-ttu-id="e0608-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-270">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-271">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-272">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-273">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-274">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0608-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-275">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e0608-276">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-276">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-277">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e0608-277">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0608-278">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-278">- Mail Read</span></span><br><span data-ttu-id="e0608-279">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="e0608-279">
      - Mail Compose</span></span><br><span data-ttu-id="e0608-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-280">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-281">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-282">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-283">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-284">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0608-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-285">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e0608-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e0608-286">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e0608-287">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-287">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-288">Office para Android</span><span class="sxs-lookup"><span data-stu-id="e0608-288">Office for Android</span></span></td>
    <td> <span data-ttu-id="e0608-289">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="e0608-289">- Mail Read</span></span><br><span data-ttu-id="e0608-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-290">
      - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-291">- <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e0608-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-292">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e0608-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-293">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e0608-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e0608-294">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e0608-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e0608-295">
      - <a href="https://docs.microsoft.com/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e0608-296">Não disponível</span><span class="sxs-lookup"><span data-stu-id="e0608-296">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e0608-297">Word</span><span class="sxs-lookup"><span data-stu-id="e0608-297">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0608-298">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e0608-298">Platform</span></span></th>
    <th><span data-ttu-id="e0608-299">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e0608-299">Extension points</span></span></th>
    <th><span data-ttu-id="e0608-300">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e0608-300">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0608-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e0608-301"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-302">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0608-302">Office Online</span></span></td>
    <td> <span data-ttu-id="e0608-303">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-303">- Taskpane</span></span><br><span data-ttu-id="e0608-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-304">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-305">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0608-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-306">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0608-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-307">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0608-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-308">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-309">-BindingEvents</span></span><br><span data-ttu-id="e0608-310">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e0608-310">customXmlParts</span></span><br><span data-ttu-id="e0608-311">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-311">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-312">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-312">
         - File</span></span><br><span data-ttu-id="e0608-313">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-313">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-314">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-314">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-315">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-315">
         -MatrixBindings</span></span><br><span data-ttu-id="e0608-316">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-316">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0608-317">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-317">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0608-318">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-318">
         -PdfFile</span></span><br><span data-ttu-id="e0608-319">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-319">
         - Selection</span></span><br><span data-ttu-id="e0608-320">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-320">
         - Settings</span></span><br><span data-ttu-id="e0608-321">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-321">
         -TableBindings</span></span><br><span data-ttu-id="e0608-322">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-322">
         -TableCoercion</span></span><br><span data-ttu-id="e0608-323">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-323">
         -TextBindings</span></span><br><span data-ttu-id="e0608-324">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-324">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-325">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0608-325">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-326">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-326">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-327">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-327">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e0608-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-328">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-329">-BindingEvents</span></span><br><span data-ttu-id="e0608-330">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-330">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-331">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e0608-331">customXmlParts</span></span><br><span data-ttu-id="e0608-332">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-332">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-333">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-333">
         - File</span></span><br><span data-ttu-id="e0608-334">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-334">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-335">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-335">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-336">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-336">
         -MatrixBindings</span></span><br><span data-ttu-id="e0608-337">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-337">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0608-338">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-338">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0608-339">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-339">
         -PdfFile</span></span><br><span data-ttu-id="e0608-340">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-340">
         - Selection</span></span><br><span data-ttu-id="e0608-341">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-341">
         - Settings</span></span><br><span data-ttu-id="e0608-342">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-342">
         -TableBindings</span></span><br><span data-ttu-id="e0608-343">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-343">
         -TableCoercion</span></span><br><span data-ttu-id="e0608-344">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-344">
         -TextBindings</span></span><br><span data-ttu-id="e0608-345">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-345">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-346">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0608-346">
         -TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-347">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-347">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-348">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-348">- Taskpane</span></span><br><span data-ttu-id="e0608-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-349">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-350">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0608-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-351">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0608-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-352">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0608-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-353">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-354">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-354">-BindingEvents</span></span><br><span data-ttu-id="e0608-355">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-355">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-356">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e0608-356">customXmlParts</span></span><br><span data-ttu-id="e0608-357">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-357">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-358">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-358">
         - File</span></span><br><span data-ttu-id="e0608-359">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-359">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-360">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-360">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-361">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-361">
         -MatrixBindings</span></span><br><span data-ttu-id="e0608-362">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-362">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0608-363">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-363">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0608-364">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-364">
         -PdfFile</span></span><br><span data-ttu-id="e0608-365">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-365">
         - Selection</span></span><br><span data-ttu-id="e0608-366">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-366">
         - Settings</span></span><br><span data-ttu-id="e0608-367">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-367">
         -TableBindings</span></span><br><span data-ttu-id="e0608-368">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-368">
         -TableCoercion</span></span><br><span data-ttu-id="e0608-369">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-369">
         -TextBindings</span></span><br><span data-ttu-id="e0608-370">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-370">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-371">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0608-371">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-372">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e0608-372">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0608-373">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-373">- Taskpane</span></span></td>
    <td> <span data-ttu-id="e0608-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-374">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0608-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-375">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0608-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-376">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0608-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0608-377">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0608-378">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-378">-BindingEvents</span></span><br><span data-ttu-id="e0608-379">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-379">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-380">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e0608-380">customXmlParts</span></span><br><span data-ttu-id="e0608-381">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-381">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-382">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-382">
         - File</span></span><br><span data-ttu-id="e0608-383">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-383">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-384">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-384">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-385">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-385">
         -MatrixBindings</span></span><br><span data-ttu-id="e0608-386">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-386">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0608-387">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-387">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0608-388">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-388">
         -PdfFile</span></span><br><span data-ttu-id="e0608-389">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-389">
         - Selection</span></span><br><span data-ttu-id="e0608-390">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-390">
         - Settings</span></span><br><span data-ttu-id="e0608-391">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-391">
         -TableBindings</span></span><br><span data-ttu-id="e0608-392">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-392">
         -TableCoercion</span></span><br><span data-ttu-id="e0608-393">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-393">
         -TextBindings</span></span><br><span data-ttu-id="e0608-394">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-394">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-395">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0608-395">
         -TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-396">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e0608-396">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0608-397">- Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-397">- Taskpane</span></span><br><span data-ttu-id="e0608-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-398">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-399">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e0608-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e0608-400">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e0608-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e0608-401">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e0608-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0608-402">
        - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0608-403">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-403">-BindingEvents</span></span><br><span data-ttu-id="e0608-404">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-404">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-405">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e0608-405">customXmlParts</span></span><br><span data-ttu-id="e0608-406">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-406">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-407">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-407">
         - File</span></span><br><span data-ttu-id="e0608-408">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-408">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-409">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-409">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-410">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-410">
         -MatrixBindings</span></span><br><span data-ttu-id="e0608-411">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-411">
         -MatrixCoercion</span></span><br><span data-ttu-id="e0608-412">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-412">
         -OoxmlCoercion</span></span><br><span data-ttu-id="e0608-413">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-413">
         -PdfFile</span></span><br><span data-ttu-id="e0608-414">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-414">
         - Selection</span></span><br><span data-ttu-id="e0608-415">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-415">
         - Settings</span></span><br><span data-ttu-id="e0608-416">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-416">
         -TableBindings</span></span><br><span data-ttu-id="e0608-417">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-417">
         -TableCoercion</span></span><br><span data-ttu-id="e0608-418">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e0608-418">
         -TextBindings</span></span><br><span data-ttu-id="e0608-419">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-419">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-420">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e0608-420">
         -TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e0608-421">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e0608-421">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0608-422">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e0608-422">Platform</span></span></th>
    <th><span data-ttu-id="e0608-423">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e0608-423">Extension points</span></span></th>
    <th><span data-ttu-id="e0608-424">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e0608-424">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0608-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e0608-425"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-426">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0608-426">Office Online</span></span></td>
    <td> <span data-ttu-id="e0608-427">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-427">- Content</span></span><br><span data-ttu-id="e0608-428">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-428">
         - Taskpane</span></span><br><span data-ttu-id="e0608-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-429">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-430">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-431">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0608-431">-ActiveView</span></span><br><span data-ttu-id="e0608-432">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-432">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-433">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-433">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-434">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-434">
         - File</span></span><br><span data-ttu-id="e0608-435">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-435">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-436">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-436">
         -PdfFile</span></span><br><span data-ttu-id="e0608-437">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-437">
         - Selection</span></span><br><span data-ttu-id="e0608-438">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-438">
         - Settings</span></span><br><span data-ttu-id="e0608-439">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-439">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-440">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-440">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-441">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-441">- Content</span></span><br><span data-ttu-id="e0608-442">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-442">
         - Taskpane</span></span><br>
    </td>
    <td> <span data-ttu-id="e0608-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e0608-443">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e0608-444">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0608-444">-ActiveView</span></span><br><span data-ttu-id="e0608-445">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-445">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-446">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-446">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-447">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-447">
         - File</span></span><br><span data-ttu-id="e0608-448">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-448">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-449">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-449">
         -PdfFile</span></span><br><span data-ttu-id="e0608-450">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-450">
         - Selection</span></span><br><span data-ttu-id="e0608-451">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-451">
         - Settings</span></span><br><span data-ttu-id="e0608-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-452">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-453">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="e0608-453">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e0608-454">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-454">- Content</span></span><br><span data-ttu-id="e0608-455">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-455">
         - Taskpane</span></span><br><span data-ttu-id="e0608-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-456">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-457">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-458">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0608-458">-ActiveView</span></span><br><span data-ttu-id="e0608-459">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-459">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-460">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-460">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-461">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-461">
         - File</span></span><br><span data-ttu-id="e0608-462">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-462">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-463">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-463">
         -PdfFile</span></span><br><span data-ttu-id="e0608-464">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-464">
         - Selection</span></span><br><span data-ttu-id="e0608-465">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-465">
         - Settings</span></span><br><span data-ttu-id="e0608-466">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-466">
         -TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-467">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="e0608-467">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e0608-468">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-468">- Content</span></span><br><span data-ttu-id="e0608-469">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-469">
         - Taskpane</span></span></td>
    <td> <span data-ttu-id="e0608-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-470">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e0608-471">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0608-471">-ActiveView</span></span><br><span data-ttu-id="e0608-472">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-472">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-473">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-473">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-474">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-474">
         - File</span></span><br><span data-ttu-id="e0608-475">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-475">
         -PdfFile</span></span><br><span data-ttu-id="e0608-476">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-476">
         - Selection</span></span><br><span data-ttu-id="e0608-477">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-477">
         - Settings</span></span><br><span data-ttu-id="e0608-478">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-478">
         -TextCoercion</span></span><br><span data-ttu-id="e0608-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-479">
         -ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-480">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="e0608-480">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e0608-481">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-481">- Content</span></span><br><span data-ttu-id="e0608-482">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-482">
         - Taskpane</span></span><br><span data-ttu-id="e0608-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-483">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-484">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-485">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e0608-485">-ActiveView</span></span><br><span data-ttu-id="e0608-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e0608-486">
         -CompressedFile</span></span><br><span data-ttu-id="e0608-487">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-487">
         -DocumentEvents</span></span><br><span data-ttu-id="e0608-488">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="e0608-488">
         - File</span></span><br><span data-ttu-id="e0608-489">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-489">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-490">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e0608-490">
         -PdfFile</span></span><br><span data-ttu-id="e0608-491">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="e0608-491">
         - Selection</span></span><br><span data-ttu-id="e0608-492">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-492">
         - Settings</span></span><br><span data-ttu-id="e0608-493">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-493">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e0608-494">OneNote</span><span class="sxs-lookup"><span data-stu-id="e0608-494">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e0608-495">Plataforma</span><span class="sxs-lookup"><span data-stu-id="e0608-495">Platform</span></span></th>
    <th><span data-ttu-id="e0608-496">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="e0608-496">Extension points</span></span></th>
    <th><span data-ttu-id="e0608-497">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="e0608-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="e0608-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="e0608-498"><a href="https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr> 
  </tr>
  <tr>
    <td><span data-ttu-id="e0608-499">Office Online</span><span class="sxs-lookup"><span data-stu-id="e0608-499">Office Online</span></span></td>
    <td> <span data-ttu-id="e0608-500">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="e0608-500">- Content</span></span><br><span data-ttu-id="e0608-501">
         - Painel de tarefas</span><span class="sxs-lookup"><span data-stu-id="e0608-501">
         - Taskpane</span></span><br><span data-ttu-id="e0608-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="e0608-502">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e0608-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-503">- <a href="https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e0608-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e0608-504">
         - <a href="https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e0608-505">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e0608-505">-DocumentEvents</span></span><br><span data-ttu-id="e0608-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-506">
         -HtmlCoercion</span></span><br><span data-ttu-id="e0608-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-507">
         -ImageCoercion</span></span><br><span data-ttu-id="e0608-508">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="e0608-508">
         - Settings</span></span><br><span data-ttu-id="e0608-509">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e0608-509">
         -TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e0608-510">Veja também</span><span class="sxs-lookup"><span data-stu-id="e0608-510">See also</span></span>

- [<span data-ttu-id="e0608-511">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="e0608-511">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e0608-512">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="e0608-512">Common API requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e0608-513">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="e0608-513">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/javascript/office/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e0608-514">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="e0608-514">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/javascript/office/javascript-api-for-office)
