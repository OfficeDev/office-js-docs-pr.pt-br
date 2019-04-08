---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477590"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="16513-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="16513-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="16513-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="16513-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="16513-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="16513-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="16513-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="16513-106">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="16513-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="16513-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="16513-108">Excel</span><span class="sxs-lookup"><span data-stu-id="16513-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="16513-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="16513-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="16513-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-111">API requirement sets</span></span></th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-112">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-112">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="16513-113">Office Online</span></span></td>
    <td> - <span data-ttu-id="16513-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-114">TaskPane</span></span><br>
        - <span data-ttu-id="16513-115">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-115">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-116">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-116">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-117">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-117">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-118">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-118">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-119">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-119">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-120">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-120">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-121">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-122">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-122">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-123">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-123">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-124">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-124">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-125">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-125">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="16513-126">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-126">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-127">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-127">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-128">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-128">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-129">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-129">File</span></span><br>
        - <span data-ttu-id="16513-130">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-130">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-131">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-131">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-132">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-132">Selection</span></span><br>
        - <span data-ttu-id="16513-133">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-133">Settings</span></span><br>
        - <span data-ttu-id="16513-134">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-134">TableBindings</span></span><br>
        - <span data-ttu-id="16513-135">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-135">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-136">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-136">TextBindings</span></span><br>
        - <span data-ttu-id="16513-137">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-137">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-138">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-138">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-139">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-139">TaskPane</span></span><br>
        - <span data-ttu-id="16513-140">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-140">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-141">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-141">Add-in Commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-142">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-142">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-143">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-143">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-144">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-144">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-145">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-145">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-146">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-146">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-147">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-147">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-148">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-148">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-149">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-149">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-150">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-150">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="16513-151">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-151">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-152">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-152">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-153">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-153">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-154">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-154">File</span></span><br>
        - <span data-ttu-id="16513-155">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-155">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-156">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-156">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-157">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-157">Selection</span></span><br>
        - <span data-ttu-id="16513-158">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-158">Settings</span></span><br>
        - <span data-ttu-id="16513-159">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-159">TableBindings</span></span><br>
        - <span data-ttu-id="16513-160">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-160">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-161">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-161">TextBindings</span></span><br>
        - <span data-ttu-id="16513-162">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-162">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-163">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-163">Office 2019 for Windows</span></span></td>
    <td>- <span data-ttu-id="16513-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-164">TaskPane</span></span><br>
        - <span data-ttu-id="16513-165">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-165">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-166">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-166">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-167">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-167">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-168">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-168">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-169">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-169">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-170">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-170">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-171">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-171">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-172">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-172">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-173">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-173">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-174">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-174">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-175">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-175">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="16513-176">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-176">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-177">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-177">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-178">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-178">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-179">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-179">File</span></span><br>
        - <span data-ttu-id="16513-180">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-180">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-181">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-181">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-182">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-182">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-183">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-183">Selection</span></span><br>
        - <span data-ttu-id="16513-184">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-184">Settings</span></span><br>
        - <span data-ttu-id="16513-185">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-185">TableBindings</span></span><br>
        - <span data-ttu-id="16513-186">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-186">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-187">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-187">TextBindings</span></span><br>
        - <span data-ttu-id="16513-188">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-188">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-189">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-189">Office 2016 for Windows</span></span></td>
    <td>- <span data-ttu-id="16513-190">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-190">TaskPane</span></span><br>
        - <span data-ttu-id="16513-191">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-191">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-192">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-192">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-193">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-193">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="16513-194">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-194">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-195">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-195">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-196">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-196">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-197">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-197">File</span></span><br>
        - <span data-ttu-id="16513-198">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-198">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-199">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-199">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-200">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-200">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-201">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-201">Selection</span></span><br>
        - <span data-ttu-id="16513-202">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-202">Settings</span></span><br>
        - <span data-ttu-id="16513-203">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-203">TableBindings</span></span><br>
        - <span data-ttu-id="16513-204">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-204">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-205">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-205">TextBindings</span></span><br>
        - <span data-ttu-id="16513-206">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-206">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-207">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-207">Office 2013 for Windows</span></span></td>
    <td>
        - <span data-ttu-id="16513-208">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-208">TaskPane</span></span><br>
        - <span data-ttu-id="16513-209">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-209">Content</span></span></td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-210">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-210">DialogApi 1.1</span></span></a>*</td>
    <td>
        - <span data-ttu-id="16513-211">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-211">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-212">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-212">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-213">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-213">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-214">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-214">File</span></span><br>
        - <span data-ttu-id="16513-215">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-215">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-216">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-216">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-217">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-217">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-218">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-218">Selection</span></span><br>
        - <span data-ttu-id="16513-219">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-219">Settings</span></span><br>
        - <span data-ttu-id="16513-220">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-220">TableBindings</span></span><br>
        - <span data-ttu-id="16513-221">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-221">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-222">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-222">TextBindings</span></span><br>
        - <span data-ttu-id="16513-223">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-223">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-224">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="16513-224">Office 365 for iPad</span></span></td>
    <td>- <span data-ttu-id="16513-225">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-225">TaskPane</span></span><br>
        - <span data-ttu-id="16513-226">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-226">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-227">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-227">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-228">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-228">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-229">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-229">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-230">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-230">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-231">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-231">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-232">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-232">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-233">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-233">ExcelApi 1.7</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-234">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-234">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-235">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-235">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="16513-236">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-236">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-237">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-237">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-238">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-238">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-239">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-239">File</span></span><br>
        - <span data-ttu-id="16513-240">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-240">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-241">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-241">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-242">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-242">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-243">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-243">Selection</span></span><br>
        - <span data-ttu-id="16513-244">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-244">Settings</span></span><br>
        - <span data-ttu-id="16513-245">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-245">TableBindings</span></span><br>
        - <span data-ttu-id="16513-246">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-246">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-247">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-247">TextBindings</span></span><br>
        - <span data-ttu-id="16513-248">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-248">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-249">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-249">Office 365 for Mac</span></span></td>
    <td>- <span data-ttu-id="16513-250">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-250">TaskPane</span></span><br>
        - <span data-ttu-id="16513-251">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-251">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-252">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-252">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-253">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-253">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-254">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-254">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-255">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-255">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-256">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-256">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-257">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-257">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-258">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-258">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-259">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-259">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-260">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-260">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-261">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-261">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="16513-262">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-262">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-263">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-263">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-264">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-264">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-265">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-265">File</span></span><br>
        - <span data-ttu-id="16513-266">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-266">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-267">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-267">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-268">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-268">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-269">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-269">PdfFile</span></span><br>
        - <span data-ttu-id="16513-270">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-270">Selection</span></span><br>
        - <span data-ttu-id="16513-271">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-271">Settings</span></span><br>
        - <span data-ttu-id="16513-272">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-272">TableBindings</span></span><br>
        - <span data-ttu-id="16513-273">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-273">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-274">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-274">TextBindings</span></span><br>
        - <span data-ttu-id="16513-275">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-275">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-276">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-276">Office 2019 for Mac</span></span></td>
    <td>- <span data-ttu-id="16513-277">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-277">TaskPane</span></span><br>
        - <span data-ttu-id="16513-278">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-278">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-279">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-279">Add-in Commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-280">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-280">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-281">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-281">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-282">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-282">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-283">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-283">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-284">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-284">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-285">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-285">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-286">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-286">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-287">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="16513-287">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-288">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-288">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="16513-289">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-289">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-290">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-290">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-291">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-291">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-292">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-292">File</span></span><br>
        - <span data-ttu-id="16513-293">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-293">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-294">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-294">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-295">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-295">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-296">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-296">PdfFile</span></span><br>
        - <span data-ttu-id="16513-297">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-297">Selection</span></span><br>
        - <span data-ttu-id="16513-298">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-298">Settings</span></span><br>
        - <span data-ttu-id="16513-299">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-299">TableBindings</span></span><br>
        - <span data-ttu-id="16513-300">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-300">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-301">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-301">TextBindings</span></span><br>
        - <span data-ttu-id="16513-302">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-302">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-303">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-303">Office 2016 for Mac</span></span></td>
    <td>- <span data-ttu-id="16513-304">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-304">TaskPane</span></span><br>
        - <span data-ttu-id="16513-305">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-305">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="16513-306">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-306">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-307">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-307">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="16513-308">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-308">BindingEvents</span></span><br>
        - <span data-ttu-id="16513-309">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-309">CompressedFile</span></span><br>
        - <span data-ttu-id="16513-310">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-310">DocumentEvents</span></span><br>
        - <span data-ttu-id="16513-311">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-311">File</span></span><br>
        - <span data-ttu-id="16513-312">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-312">ImageCoercion</span></span><br>
        - <span data-ttu-id="16513-313">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-313">MatrixBindings</span></span><br>
        - <span data-ttu-id="16513-314">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-314">MatrixCoercion</span></span><br>
        - <span data-ttu-id="16513-315">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-315">PdfFile</span></span><br>
        - <span data-ttu-id="16513-316">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-316">Selection</span></span><br>
        - <span data-ttu-id="16513-317">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-317">Settings</span></span><br>
        - <span data-ttu-id="16513-318">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-318">TableBindings</span></span><br>
        - <span data-ttu-id="16513-319">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-319">TableCoercion</span></span><br>
        - <span data-ttu-id="16513-320">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-320">TextBindings</span></span><br>
        - <span data-ttu-id="16513-321">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-321">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="16513-322">&ast; – Adicionado com atualizações pós-lançamento.</span><span class="sxs-lookup"><span data-stu-id="16513-322">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="outlook"></a><span data-ttu-id="16513-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="16513-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16513-324">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-324">Platform</span></span></th>
    <th><span data-ttu-id="16513-325">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-325">Extension points</span></span></th>
    <th><span data-ttu-id="16513-326">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-326">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-327">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-327">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="16513-328">Office Online</span></span></td>
    <td> - <span data-ttu-id="16513-329">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-329">Mail Read</span></span><br>
      - <span data-ttu-id="16513-330">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-330">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-331">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-331">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-332">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-332">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-333">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-333">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-334">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-334">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-335">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-335">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-336">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-336">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-337">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-337">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="16513-338">Caixa de correio 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-338">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="16513-339">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-340">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-340">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-341">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-341">Mail Read</span></span><br>
      - <span data-ttu-id="16513-342">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-342">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-343">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-343">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="16513-344">Módulos</span><span class="sxs-lookup"><span data-stu-id="16513-344">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-345">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-345">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-346">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-346">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-347">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-347">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-348">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-348">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-349">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-349">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-350">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-350">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="16513-351">Caixa de correio 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-351">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="16513-352">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-353">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-353">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-354">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-354">Mail Read</span></span><br>
      - <span data-ttu-id="16513-355">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-355">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-356">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-356">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="16513-357">Módulos</span><span class="sxs-lookup"><span data-stu-id="16513-357">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-358">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-358">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-359">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-359">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-360">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-360">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-361">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-361">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-362">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-362">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-363">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-363">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="16513-364">Caixa de correio 1.7</span><span class="sxs-lookup"><span data-stu-id="16513-364">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="16513-365">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-366">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-366">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-367">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-367">Mail Read</span></span><br>
      - <span data-ttu-id="16513-368">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-368">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-369">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-369">Add-in Commands</span></span></a><br>
      - <span data-ttu-id="16513-370">Módulos</span><span class="sxs-lookup"><span data-stu-id="16513-370">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-371">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-371">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-372">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-372">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-373">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-373">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-374">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-374">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="16513-375">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-376">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-376">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-377">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-377">Mail Read</span></span><br>
      - <span data-ttu-id="16513-378">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-378">Mail Compose</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-379">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-379">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-380">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-380">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-381">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-381">Mailbox 1.3</span></span></a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-382">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-382">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="16513-383">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-384">Office 365 para iOS</span><span class="sxs-lookup"><span data-stu-id="16513-384">Office 365 for iOS</span></span></td>
    <td> - <span data-ttu-id="16513-385">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-385">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-386">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-386">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-387">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-387">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-388">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-388">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-389">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-389">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-390">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-390">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-391">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-391">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="16513-392">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-393">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-393">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-394">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-394">Mail Read</span></span><br>
      - <span data-ttu-id="16513-395">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-395">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-396">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-396">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-397">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-397">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-398">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-398">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-399">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-399">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-400">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-400">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-401">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-401">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-402">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-402">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="16513-403">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-404">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-404">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-405">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-405">Mail Read</span></span><br>
      - <span data-ttu-id="16513-406">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-406">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-407">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-407">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-408">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-408">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-409">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-409">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-410">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-410">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-411">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-411">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-412">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-412">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-413">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-413">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="16513-414">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-415">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-415">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-416">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-416">Mail Read</span></span><br>
      - <span data-ttu-id="16513-417">Composição de email</span><span class="sxs-lookup"><span data-stu-id="16513-417">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-418">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-418">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-419">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-419">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-420">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-420">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-421">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-421">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-422">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-422">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-423">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-423">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="16513-424">Caixa de correio 1.6</span><span class="sxs-lookup"><span data-stu-id="16513-424">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="16513-425">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-426">Office 365 para Android</span><span class="sxs-lookup"><span data-stu-id="16513-426">Office 365 for Android</span></span></td>
    <td> - <span data-ttu-id="16513-427">Email lido</span><span class="sxs-lookup"><span data-stu-id="16513-427">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-428">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-428">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="16513-429">Caixa de correio 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-429">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="16513-430">Caixa de correio 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-430">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="16513-431">Caixa de correio 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-431">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="16513-432">Caixa de correio 1.4</span><span class="sxs-lookup"><span data-stu-id="16513-432">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="16513-433">Caixa de correio 1.5</span><span class="sxs-lookup"><span data-stu-id="16513-433">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="16513-434">Não disponível</span><span class="sxs-lookup"><span data-stu-id="16513-434">Not available</span></span></td>
  </tr>
</table>

*<span data-ttu-id="16513-435">&ast; – Adicionado com atualizações pós-lançamento.</span><span class="sxs-lookup"><span data-stu-id="16513-435">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="word"></a><span data-ttu-id="16513-436">Word</span><span class="sxs-lookup"><span data-stu-id="16513-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16513-437">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-437">Platform</span></span></th>
    <th><span data-ttu-id="16513-438">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-438">Extension points</span></span></th>
    <th><span data-ttu-id="16513-439">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-439">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-440">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-440">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="16513-441">Office Online</span></span></td>
    <td> - <span data-ttu-id="16513-442">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-442">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-443">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-443">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-444">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-444">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-445">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-445">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-446">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-446">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-447">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-447">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-448">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-448">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-449">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-449">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-450">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-450">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-451">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-451">File</span></span><br>
         - <span data-ttu-id="16513-452">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-452">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-453">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-453">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-454">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-454">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-455">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-455">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-456">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-456">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-457">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-457">PdfFile</span></span><br>
         - <span data-ttu-id="16513-458">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-458">Selection</span></span><br>
         - <span data-ttu-id="16513-459">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-459">Settings</span></span><br>
         - <span data-ttu-id="16513-460">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-460">TableBindings</span></span><br>
         - <span data-ttu-id="16513-461">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-461">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-462">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-462">TextBindings</span></span><br>
         - <span data-ttu-id="16513-463">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-463">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-464">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-464">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-465">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-465">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-466">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-466">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-467">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-467">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-468">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-468">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-469">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-469">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-470">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-470">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-471">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-471">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-472">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-472">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-473">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-473">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-474">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-474">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-475">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-475">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-476">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-476">File</span></span><br>
         - <span data-ttu-id="16513-477">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-477">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-478">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-478">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-479">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-479">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-480">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-480">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-481">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-481">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-482">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-482">PdfFile</span></span><br>
         - <span data-ttu-id="16513-483">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-483">Selection</span></span><br>
         - <span data-ttu-id="16513-484">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-484">Settings</span></span><br>
         - <span data-ttu-id="16513-485">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-485">TableBindings</span></span><br>
         - <span data-ttu-id="16513-486">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-486">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-487">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-487">TextBindings</span></span><br>
         - <span data-ttu-id="16513-488">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-488">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-489">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-489">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-490">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-490">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-491">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-491">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-492">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-492">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-493">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-493">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-494">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-494">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-495">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-495">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-496">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-496">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-497">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-497">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-498">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-498">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-499">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-499">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-500">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-500">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-501">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-501">File</span></span><br>
         - <span data-ttu-id="16513-502">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-502">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-503">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-503">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-504">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-504">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-505">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-505">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-506">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-506">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-507">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-507">PdfFile</span></span><br>
         - <span data-ttu-id="16513-508">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-508">Selection</span></span><br>
         - <span data-ttu-id="16513-509">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-509">Settings</span></span><br>
         - <span data-ttu-id="16513-510">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-510">TableBindings</span></span><br>
         - <span data-ttu-id="16513-511">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-511">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-512">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-512">TextBindings</span></span><br>
         - <span data-ttu-id="16513-513">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-513">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-514">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-514">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-515">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-515">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-516">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-516">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-517">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-517">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-518">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-518">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-519">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-519">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-520">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-520">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-521">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-521">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-522">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-522">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-523">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-523">File</span></span><br>
         - <span data-ttu-id="16513-524">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-524">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-525">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-525">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-526">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-526">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-527">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-527">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-528">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-528">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-529">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-529">PdfFile</span></span><br>
         - <span data-ttu-id="16513-530">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-530">Selection</span></span><br>
         - <span data-ttu-id="16513-531">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-531">Settings</span></span><br>
         - <span data-ttu-id="16513-532">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-532">TableBindings</span></span><br>
         - <span data-ttu-id="16513-533">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-533">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-534">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-534">TextBindings</span></span><br>
         - <span data-ttu-id="16513-535">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-535">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-536">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-536">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-537">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-537">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-538">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-538">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-539">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-539">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-540">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-540">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-541">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-541">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-542">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-542">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-543">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-543">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-544">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-544">File</span></span><br>
         - <span data-ttu-id="16513-545">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-545">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-546">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-546">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-547">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-547">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-548">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-548">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-549">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-549">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-550">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-550">PdfFile</span></span><br>
         - <span data-ttu-id="16513-551">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-551">Selection</span></span><br>
         - <span data-ttu-id="16513-552">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-552">Settings</span></span><br>
         - <span data-ttu-id="16513-553">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-553">TableBindings</span></span><br>
         - <span data-ttu-id="16513-554">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-554">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-555">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-555">TextBindings</span></span><br>
         - <span data-ttu-id="16513-556">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-556">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-557">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-557">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-558">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="16513-558">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="16513-559">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-559">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-560">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-560">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-561">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-561">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-562">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-562">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-563">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-563">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="16513-564">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-564">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-565">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-565">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-566">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-566">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-567">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-567">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-568">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-568">File</span></span><br>
         - <span data-ttu-id="16513-569">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-569">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-570">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-570">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-571">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-571">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-572">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-572">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-573">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-573">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-574">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-574">PdfFile</span></span><br>
         - <span data-ttu-id="16513-575">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-575">Selection</span></span><br>
         - <span data-ttu-id="16513-576">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-576">Settings</span></span><br>
         - <span data-ttu-id="16513-577">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-577">TableBindings</span></span><br>
         - <span data-ttu-id="16513-578">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-578">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-579">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-579">TextBindings</span></span><br>
         - <span data-ttu-id="16513-580">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-580">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-581">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-581">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-582">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-582">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-583">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-583">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-584">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-584">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-585">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-585">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-586">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-586">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-587">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-587">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-588">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-588">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="16513-589">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-589">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-590">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-590">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-591">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-591">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-592">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-592">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-593">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-593">File</span></span><br>
         - <span data-ttu-id="16513-594">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-594">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-595">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-595">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-596">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-596">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-597">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-597">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-598">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-598">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-599">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-599">PdfFile</span></span><br>
         - <span data-ttu-id="16513-600">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-600">Selection</span></span><br>
         - <span data-ttu-id="16513-601">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-601">Settings</span></span><br>
         - <span data-ttu-id="16513-602">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-602">TableBindings</span></span><br>
         - <span data-ttu-id="16513-603">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-603">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-604">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-604">TextBindings</span></span><br>
         - <span data-ttu-id="16513-605">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-605">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-606">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-606">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-607">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-607">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-608">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-608">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-609">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-609">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-610">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-610">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-611">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="16513-611">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-612">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="16513-612">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-613">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-613">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="16513-614">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-614">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-615">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-615">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-616">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-616">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-617">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-617">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-618">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-618">File</span></span><br>
         - <span data-ttu-id="16513-619">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-619">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-620">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-620">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-621">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-621">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-622">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-622">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-623">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-623">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-624">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-624">PdfFile</span></span><br>
         - <span data-ttu-id="16513-625">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-625">Selection</span></span><br>
         - <span data-ttu-id="16513-626">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-626">Settings</span></span><br>
         - <span data-ttu-id="16513-627">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-627">TableBindings</span></span><br>
         - <span data-ttu-id="16513-628">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-628">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-629">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-629">TextBindings</span></span><br>
         - <span data-ttu-id="16513-630">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-630">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-631">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-631">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-632">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-632">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-633">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-633">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="16513-634">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-634">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-635">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-635">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-636">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="16513-636">BindingEvents</span></span><br>
         - <span data-ttu-id="16513-637">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-637">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-638">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="16513-638">CustomXmlParts</span></span><br>
         - <span data-ttu-id="16513-639">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-639">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-640">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-640">File</span></span><br>
         - <span data-ttu-id="16513-641">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-641">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-642">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-642">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-643">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="16513-643">MatrixBindings</span></span><br>
         - <span data-ttu-id="16513-644">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-644">MatrixCoercion</span></span><br>
         - <span data-ttu-id="16513-645">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-645">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="16513-646">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-646">PdfFile</span></span><br>
         - <span data-ttu-id="16513-647">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-647">Selection</span></span><br>
         - <span data-ttu-id="16513-648">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-648">Settings</span></span><br>
         - <span data-ttu-id="16513-649">TableBindings</span><span class="sxs-lookup"><span data-stu-id="16513-649">TableBindings</span></span><br>
         - <span data-ttu-id="16513-650">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-650">TableCoercion</span></span><br>
         - <span data-ttu-id="16513-651">TextBindings</span><span class="sxs-lookup"><span data-stu-id="16513-651">TextBindings</span></span><br>
         - <span data-ttu-id="16513-652">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-652">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-653">TextFile</span><span class="sxs-lookup"><span data-stu-id="16513-653">TextFile</span></span> </td>
  </tr>
</table>

*<span data-ttu-id="16513-654">&ast; – Adicionado com atualizações pós-lançamento.</span><span class="sxs-lookup"><span data-stu-id="16513-654">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="powerpoint"></a><span data-ttu-id="16513-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="16513-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16513-656">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-656">Platform</span></span></th>
    <th><span data-ttu-id="16513-657">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-657">Extension points</span></span></th>
    <th><span data-ttu-id="16513-658">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-658">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-659">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-659">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="16513-660">Office Online</span></span></td>
    <td> - <span data-ttu-id="16513-661">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-661">Content</span></span><br>
         - <span data-ttu-id="16513-662">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-662">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-663">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-663">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-664">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-664">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-665">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-665">ActiveView</span></span><br>
         - <span data-ttu-id="16513-666">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-666">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-667">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-667">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-668">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-668">File</span></span><br>
         - <span data-ttu-id="16513-669">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-669">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-670">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-670">PdfFile</span></span><br>
         - <span data-ttu-id="16513-671">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-671">Selection</span></span><br>
         - <span data-ttu-id="16513-672">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-672">Settings</span></span><br>
         - <span data-ttu-id="16513-673">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-673">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-674">Office 365 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-674">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-675">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-675">Content</span></span><br>
         - <span data-ttu-id="16513-676">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-676">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-677">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-677">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-678">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-678">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-679">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-679">ActiveView</span></span><br>
         - <span data-ttu-id="16513-680">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-680">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-681">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-681">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-682">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-682">File</span></span><br>
         - <span data-ttu-id="16513-683">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-683">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-684">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-684">PdfFile</span></span><br>
         - <span data-ttu-id="16513-685">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-685">Selection</span></span><br>
         - <span data-ttu-id="16513-686">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-686">Settings</span></span><br>
         - <span data-ttu-id="16513-687">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-687">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-688">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-688">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-689">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-689">Content</span></span><br>
         - <span data-ttu-id="16513-690">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-690">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-691">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-691">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-692">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-692">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-693">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-693">ActiveView</span></span><br>
         - <span data-ttu-id="16513-694">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-694">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-695">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-695">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-696">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-696">File</span></span><br>
         - <span data-ttu-id="16513-697">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-697">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-698">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-698">PdfFile</span></span><br>
         - <span data-ttu-id="16513-699">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-699">Selection</span></span><br>
         - <span data-ttu-id="16513-700">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-700">Settings</span></span><br>
         - <span data-ttu-id="16513-701">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-701">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-702">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-702">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-703">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-703">Content</span></span><br>
         - <span data-ttu-id="16513-704">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-704">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-705">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-705">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-706">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-706">ActiveView</span></span><br>
         - <span data-ttu-id="16513-707">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-707">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-708">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-708">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-709">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-709">File</span></span><br>
         - <span data-ttu-id="16513-710">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-710">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-711">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-711">PdfFile</span></span><br>
         - <span data-ttu-id="16513-712">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-712">Selection</span></span><br>
         - <span data-ttu-id="16513-713">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-713">Settings</span></span><br>
         - <span data-ttu-id="16513-714">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-714">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-715">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-715">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-716">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-716">Content</span></span><br>
         - <span data-ttu-id="16513-717">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-717">TaskPane</span></span><br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-718">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-718">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-719">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-719">ActiveView</span></span><br>
         - <span data-ttu-id="16513-720">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-720">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-721">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-721">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-722">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-722">File</span></span><br>
         - <span data-ttu-id="16513-723">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-723">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-724">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-724">PdfFile</span></span><br>
         - <span data-ttu-id="16513-725">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-725">Selection</span></span><br>
         - <span data-ttu-id="16513-726">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-726">Settings</span></span><br>
         - <span data-ttu-id="16513-727">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-727">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-728">Office 365 para iPad</span><span class="sxs-lookup"><span data-stu-id="16513-728">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="16513-729">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-729">Content</span></span><br>
         - <span data-ttu-id="16513-730">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-730">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-731">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-731">DialogApi 1.1</span></span></a></td>
     <td> - <span data-ttu-id="16513-732">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-732">ActiveView</span></span><br>
         - <span data-ttu-id="16513-733">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-733">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-734">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-734">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-735">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-735">File</span></span><br>
         - <span data-ttu-id="16513-736">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-736">PdfFile</span></span><br>
         - <span data-ttu-id="16513-737">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-737">Selection</span></span><br>
         - <span data-ttu-id="16513-738">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-738">Settings</span></span><br>
         - <span data-ttu-id="16513-739">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-739">TextCoercion</span></span><br>
         - <span data-ttu-id="16513-740">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-740">ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-741">Office 365 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-741">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-742">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-742">Content</span></span><br>
         - <span data-ttu-id="16513-743">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-743">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-744">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-744">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-745">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-745">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-746">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-746">ActiveView</span></span><br>
         - <span data-ttu-id="16513-747">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-747">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-748">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-748">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-749">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-749">File</span></span><br>
         - <span data-ttu-id="16513-750">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-750">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-751">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-751">PdfFile</span></span><br>
         - <span data-ttu-id="16513-752">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-752">Selection</span></span><br>
         - <span data-ttu-id="16513-753">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-753">Settings</span></span><br>
         - <span data-ttu-id="16513-754">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-754">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-755">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-755">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-756">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-756">Content</span></span><br>
         - <span data-ttu-id="16513-757">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-757">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-758">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-758">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-759">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-759">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-760">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-760">ActiveView</span></span><br>
         - <span data-ttu-id="16513-761">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-761">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-762">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-762">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-763">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-763">File</span></span><br>
         - <span data-ttu-id="16513-764">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-764">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-765">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-765">PdfFile</span></span><br>
         - <span data-ttu-id="16513-766">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-766">Selection</span></span><br>
         - <span data-ttu-id="16513-767">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-767">Settings</span></span><br>
         - <span data-ttu-id="16513-768">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-768">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-769">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="16513-769">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="16513-770">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-770">Content</span></span><br>
         - <span data-ttu-id="16513-771">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-771">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-772">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-772">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="16513-773">ActiveView</span><span class="sxs-lookup"><span data-stu-id="16513-773">ActiveView</span></span><br>
         - <span data-ttu-id="16513-774">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="16513-774">CompressedFile</span></span><br>
         - <span data-ttu-id="16513-775">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-775">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-776">Arquivo</span><span class="sxs-lookup"><span data-stu-id="16513-776">File</span></span><br>
         - <span data-ttu-id="16513-777">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-777">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-778">PdfFile</span><span class="sxs-lookup"><span data-stu-id="16513-778">PdfFile</span></span><br>
         - <span data-ttu-id="16513-779">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-779">Selection</span></span><br>
         - <span data-ttu-id="16513-780">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-780">Settings</span></span><br>
         - <span data-ttu-id="16513-781">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-781">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="16513-782">&ast; – Adicionado com atualizações pós-lançamento.</span><span class="sxs-lookup"><span data-stu-id="16513-782">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="onenote"></a><span data-ttu-id="16513-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="16513-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16513-784">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-784">Platform</span></span></th>
    <th><span data-ttu-id="16513-785">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-785">Extension points</span></span></th>
    <th><span data-ttu-id="16513-786">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-786">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-787">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-787">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="16513-788">Office Online</span></span></td>
    <td> - <span data-ttu-id="16513-789">Conteúdo</span><span class="sxs-lookup"><span data-stu-id="16513-789">Content</span></span><br>
         - <span data-ttu-id="16513-790">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-790">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="16513-791">Comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-791">Add-in Commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets"><span data-ttu-id="16513-792">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-792">OneNoteApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-793">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-793">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-794">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="16513-794">DocumentEvents</span></span><br>
         - <span data-ttu-id="16513-795">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-795">HtmlCoercion</span></span><br>
         - <span data-ttu-id="16513-796">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-796">ImageCoercion</span></span><br>
         - <span data-ttu-id="16513-797">Configurações</span><span class="sxs-lookup"><span data-stu-id="16513-797">Settings</span></span><br>
         - <span data-ttu-id="16513-798">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-798">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="16513-799">Project</span><span class="sxs-lookup"><span data-stu-id="16513-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="16513-800">Plataforma</span><span class="sxs-lookup"><span data-stu-id="16513-800">Platform</span></span></th>
    <th><span data-ttu-id="16513-801">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="16513-801">Extension points</span></span></th>
    <th><span data-ttu-id="16513-802">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="16513-802">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="16513-803">APIs comuns</span><span class="sxs-lookup"><span data-stu-id="16513-803">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-804">Office 2019 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-804">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-805">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-805">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-806">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-806">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-807">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-807">Selection</span></span><br>
         - <span data-ttu-id="16513-808">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-808">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-809">Office 2016 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-809">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-810">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-810">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-811">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-811">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-812">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-812">Selection</span></span><br>
         - <span data-ttu-id="16513-813">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-813">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="16513-814">Office 2013 para Windows</span><span class="sxs-lookup"><span data-stu-id="16513-814">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="16513-815">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16513-815">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="16513-816">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="16513-816">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="16513-817">Seleção</span><span class="sxs-lookup"><span data-stu-id="16513-817">Selection</span></span><br>
         - <span data-ttu-id="16513-818">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="16513-818">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="16513-819">Confira também</span><span class="sxs-lookup"><span data-stu-id="16513-819">See also</span></span>

- [<span data-ttu-id="16513-820">Visão geral da plataforma de Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="16513-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="16513-821">Conjuntos de requisitos da API Comum</span><span class="sxs-lookup"><span data-stu-id="16513-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="16513-822">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="16513-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="16513-823">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="16513-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="16513-824">Office 2016 e o histórico de atualização de 2019 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="16513-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="16513-825">histórico de atualização do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="16513-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="16513-826">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16513-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="16513-827">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="16513-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)