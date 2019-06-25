---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 06/13/2019
localization_priority: Priority
ms.openlocfilehash: 82c276c802cab66ae4f5443d0d556bc42ee57841
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128619"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="48036-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="48036-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="48036-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="48036-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="48036-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="48036-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="48036-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="48036-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="48036-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="48036-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="48036-108">Excel</span><span class="sxs-lookup"><span data-stu-id="48036-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="48036-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="48036-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="48036-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="48036-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="48036-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-114">- TaskPane</span></span><br><span data-ttu-id="48036-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-115">
        - Content</span></span><br><span data-ttu-id="48036-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-116">
        - Custom Functions</span></span><br><span data-ttu-id="48036-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="48036-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="48036-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="48036-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="48036-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-128">
        - BindingEvents</span></span><br><span data-ttu-id="48036-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-129">
        - CompressedFile</span></span><br><span data-ttu-id="48036-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-130">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-131">
        - File</span></span><br><span data-ttu-id="48036-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-132">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-134">
        - Selection</span></span><br><span data-ttu-id="48036-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-135">
        - Settings</span></span><br><span data-ttu-id="48036-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-136">
        - TableBindings</span></span><br><span data-ttu-id="48036-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-137">
        - TableCoercion</span></span><br><span data-ttu-id="48036-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-138">
        - TextBindings</span></span><br><span data-ttu-id="48036-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-140">Office on Windows</span></span><br><span data-ttu-id="48036-141">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-142">- TaskPane</span></span><br><span data-ttu-id="48036-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-143">
        - Content</span></span><br><span data-ttu-id="48036-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-144">
        - Custom Functions</span></span><br><span data-ttu-id="48036-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="48036-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="48036-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="48036-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="48036-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-156">
        - BindingEvents</span></span><br><span data-ttu-id="48036-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-157">
        - CompressedFile</span></span><br><span data-ttu-id="48036-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-158">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-159">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-159">
        - File</span></span><br><span data-ttu-id="48036-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-160">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-162">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-162">
        - Selection</span></span><br><span data-ttu-id="48036-163">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-163">
        - Settings</span></span><br><span data-ttu-id="48036-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-164">
        - TableBindings</span></span><br><span data-ttu-id="48036-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-165">
        - TableCoercion</span></span><br><span data-ttu-id="48036-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-166">
        - TextBindings</span></span><br><span data-ttu-id="48036-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-168">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-168">Office 2019 on Windows</span></span><br><span data-ttu-id="48036-169">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="48036-170">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-170">- TaskPane</span></span><br><span data-ttu-id="48036-171">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-171">
        - Content</span></span><br><span data-ttu-id="48036-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="48036-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-182">- BindingEvents</span></span><br><span data-ttu-id="48036-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-183">
        - CompressedFile</span></span><br><span data-ttu-id="48036-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-184">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-185">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-185">
        - File</span></span><br><span data-ttu-id="48036-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-186">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-187">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-189">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-189">
        - Selection</span></span><br><span data-ttu-id="48036-190">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-190">
        - Settings</span></span><br><span data-ttu-id="48036-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-191">
        - TableBindings</span></span><br><span data-ttu-id="48036-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-192">
        - TableCoercion</span></span><br><span data-ttu-id="48036-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-193">
        - TextBindings</span></span><br><span data-ttu-id="48036-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-195">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-195">Office 2016 on Windows</span></span><br><span data-ttu-id="48036-196">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="48036-197">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-197">- TaskPane</span></span><br><span data-ttu-id="48036-198">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-198">
        - Content</span></span></td>
    <td><span data-ttu-id="48036-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="48036-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-201">- BindingEvents</span></span><br><span data-ttu-id="48036-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-202">
        - CompressedFile</span></span><br><span data-ttu-id="48036-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-203">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-204">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-204">
        - File</span></span><br><span data-ttu-id="48036-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-205">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-206">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-208">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-208">
        - Selection</span></span><br><span data-ttu-id="48036-209">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-209">
        - Settings</span></span><br><span data-ttu-id="48036-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-210">
        - TableBindings</span></span><br><span data-ttu-id="48036-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-211">
        - TableCoercion</span></span><br><span data-ttu-id="48036-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-212">
        - TextBindings</span></span><br><span data-ttu-id="48036-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-214">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-214">Office 2013 on Windows</span></span><br><span data-ttu-id="48036-215">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="48036-216">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-216">
        - TaskPane</span></span><br><span data-ttu-id="48036-217">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="48036-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="48036-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="48036-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-219">
        - BindingEvents</span></span><br><span data-ttu-id="48036-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-220">
        - CompressedFile</span></span><br><span data-ttu-id="48036-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-221">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-222">
        - File</span></span><br><span data-ttu-id="48036-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-223">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-224">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-226">
        - Selection</span></span><br><span data-ttu-id="48036-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-227">
        - Settings</span></span><br><span data-ttu-id="48036-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-228">
        - TableBindings</span></span><br><span data-ttu-id="48036-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-229">
        - TableCoercion</span></span><br><span data-ttu-id="48036-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-230">
        - TextBindings</span></span><br><span data-ttu-id="48036-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-232">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="48036-232">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="48036-233">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-233">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="48036-234">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-234">- TaskPane</span></span><br><span data-ttu-id="48036-235">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-235">
        - Content</span></span><br><span data-ttu-id="48036-236">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="48036-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="48036-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="48036-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-247">- BindingEvents</span></span><br><span data-ttu-id="48036-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-248">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-249">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-249">
        - File</span></span><br><span data-ttu-id="48036-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-250">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-251">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-253">
        - Selection</span></span><br><span data-ttu-id="48036-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-254">
        - Settings</span></span><br><span data-ttu-id="48036-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-255">
        - TableBindings</span></span><br><span data-ttu-id="48036-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-256">
        - TableCoercion</span></span><br><span data-ttu-id="48036-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-257">
        - TextBindings</span></span><br><span data-ttu-id="48036-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-259">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-259">Office apps on Mac</span></span><br><span data-ttu-id="48036-260">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-260">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="48036-261">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-261">- TaskPane</span></span><br><span data-ttu-id="48036-262">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-262">
        - Content</span></span><br><span data-ttu-id="48036-263">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-263">
        - Custom Functions</span></span><br><span data-ttu-id="48036-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="48036-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="48036-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="48036-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-275">- BindingEvents</span></span><br><span data-ttu-id="48036-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-276">
        - CompressedFile</span></span><br><span data-ttu-id="48036-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-277">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-278">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-278">
        - File</span></span><br><span data-ttu-id="48036-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-279">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-280">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-282">
        - PdfFile</span></span><br><span data-ttu-id="48036-283">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-283">
        - Selection</span></span><br><span data-ttu-id="48036-284">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-284">
        - Settings</span></span><br><span data-ttu-id="48036-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-285">
        - TableBindings</span></span><br><span data-ttu-id="48036-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-286">
        - TableCoercion</span></span><br><span data-ttu-id="48036-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-287">
        - TextBindings</span></span><br><span data-ttu-id="48036-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-289">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-289">Office 2019 for Mac</span></span><br><span data-ttu-id="48036-290">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="48036-291">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-291">- TaskPane</span></span><br><span data-ttu-id="48036-292">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-292">
        - Content</span></span><br><span data-ttu-id="48036-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="48036-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="48036-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="48036-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="48036-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="48036-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="48036-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="48036-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="48036-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="48036-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="48036-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-303">- BindingEvents</span></span><br><span data-ttu-id="48036-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-304">
        - CompressedFile</span></span><br><span data-ttu-id="48036-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-305">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-306">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-306">
        - File</span></span><br><span data-ttu-id="48036-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-307">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-308">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-310">
        - PdfFile</span></span><br><span data-ttu-id="48036-311">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-311">
        - Selection</span></span><br><span data-ttu-id="48036-312">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-312">
        - Settings</span></span><br><span data-ttu-id="48036-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-313">
        - TableBindings</span></span><br><span data-ttu-id="48036-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-314">
        - TableCoercion</span></span><br><span data-ttu-id="48036-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-315">
        - TextBindings</span></span><br><span data-ttu-id="48036-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-317">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-317">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="48036-318">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="48036-319">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-319">- TaskPane</span></span><br><span data-ttu-id="48036-320">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-320">
        - Content</span></span></td>
    <td><span data-ttu-id="48036-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="48036-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="48036-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-323">- BindingEvents</span></span><br><span data-ttu-id="48036-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-324">
        - CompressedFile</span></span><br><span data-ttu-id="48036-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-325">
        - DocumentEvents</span></span><br><span data-ttu-id="48036-326">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-326">
        - File</span></span><br><span data-ttu-id="48036-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-327">
        - ImageCoercion</span></span><br><span data-ttu-id="48036-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-328">
        - MatrixBindings</span></span><br><span data-ttu-id="48036-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="48036-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-330">
        - PdfFile</span></span><br><span data-ttu-id="48036-331">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-331">
        - Selection</span></span><br><span data-ttu-id="48036-332">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-332">
        - Settings</span></span><br><span data-ttu-id="48036-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-333">
        - TableBindings</span></span><br><span data-ttu-id="48036-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-334">
        - TableCoercion</span></span><br><span data-ttu-id="48036-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-335">
        - TextBindings</span></span><br><span data-ttu-id="48036-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="48036-337">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="48036-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="48036-338">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="48036-339">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="48036-340">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="48036-341">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="48036-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-343">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-343">Office on the web</span></span></td>
    <td><span data-ttu-id="48036-344">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="48036-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-346">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-346">Office on Windows</span></span><br><span data-ttu-id="48036-347">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-347">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="48036-348">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="48036-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-350">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="48036-350">Office for Mac</span></span><br><span data-ttu-id="48036-351">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="48036-352">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="48036-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="48036-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="48036-354">Outlook</span><span class="sxs-lookup"><span data-stu-id="48036-354">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="48036-355">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-355">Platform</span></span></th>
    <th><span data-ttu-id="48036-356">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-356">Extension points</span></span></th>
    <th><span data-ttu-id="48036-357">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="48036-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-359">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-359">Office on the web</span></span><br><span data-ttu-id="48036-360">(novo)</span><span class="sxs-lookup"><span data-stu-id="48036-360">New</span></span></td>
    <td> <span data-ttu-id="48036-361">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-361">- Mail Read</span></span><br><span data-ttu-id="48036-362">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-362">
      - Mail Compose</span></span><br><span data-ttu-id="48036-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="48036-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="48036-371">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-372">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-372">Office on the web</span></span><br><span data-ttu-id="48036-373">(clássico)</span><span class="sxs-lookup"><span data-stu-id="48036-373">Classic.</span></span></td>
    <td> <span data-ttu-id="48036-374">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-374">- Mail Read</span></span><br><span data-ttu-id="48036-375">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-375">
      - Mail Compose</span></span><br><span data-ttu-id="48036-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="48036-383">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-384">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-384">Office on Windows</span></span><br><span data-ttu-id="48036-385">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-385">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-386">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-386">- Mail Read</span></span><br><span data-ttu-id="48036-387">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-387">
      - Mail Compose</span></span><br><span data-ttu-id="48036-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="48036-389">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="48036-389">
      - Modules</span></span></td>
    <td> <span data-ttu-id="48036-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="48036-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="48036-397">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-398">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-398">Office 2019 on Windows</span></span><br><span data-ttu-id="48036-399">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-399">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-400">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-400">- Mail Read</span></span><br><span data-ttu-id="48036-401">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-401">
      - Mail Compose</span></span><br><span data-ttu-id="48036-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="48036-403">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="48036-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="48036-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="48036-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="48036-411">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-412">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-412">Office 2016 on Windows</span></span><br><span data-ttu-id="48036-413">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-413">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-414">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-414">- Mail Read</span></span><br><span data-ttu-id="48036-415">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-415">
      - Mail Compose</span></span><br><span data-ttu-id="48036-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="48036-417">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="48036-417">
      - Modules</span></span></td>
    <td> <span data-ttu-id="48036-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="48036-422">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-423">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-423">Office 2013 on Windows</span></span><br><span data-ttu-id="48036-424">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-424">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-425">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-425">- Mail Read</span></span><br><span data-ttu-id="48036-426">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-426">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="48036-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="48036-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="48036-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-432">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="48036-432">Office apps on iOS</span></span><br><span data-ttu-id="48036-433">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-433">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-434">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-434">- Mail Read</span></span><br><span data-ttu-id="48036-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="48036-441">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-442">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-442">Office apps on Mac</span></span><br><span data-ttu-id="48036-443">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-443">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-444">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-444">- Mail Read</span></span><br><span data-ttu-id="48036-445">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-445">
      - Mail Compose</span></span><br><span data-ttu-id="48036-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="48036-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="48036-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="48036-454">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-455">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-455">Office 2019 for Mac</span></span><br><span data-ttu-id="48036-456">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-456">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-457">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-457">- Mail Read</span></span><br><span data-ttu-id="48036-458">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-458">
      - Mail Compose</span></span><br><span data-ttu-id="48036-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="48036-466">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-467">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-467">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="48036-468">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-468">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-469">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-469">- Mail Read</span></span><br><span data-ttu-id="48036-470">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="48036-470">
      - Mail Compose</span></span><br><span data-ttu-id="48036-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="48036-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="48036-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="48036-478">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-479">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="48036-479">Office apps on Android</span></span><br><span data-ttu-id="48036-480">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-480">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-481">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="48036-481">- Mail Read</span></span><br><span data-ttu-id="48036-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="48036-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="48036-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="48036-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="48036-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="48036-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="48036-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="48036-488">Não disponível</span><span class="sxs-lookup"><span data-stu-id="48036-488">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="48036-489">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="48036-489">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="48036-490">Word</span><span class="sxs-lookup"><span data-stu-id="48036-490">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="48036-491">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-491">Platform</span></span></th>
    <th><span data-ttu-id="48036-492">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-492">Extension points</span></span></th>
    <th><span data-ttu-id="48036-493">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-493">API requirement sets</span></span></th>
    <th><span data-ttu-id="48036-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-495">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-495">Office on the web</span></span></td>
    <td> <span data-ttu-id="48036-496">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-496">- TaskPane</span></span><br><span data-ttu-id="48036-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-502">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-502">- BindingEvents</span></span><br><span data-ttu-id="48036-503">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-503">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-504">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-504">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-505">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-505">
         - File</span></span><br><span data-ttu-id="48036-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-506">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-507">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-508">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-508">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-509">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-509">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-510">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-510">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-511">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-511">
         - PdfFile</span></span><br><span data-ttu-id="48036-512">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-512">
         - Selection</span></span><br><span data-ttu-id="48036-513">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-513">
         - Settings</span></span><br><span data-ttu-id="48036-514">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-514">
         - TableBindings</span></span><br><span data-ttu-id="48036-515">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-515">
         - TableCoercion</span></span><br><span data-ttu-id="48036-516">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-516">
         - TextBindings</span></span><br><span data-ttu-id="48036-517">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-517">
         - TextCoercion</span></span><br><span data-ttu-id="48036-518">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-518">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-519">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-519">Office on Windows</span></span><br><span data-ttu-id="48036-520">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-520">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-521">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-521">- TaskPane</span></span><br><span data-ttu-id="48036-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-527">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-527">- BindingEvents</span></span><br><span data-ttu-id="48036-528">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-528">
         - CompressedFile</span></span><br><span data-ttu-id="48036-529">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-529">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-530">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-530">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-531">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-531">
         - File</span></span><br><span data-ttu-id="48036-532">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-532">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-533">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-533">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-534">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-534">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-535">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-535">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-536">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-536">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-537">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-537">
         - PdfFile</span></span><br><span data-ttu-id="48036-538">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-538">
         - Selection</span></span><br><span data-ttu-id="48036-539">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-539">
         - Settings</span></span><br><span data-ttu-id="48036-540">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-540">
         - TableBindings</span></span><br><span data-ttu-id="48036-541">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-541">
         - TableCoercion</span></span><br><span data-ttu-id="48036-542">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-542">
         - TextBindings</span></span><br><span data-ttu-id="48036-543">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-543">
         - TextCoercion</span></span><br><span data-ttu-id="48036-544">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-544">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-545">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-545">Office 2019 on Windows</span></span><br><span data-ttu-id="48036-546">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-546">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-547">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-547">- TaskPane</span></span><br><span data-ttu-id="48036-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-553">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-553">- BindingEvents</span></span><br><span data-ttu-id="48036-554">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-554">
         - CompressedFile</span></span><br><span data-ttu-id="48036-555">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-555">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-556">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-557">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-557">
         - File</span></span><br><span data-ttu-id="48036-558">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-558">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-559">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-559">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-560">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-560">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-561">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-561">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-562">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-562">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-563">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-563">
         - PdfFile</span></span><br><span data-ttu-id="48036-564">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-564">
         - Selection</span></span><br><span data-ttu-id="48036-565">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-565">
         - Settings</span></span><br><span data-ttu-id="48036-566">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-566">
         - TableBindings</span></span><br><span data-ttu-id="48036-567">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-567">
         - TableCoercion</span></span><br><span data-ttu-id="48036-568">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-568">
         - TextBindings</span></span><br><span data-ttu-id="48036-569">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-569">
         - TextCoercion</span></span><br><span data-ttu-id="48036-570">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-570">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-571">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-571">Office 2016 on Windows</span></span><br><span data-ttu-id="48036-572">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-572">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-573">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-573">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="48036-576">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-576">- BindingEvents</span></span><br><span data-ttu-id="48036-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-577">
         - CompressedFile</span></span><br><span data-ttu-id="48036-578">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-578">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-579">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-580">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-580">
         - File</span></span><br><span data-ttu-id="48036-581">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-581">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-582">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-583">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-583">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-584">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-584">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-585">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-585">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-586">
         - PdfFile</span></span><br><span data-ttu-id="48036-587">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-587">
         - Selection</span></span><br><span data-ttu-id="48036-588">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-588">
         - Settings</span></span><br><span data-ttu-id="48036-589">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-589">
         - TableBindings</span></span><br><span data-ttu-id="48036-590">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-590">
         - TableCoercion</span></span><br><span data-ttu-id="48036-591">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-591">
         - TextBindings</span></span><br><span data-ttu-id="48036-592">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-592">
         - TextCoercion</span></span><br><span data-ttu-id="48036-593">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-593">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-594">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-594">Office 2013 on Windows</span></span><br><span data-ttu-id="48036-595">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-595">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-596">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-596">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="48036-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="48036-598">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-598">- BindingEvents</span></span><br><span data-ttu-id="48036-599">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-599">
         - CompressedFile</span></span><br><span data-ttu-id="48036-600">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-600">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-601">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-601">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-602">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-602">
         - File</span></span><br><span data-ttu-id="48036-603">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-603">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-604">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-604">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-605">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-605">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-606">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-606">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-607">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-607">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-608">
         - PdfFile</span></span><br><span data-ttu-id="48036-609">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-609">
         - Selection</span></span><br><span data-ttu-id="48036-610">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-610">
         - Settings</span></span><br><span data-ttu-id="48036-611">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-611">
         - TableBindings</span></span><br><span data-ttu-id="48036-612">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-612">
         - TableCoercion</span></span><br><span data-ttu-id="48036-613">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-613">
         - TextBindings</span></span><br><span data-ttu-id="48036-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-614">
         - TextCoercion</span></span><br><span data-ttu-id="48036-615">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-615">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-616">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="48036-616">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="48036-617">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-617">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-618">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-618">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="48036-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="48036-623">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-623">- BindingEvents</span></span><br><span data-ttu-id="48036-624">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-624">
         - CompressedFile</span></span><br><span data-ttu-id="48036-625">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-625">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-626">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-627">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-627">
         - File</span></span><br><span data-ttu-id="48036-628">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-628">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-629">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-629">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-630">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-630">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-631">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-631">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-632">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-632">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-633">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-633">
         - PdfFile</span></span><br><span data-ttu-id="48036-634">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-634">
         - Selection</span></span><br><span data-ttu-id="48036-635">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-635">
         - Settings</span></span><br><span data-ttu-id="48036-636">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-636">
         - TableBindings</span></span><br><span data-ttu-id="48036-637">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-637">
         - TableCoercion</span></span><br><span data-ttu-id="48036-638">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-638">
         - TextBindings</span></span><br><span data-ttu-id="48036-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-639">
         - TextCoercion</span></span><br><span data-ttu-id="48036-640">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-640">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-641">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-641">Office apps on Mac</span></span><br><span data-ttu-id="48036-642">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-642">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-643">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-643">- TaskPane</span></span><br><span data-ttu-id="48036-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="48036-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="48036-649">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-649">- BindingEvents</span></span><br><span data-ttu-id="48036-650">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-650">
         - CompressedFile</span></span><br><span data-ttu-id="48036-651">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-651">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-652">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-652">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-653">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-653">
         - File</span></span><br><span data-ttu-id="48036-654">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-654">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-655">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-655">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-656">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-656">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-657">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-657">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-658">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-658">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-659">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-659">
         - PdfFile</span></span><br><span data-ttu-id="48036-660">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-660">
         - Selection</span></span><br><span data-ttu-id="48036-661">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-661">
         - Settings</span></span><br><span data-ttu-id="48036-662">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-662">
         - TableBindings</span></span><br><span data-ttu-id="48036-663">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-663">
         - TableCoercion</span></span><br><span data-ttu-id="48036-664">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-664">
         - TextBindings</span></span><br><span data-ttu-id="48036-665">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-665">
         - TextCoercion</span></span><br><span data-ttu-id="48036-666">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-666">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-667">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-667">Office 2019 for Mac</span></span><br><span data-ttu-id="48036-668">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-668">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-669">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-669">- TaskPane</span></span><br><span data-ttu-id="48036-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="48036-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="48036-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="48036-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="48036-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="48036-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="48036-675">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-675">- BindingEvents</span></span><br><span data-ttu-id="48036-676">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-676">
         - CompressedFile</span></span><br><span data-ttu-id="48036-677">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-677">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-678">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-678">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-679">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-679">
         - File</span></span><br><span data-ttu-id="48036-680">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-680">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-681">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-681">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-682">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-682">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-683">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-683">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-684">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-684">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-685">
         - PdfFile</span></span><br><span data-ttu-id="48036-686">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-686">
         - Selection</span></span><br><span data-ttu-id="48036-687">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-687">
         - Settings</span></span><br><span data-ttu-id="48036-688">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-688">
         - TableBindings</span></span><br><span data-ttu-id="48036-689">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-689">
         - TableCoercion</span></span><br><span data-ttu-id="48036-690">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-690">
         - TextBindings</span></span><br><span data-ttu-id="48036-691">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-691">
         - TextCoercion</span></span><br><span data-ttu-id="48036-692">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-692">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-693">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-693">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="48036-694">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-694">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-695">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-695">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="48036-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="48036-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="48036-698">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="48036-698">- BindingEvents</span></span><br><span data-ttu-id="48036-699">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-699">
         - CompressedFile</span></span><br><span data-ttu-id="48036-700">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="48036-700">
         - CustomXmlParts</span></span><br><span data-ttu-id="48036-701">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-701">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-702">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-702">
         - File</span></span><br><span data-ttu-id="48036-703">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-703">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-704">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-704">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-705">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="48036-705">
         - MatrixBindings</span></span><br><span data-ttu-id="48036-706">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-706">
         - MatrixCoercion</span></span><br><span data-ttu-id="48036-707">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-707">
         - OoxmlCoercion</span></span><br><span data-ttu-id="48036-708">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-708">
         - PdfFile</span></span><br><span data-ttu-id="48036-709">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-709">
         - Selection</span></span><br><span data-ttu-id="48036-710">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-710">
         - Settings</span></span><br><span data-ttu-id="48036-711">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="48036-711">
         - TableBindings</span></span><br><span data-ttu-id="48036-712">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-712">
         - TableCoercion</span></span><br><span data-ttu-id="48036-713">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="48036-713">
         - TextBindings</span></span><br><span data-ttu-id="48036-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-714">
         - TextCoercion</span></span><br><span data-ttu-id="48036-715">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="48036-715">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="48036-716">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="48036-716">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="48036-717">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="48036-717">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="48036-718">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-718">Platform</span></span></th>
    <th><span data-ttu-id="48036-719">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-719">Extension points</span></span></th>
    <th><span data-ttu-id="48036-720">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-720">API requirement sets</span></span></th>
    <th><span data-ttu-id="48036-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-722">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-722">Office on the web</span></span></td>
    <td> <span data-ttu-id="48036-723">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-723">- Content</span></span><br><span data-ttu-id="48036-724">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-724">
         - TaskPane</span></span><br><span data-ttu-id="48036-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-727">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-727">- ActiveView</span></span><br><span data-ttu-id="48036-728">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-728">
         - CompressedFile</span></span><br><span data-ttu-id="48036-729">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-729">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-730">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-730">
         - File</span></span><br><span data-ttu-id="48036-731">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-731">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-732">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-732">
         - PdfFile</span></span><br><span data-ttu-id="48036-733">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-733">
         - Selection</span></span><br><span data-ttu-id="48036-734">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-734">
         - Settings</span></span><br><span data-ttu-id="48036-735">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-735">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-736">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-736">Office on Windows</span></span><br><span data-ttu-id="48036-737">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-737">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-738">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-738">- Content</span></span><br><span data-ttu-id="48036-739">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-739">
         - TaskPane</span></span><br><span data-ttu-id="48036-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-742">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-742">- ActiveView</span></span><br><span data-ttu-id="48036-743">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-743">
         - CompressedFile</span></span><br><span data-ttu-id="48036-744">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-744">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-745">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-745">
         - File</span></span><br><span data-ttu-id="48036-746">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-746">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-747">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-747">
         - PdfFile</span></span><br><span data-ttu-id="48036-748">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-748">
         - Selection</span></span><br><span data-ttu-id="48036-749">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-749">
         - Settings</span></span><br><span data-ttu-id="48036-750">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-750">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-751">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-751">Office 2019 on Windows</span></span><br><span data-ttu-id="48036-752">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-752">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-753">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-753">- Content</span></span><br><span data-ttu-id="48036-754">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-754">
         - TaskPane</span></span><br><span data-ttu-id="48036-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-757">- ActiveView</span></span><br><span data-ttu-id="48036-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-758">
         - CompressedFile</span></span><br><span data-ttu-id="48036-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-759">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-760">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-760">
         - File</span></span><br><span data-ttu-id="48036-761">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-761">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-762">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-762">
         - PdfFile</span></span><br><span data-ttu-id="48036-763">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-763">
         - Selection</span></span><br><span data-ttu-id="48036-764">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-764">
         - Settings</span></span><br><span data-ttu-id="48036-765">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-765">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-766">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-766">Office 2016 on Windows</span></span><br><span data-ttu-id="48036-767">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-767">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-768">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-768">- Content</span></span><br><span data-ttu-id="48036-769">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-769">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="48036-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="48036-771">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-771">- ActiveView</span></span><br><span data-ttu-id="48036-772">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-772">
         - CompressedFile</span></span><br><span data-ttu-id="48036-773">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-773">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-774">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-774">
         - File</span></span><br><span data-ttu-id="48036-775">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-775">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-776">
         - PdfFile</span></span><br><span data-ttu-id="48036-777">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-777">
         - Selection</span></span><br><span data-ttu-id="48036-778">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-778">
         - Settings</span></span><br><span data-ttu-id="48036-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-780">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-780">Office 2013 on Windows</span></span><br><span data-ttu-id="48036-781">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-782">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-782">- Content</span></span><br><span data-ttu-id="48036-783">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-783">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="48036-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="48036-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="48036-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-785">- ActiveView</span></span><br><span data-ttu-id="48036-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-786">
         - CompressedFile</span></span><br><span data-ttu-id="48036-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-787">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-788">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-788">
         - File</span></span><br><span data-ttu-id="48036-789">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-789">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-790">
         - PdfFile</span></span><br><span data-ttu-id="48036-791">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-791">
         - Selection</span></span><br><span data-ttu-id="48036-792">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-792">
         - Settings</span></span><br><span data-ttu-id="48036-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-794">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="48036-794">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="48036-795">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-795">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-796">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-796">- Content</span></span><br><span data-ttu-id="48036-797">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-797">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-799">- ActiveView</span></span><br><span data-ttu-id="48036-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-800">
         - CompressedFile</span></span><br><span data-ttu-id="48036-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-801">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-802">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-802">
         - File</span></span><br><span data-ttu-id="48036-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-803">
         - PdfFile</span></span><br><span data-ttu-id="48036-804">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-804">
         - Selection</span></span><br><span data-ttu-id="48036-805">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-805">
         - Settings</span></span><br><span data-ttu-id="48036-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-806">
         - TextCoercion</span></span><br><span data-ttu-id="48036-807">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-807">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-808">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-808">Office apps on Mac</span></span><br><span data-ttu-id="48036-809">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="48036-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="48036-810">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-810">- Content</span></span><br><span data-ttu-id="48036-811">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-811">
         - TaskPane</span></span><br><span data-ttu-id="48036-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-814">- ActiveView</span></span><br><span data-ttu-id="48036-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-815">
         - CompressedFile</span></span><br><span data-ttu-id="48036-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-816">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-817">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-817">
         - File</span></span><br><span data-ttu-id="48036-818">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-818">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-819">
         - PdfFile</span></span><br><span data-ttu-id="48036-820">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-820">
         - Selection</span></span><br><span data-ttu-id="48036-821">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-821">
         - Settings</span></span><br><span data-ttu-id="48036-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-823">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-823">Office 2019 for Mac</span></span><br><span data-ttu-id="48036-824">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-824">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-825">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-825">- Content</span></span><br><span data-ttu-id="48036-826">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-826">
         - TaskPane</span></span><br><span data-ttu-id="48036-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-829">- ActiveView</span></span><br><span data-ttu-id="48036-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-830">
         - CompressedFile</span></span><br><span data-ttu-id="48036-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-831">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-832">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-832">
         - File</span></span><br><span data-ttu-id="48036-833">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-833">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-834">
         - PdfFile</span></span><br><span data-ttu-id="48036-835">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-835">
         - Selection</span></span><br><span data-ttu-id="48036-836">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-836">
         - Settings</span></span><br><span data-ttu-id="48036-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-838">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="48036-838">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="48036-839">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-840">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-840">- Content</span></span><br><span data-ttu-id="48036-841">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-841">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="48036-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="48036-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="48036-843">- ActiveView</span></span><br><span data-ttu-id="48036-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="48036-844">
         - CompressedFile</span></span><br><span data-ttu-id="48036-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-845">
         - DocumentEvents</span></span><br><span data-ttu-id="48036-846">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="48036-846">
         - File</span></span><br><span data-ttu-id="48036-847">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-847">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="48036-848">
         - PdfFile</span></span><br><span data-ttu-id="48036-849">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-849">
         - Selection</span></span><br><span data-ttu-id="48036-850">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-850">
         - Settings</span></span><br><span data-ttu-id="48036-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-851">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="48036-852">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="48036-852">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="48036-853">OneNote</span><span class="sxs-lookup"><span data-stu-id="48036-853">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="48036-854">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-854">Platform</span></span></th>
    <th><span data-ttu-id="48036-855">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-855">Extension points</span></span></th>
    <th><span data-ttu-id="48036-856">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-856">API requirement sets</span></span></th>
    <th><span data-ttu-id="48036-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-858">Office na Web</span><span class="sxs-lookup"><span data-stu-id="48036-858">Office on the web</span></span></td>
    <td> <span data-ttu-id="48036-859">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="48036-859">- Content</span></span><br><span data-ttu-id="48036-860">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-860">
         - TaskPane</span></span><br><span data-ttu-id="48036-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="48036-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="48036-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="48036-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-864">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="48036-864">- DocumentEvents</span></span><br><span data-ttu-id="48036-865">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-865">
         - HtmlCoercion</span></span><br><span data-ttu-id="48036-866">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-866">
         - ImageCoercion</span></span><br><span data-ttu-id="48036-867">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="48036-867">
         - Settings</span></span><br><span data-ttu-id="48036-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="48036-869">Project</span><span class="sxs-lookup"><span data-stu-id="48036-869">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="48036-870">Plataforma</span><span class="sxs-lookup"><span data-stu-id="48036-870">Platform</span></span></th>
    <th><span data-ttu-id="48036-871">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="48036-871">Extension points</span></span></th>
    <th><span data-ttu-id="48036-872">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="48036-872">API requirement sets</span></span></th>
    <th><span data-ttu-id="48036-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="48036-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-874">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-874">Office 2019 on Windows</span></span><br><span data-ttu-id="48036-875">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-875">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-876">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-876">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-878">- Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-878">- Selection</span></span><br><span data-ttu-id="48036-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-879">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-880">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-880">Office 2016 on Windows</span></span><br><span data-ttu-id="48036-881">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-881">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-882">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-882">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-884">- Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-884">- Selection</span></span><br><span data-ttu-id="48036-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-885">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="48036-886">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="48036-886">Office 2013 on Windows</span></span><br><span data-ttu-id="48036-887">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="48036-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="48036-888">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="48036-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="48036-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="48036-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="48036-890">- Seleção</span><span class="sxs-lookup"><span data-stu-id="48036-890">- Selection</span></span><br><span data-ttu-id="48036-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="48036-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="48036-892">Confira também</span><span class="sxs-lookup"><span data-stu-id="48036-892">See also</span></span>

- [<span data-ttu-id="48036-893">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="48036-893">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="48036-894">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="48036-894">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="48036-895">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="48036-895">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="48036-896">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="48036-896">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="48036-897">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="48036-897">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="48036-898">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="48036-898">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="48036-899">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="48036-899">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="48036-900">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="48036-900">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="48036-901">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="48036-901">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="48036-902">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="48036-902">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="48036-903">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="48036-903">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
