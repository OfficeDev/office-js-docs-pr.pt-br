---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 19f2fa7f744345823c2700b04524ec20705035a8
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952366"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="42176-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="42176-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="42176-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="42176-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="42176-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="42176-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="42176-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="42176-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="42176-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="42176-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="42176-108">Excel</span><span class="sxs-lookup"><span data-stu-id="42176-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="42176-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="42176-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="42176-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="42176-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-113">Office Online</span></span></td>
    <td> <span data-ttu-id="42176-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-114">- TaskPane</span></span><br><span data-ttu-id="42176-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-115">
        - Content</span></span><br><span data-ttu-id="42176-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-116">
        -Custom Functions</span></span><br><span data-ttu-id="42176-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="42176-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="42176-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="42176-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="42176-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-128">
        - BindingEvents</span></span><br><span data-ttu-id="42176-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-129">
        - CompressedFile</span></span><br><span data-ttu-id="42176-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-130">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-131">
        - File</span></span><br><span data-ttu-id="42176-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-132">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-134">
        - Selection</span></span><br><span data-ttu-id="42176-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-135">
        - Settings</span></span><br><span data-ttu-id="42176-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-136">
        - TableBindings</span></span><br><span data-ttu-id="42176-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-137">
        - TableCoercion</span></span><br><span data-ttu-id="42176-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-138">
        - TextBindings</span></span><br><span data-ttu-id="42176-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-140">Office apps on Windows</span></span><br><span data-ttu-id="42176-141">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-142">- TaskPane</span></span><br><span data-ttu-id="42176-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-143">
        - Content</span></span><br><span data-ttu-id="42176-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-144">
        -Custom Functions</span></span><br><span data-ttu-id="42176-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="42176-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="42176-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="42176-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="42176-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-156">
        - BindingEvents</span></span><br><span data-ttu-id="42176-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-157">
        - CompressedFile</span></span><br><span data-ttu-id="42176-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-158">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-159">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-159">
        - File</span></span><br><span data-ttu-id="42176-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-160">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-162">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-162">
        - Selection</span></span><br><span data-ttu-id="42176-163">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-163">
        - Settings</span></span><br><span data-ttu-id="42176-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-164">
        - TableBindings</span></span><br><span data-ttu-id="42176-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-165">
        - TableCoercion</span></span><br><span data-ttu-id="42176-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-166">
        - TextBindings</span></span><br><span data-ttu-id="42176-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-168">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-168">Office 2019 for Windows</span></span><br><span data-ttu-id="42176-169">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="42176-170">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-170">- TaskPane</span></span><br><span data-ttu-id="42176-171">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-171">
        - Content</span></span><br><span data-ttu-id="42176-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="42176-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-182">- BindingEvents</span></span><br><span data-ttu-id="42176-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-183">
        - CompressedFile</span></span><br><span data-ttu-id="42176-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-184">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-185">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-185">
        - File</span></span><br><span data-ttu-id="42176-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-186">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-187">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-189">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-189">
        - Selection</span></span><br><span data-ttu-id="42176-190">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-190">
        - Settings</span></span><br><span data-ttu-id="42176-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-191">
        - TableBindings</span></span><br><span data-ttu-id="42176-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-192">
        - TableCoercion</span></span><br><span data-ttu-id="42176-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-193">
        - TextBindings</span></span><br><span data-ttu-id="42176-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-195">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-195">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="42176-196">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="42176-197">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-197">- TaskPane</span></span><br><span data-ttu-id="42176-198">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-198">
        - Content</span></span></td>
    <td><span data-ttu-id="42176-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="42176-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-201">- BindingEvents</span></span><br><span data-ttu-id="42176-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-202">
        - CompressedFile</span></span><br><span data-ttu-id="42176-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-203">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-204">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-204">
        - File</span></span><br><span data-ttu-id="42176-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-205">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-206">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-208">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-208">
        - Selection</span></span><br><span data-ttu-id="42176-209">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-209">
        - Settings</span></span><br><span data-ttu-id="42176-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-210">
        - TableBindings</span></span><br><span data-ttu-id="42176-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-211">
        - TableCoercion</span></span><br><span data-ttu-id="42176-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-212">
        - TextBindings</span></span><br><span data-ttu-id="42176-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-214">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-214">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="42176-215">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="42176-216">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-216">
        - TaskPane</span></span><br><span data-ttu-id="42176-217">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="42176-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="42176-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="42176-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-219">
        - BindingEvents</span></span><br><span data-ttu-id="42176-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-220">
        - CompressedFile</span></span><br><span data-ttu-id="42176-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-221">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-222">
        - File</span></span><br><span data-ttu-id="42176-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-223">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-224">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-226">
        - Selection</span></span><br><span data-ttu-id="42176-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-227">
        - Settings</span></span><br><span data-ttu-id="42176-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-228">
        - TableBindings</span></span><br><span data-ttu-id="42176-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-229">
        - TableCoercion</span></span><br><span data-ttu-id="42176-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-230">
        - TextBindings</span></span><br><span data-ttu-id="42176-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-232">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="42176-232">Office for iPad</span></span><br><span data-ttu-id="42176-233">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="42176-234">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-234">- TaskPane</span></span><br><span data-ttu-id="42176-235">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-235">
        - Content</span></span><br><span data-ttu-id="42176-236">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-236">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="42176-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="42176-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="42176-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-247">- BindingEvents</span></span><br><span data-ttu-id="42176-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-248">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-249">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-249">
        - File</span></span><br><span data-ttu-id="42176-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-250">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-251">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-253">
        - Selection</span></span><br><span data-ttu-id="42176-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-254">
        - Settings</span></span><br><span data-ttu-id="42176-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-255">
        - TableBindings</span></span><br><span data-ttu-id="42176-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-256">
        - TableCoercion</span></span><br><span data-ttu-id="42176-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-257">
        - TextBindings</span></span><br><span data-ttu-id="42176-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-259">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-259">Office for Mac</span></span><br><span data-ttu-id="42176-260">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="42176-261">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-261">- TaskPane</span></span><br><span data-ttu-id="42176-262">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-262">
        - Content</span></span><br><span data-ttu-id="42176-263">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-263">
        -Custom Functions</span></span><br><span data-ttu-id="42176-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="42176-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="42176-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="42176-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-275">- BindingEvents</span></span><br><span data-ttu-id="42176-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-276">
        - CompressedFile</span></span><br><span data-ttu-id="42176-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-277">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-278">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-278">
        - File</span></span><br><span data-ttu-id="42176-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-279">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-280">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-282">
        - PdfFile</span></span><br><span data-ttu-id="42176-283">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-283">
        - Selection</span></span><br><span data-ttu-id="42176-284">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-284">
        - Settings</span></span><br><span data-ttu-id="42176-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-285">
        - TableBindings</span></span><br><span data-ttu-id="42176-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-286">
        - TableCoercion</span></span><br><span data-ttu-id="42176-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-287">
        - TextBindings</span></span><br><span data-ttu-id="42176-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-289">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-289">Office 2019 for Mac</span></span><br><span data-ttu-id="42176-290">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="42176-291">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-291">- TaskPane</span></span><br><span data-ttu-id="42176-292">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-292">
        - Content</span></span><br><span data-ttu-id="42176-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="42176-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="42176-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="42176-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="42176-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="42176-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="42176-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="42176-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="42176-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="42176-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="42176-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-303">- BindingEvents</span></span><br><span data-ttu-id="42176-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-304">
        - CompressedFile</span></span><br><span data-ttu-id="42176-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-305">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-306">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-306">
        - File</span></span><br><span data-ttu-id="42176-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-307">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-308">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-310">
        - PdfFile</span></span><br><span data-ttu-id="42176-311">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-311">
        - Selection</span></span><br><span data-ttu-id="42176-312">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-312">
        - Settings</span></span><br><span data-ttu-id="42176-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-313">
        - TableBindings</span></span><br><span data-ttu-id="42176-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-314">
        - TableCoercion</span></span><br><span data-ttu-id="42176-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-315">
        - TextBindings</span></span><br><span data-ttu-id="42176-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-317">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-317">Office 2016 for Mac</span></span><br><span data-ttu-id="42176-318">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="42176-319">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-319">- TaskPane</span></span><br><span data-ttu-id="42176-320">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-320">
        - Content</span></span></td>
    <td><span data-ttu-id="42176-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="42176-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="42176-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-323">- BindingEvents</span></span><br><span data-ttu-id="42176-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-324">
        - CompressedFile</span></span><br><span data-ttu-id="42176-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-325">
        - DocumentEvents</span></span><br><span data-ttu-id="42176-326">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-326">
        - File</span></span><br><span data-ttu-id="42176-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-327">
        - ImageCoercion</span></span><br><span data-ttu-id="42176-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-328">
        - MatrixBindings</span></span><br><span data-ttu-id="42176-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="42176-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-330">
        - PdfFile</span></span><br><span data-ttu-id="42176-331">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-331">
        - Selection</span></span><br><span data-ttu-id="42176-332">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-332">
        - Settings</span></span><br><span data-ttu-id="42176-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-333">
        - TableBindings</span></span><br><span data-ttu-id="42176-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-334">
        - TableCoercion</span></span><br><span data-ttu-id="42176-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-335">
        - TextBindings</span></span><br><span data-ttu-id="42176-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="42176-337">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="42176-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="42176-338">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="42176-339">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="42176-340">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="42176-341">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="42176-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-343">Office Online</span></span></td>
    <td><span data-ttu-id="42176-344">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-344">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="42176-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-346">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-346">Office apps on Windows</span></span><br><span data-ttu-id="42176-347">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="42176-348">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-348">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="42176-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-350">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="42176-350">Office for iPad</span></span><br><span data-ttu-id="42176-351">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="42176-352">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-352">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="42176-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-354">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-354">Office for Mac</span></span><br><span data-ttu-id="42176-355">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="42176-356">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="42176-356">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="42176-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="42176-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="42176-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="42176-359">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-359">Platform</span></span></th>
    <th><span data-ttu-id="42176-360">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-360">Extension points</span></span></th>
    <th><span data-ttu-id="42176-361">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="42176-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-363">Office Online</span></span></td>
    <td> <span data-ttu-id="42176-364">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-364">- Mail Read</span></span><br><span data-ttu-id="42176-365">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-365">
      - Mail Compose</span></span><br><span data-ttu-id="42176-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="42176-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="42176-374">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-375">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-375">Office apps on Windows</span></span><br><span data-ttu-id="42176-376">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-377">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-377">- Mail Read</span></span><br><span data-ttu-id="42176-378">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-378">
      - Mail Compose</span></span><br><span data-ttu-id="42176-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="42176-380">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="42176-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="42176-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="42176-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="42176-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-389">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-389">Office 2019 for Windows</span></span><br><span data-ttu-id="42176-390">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-391">- Mail Read</span></span><br><span data-ttu-id="42176-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-392">
      - Mail Compose</span></span><br><span data-ttu-id="42176-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="42176-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="42176-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="42176-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="42176-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="42176-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="42176-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-403">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-403">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="42176-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-405">- Mail Read</span></span><br><span data-ttu-id="42176-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-406">
      - Mail Compose</span></span><br><span data-ttu-id="42176-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="42176-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="42176-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="42176-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="42176-413">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-414">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-414">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="42176-415">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-416">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-416">- Mail Read</span></span><br><span data-ttu-id="42176-417">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="42176-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="42176-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="42176-422">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-423">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="42176-423">Office for iOS</span></span><br><span data-ttu-id="42176-424">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-425">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-425">- Mail Read</span></span><br><span data-ttu-id="42176-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="42176-432">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-433">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-433">Office for Mac</span></span><br><span data-ttu-id="42176-434">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-435">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-435">- Mail Read</span></span><br><span data-ttu-id="42176-436">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-436">
      - Mail Compose</span></span><br><span data-ttu-id="42176-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="42176-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-445">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-445">Office 2019 for Mac</span></span><br><span data-ttu-id="42176-446">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-447">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-447">- Mail Read</span></span><br><span data-ttu-id="42176-448">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-448">
      - Mail Compose</span></span><br><span data-ttu-id="42176-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="42176-456">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-457">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-457">Office 2016 for Mac</span></span><br><span data-ttu-id="42176-458">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-459">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-459">- Mail Read</span></span><br><span data-ttu-id="42176-460">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="42176-460">
      - Mail Compose</span></span><br><span data-ttu-id="42176-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="42176-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="42176-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="42176-468">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-469">Office para Android</span><span class="sxs-lookup"><span data-stu-id="42176-469">Office for Android</span></span><br><span data-ttu-id="42176-470">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-470">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-471">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="42176-471">- Mail Read</span></span><br><span data-ttu-id="42176-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="42176-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="42176-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="42176-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="42176-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="42176-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="42176-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="42176-478">Não disponível</span><span class="sxs-lookup"><span data-stu-id="42176-478">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="42176-479">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="42176-479">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="42176-480">Word</span><span class="sxs-lookup"><span data-stu-id="42176-480">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="42176-481">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-481">Platform</span></span></th>
    <th><span data-ttu-id="42176-482">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-482">Extension points</span></span></th>
    <th><span data-ttu-id="42176-483">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-483">API requirement sets</span></span></th>
    <th><span data-ttu-id="42176-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-485">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-485">Office Online</span></span></td>
    <td> <span data-ttu-id="42176-486">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-486">- TaskPane</span></span><br><span data-ttu-id="42176-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-492">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-492">- BindingEvents</span></span><br><span data-ttu-id="42176-493">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-493">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-494">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-494">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-495">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-495">
         - File</span></span><br><span data-ttu-id="42176-496">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-496">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-497">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-497">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-498">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-498">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-499">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-499">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-500">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-500">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-501">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-501">
         - PdfFile</span></span><br><span data-ttu-id="42176-502">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-502">
         - Selection</span></span><br><span data-ttu-id="42176-503">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-503">
         - Settings</span></span><br><span data-ttu-id="42176-504">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-504">
         - TableBindings</span></span><br><span data-ttu-id="42176-505">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-505">
         - TableCoercion</span></span><br><span data-ttu-id="42176-506">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-506">
         - TextBindings</span></span><br><span data-ttu-id="42176-507">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-507">
         - TextCoercion</span></span><br><span data-ttu-id="42176-508">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-508">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-509">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-509">Office apps on Windows</span></span><br><span data-ttu-id="42176-510">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-510">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-511">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-511">- TaskPane</span></span><br><span data-ttu-id="42176-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-517">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-517">- BindingEvents</span></span><br><span data-ttu-id="42176-518">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-518">
         - CompressedFile</span></span><br><span data-ttu-id="42176-519">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-519">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-520">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-520">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-521">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-521">
         - File</span></span><br><span data-ttu-id="42176-522">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-522">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-523">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-523">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-524">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-524">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-525">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-525">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-526">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-526">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-527">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-527">
         - PdfFile</span></span><br><span data-ttu-id="42176-528">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-528">
         - Selection</span></span><br><span data-ttu-id="42176-529">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-529">
         - Settings</span></span><br><span data-ttu-id="42176-530">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-530">
         - TableBindings</span></span><br><span data-ttu-id="42176-531">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-531">
         - TableCoercion</span></span><br><span data-ttu-id="42176-532">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-532">
         - TextBindings</span></span><br><span data-ttu-id="42176-533">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-533">
         - TextCoercion</span></span><br><span data-ttu-id="42176-534">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-534">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-535">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-535">Office 2019 for Windows</span></span><br><span data-ttu-id="42176-536">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-536">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-537">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-537">- TaskPane</span></span><br><span data-ttu-id="42176-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-543">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-543">- BindingEvents</span></span><br><span data-ttu-id="42176-544">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-544">
         - CompressedFile</span></span><br><span data-ttu-id="42176-545">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-545">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-546">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-546">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-547">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-547">
         - File</span></span><br><span data-ttu-id="42176-548">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-548">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-549">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-549">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-550">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-550">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-551">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-551">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-552">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-552">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-553">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-553">
         - PdfFile</span></span><br><span data-ttu-id="42176-554">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-554">
         - Selection</span></span><br><span data-ttu-id="42176-555">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-555">
         - Settings</span></span><br><span data-ttu-id="42176-556">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-556">
         - TableBindings</span></span><br><span data-ttu-id="42176-557">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-557">
         - TableCoercion</span></span><br><span data-ttu-id="42176-558">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-558">
         - TextBindings</span></span><br><span data-ttu-id="42176-559">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-559">
         - TextCoercion</span></span><br><span data-ttu-id="42176-560">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-560">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-561">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-561">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="42176-562">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-562">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-563">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-563">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="42176-566">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-566">- BindingEvents</span></span><br><span data-ttu-id="42176-567">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-567">
         - CompressedFile</span></span><br><span data-ttu-id="42176-568">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-568">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-569">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-570">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-570">
         - File</span></span><br><span data-ttu-id="42176-571">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-571">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-572">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-572">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-573">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-573">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-574">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-574">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-575">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-575">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-576">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-576">
         - PdfFile</span></span><br><span data-ttu-id="42176-577">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-577">
         - Selection</span></span><br><span data-ttu-id="42176-578">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-578">
         - Settings</span></span><br><span data-ttu-id="42176-579">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-579">
         - TableBindings</span></span><br><span data-ttu-id="42176-580">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-580">
         - TableCoercion</span></span><br><span data-ttu-id="42176-581">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-581">
         - TextBindings</span></span><br><span data-ttu-id="42176-582">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-582">
         - TextCoercion</span></span><br><span data-ttu-id="42176-583">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-583">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-584">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-584">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="42176-585">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-585">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-586">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-586">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="42176-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="42176-588">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-588">- BindingEvents</span></span><br><span data-ttu-id="42176-589">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-589">
         - CompressedFile</span></span><br><span data-ttu-id="42176-590">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-590">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-591">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-592">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-592">
         - File</span></span><br><span data-ttu-id="42176-593">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-593">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-594">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-595">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-595">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-596">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-596">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-597">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-597">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-598">
         - PdfFile</span></span><br><span data-ttu-id="42176-599">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-599">
         - Selection</span></span><br><span data-ttu-id="42176-600">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-600">
         - Settings</span></span><br><span data-ttu-id="42176-601">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-601">
         - TableBindings</span></span><br><span data-ttu-id="42176-602">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-602">
         - TableCoercion</span></span><br><span data-ttu-id="42176-603">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-603">
         - TextBindings</span></span><br><span data-ttu-id="42176-604">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-604">
         - TextCoercion</span></span><br><span data-ttu-id="42176-605">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-605">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-606">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="42176-606">Office for iPad</span></span><br><span data-ttu-id="42176-607">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-607">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-608">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-608">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="42176-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="42176-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-613">- BindingEvents</span></span><br><span data-ttu-id="42176-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-614">
         - CompressedFile</span></span><br><span data-ttu-id="42176-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-616">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-617">
         - File</span></span><br><span data-ttu-id="42176-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-619">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-619">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-620">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-623">
         - PdfFile</span></span><br><span data-ttu-id="42176-624">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-624">
         - Selection</span></span><br><span data-ttu-id="42176-625">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-625">
         - Settings</span></span><br><span data-ttu-id="42176-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-626">
         - TableBindings</span></span><br><span data-ttu-id="42176-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-627">
         - TableCoercion</span></span><br><span data-ttu-id="42176-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-628">
         - TextBindings</span></span><br><span data-ttu-id="42176-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-629">
         - TextCoercion</span></span><br><span data-ttu-id="42176-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-631">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-631">Office for Mac</span></span><br><span data-ttu-id="42176-632">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-632">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-633">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-633">- TaskPane</span></span><br><span data-ttu-id="42176-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="42176-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="42176-639">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-639">- BindingEvents</span></span><br><span data-ttu-id="42176-640">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-640">
         - CompressedFile</span></span><br><span data-ttu-id="42176-641">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-641">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-642">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-642">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-643">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-643">
         - File</span></span><br><span data-ttu-id="42176-644">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-644">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-645">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-645">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-646">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-646">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-647">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-647">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-648">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-648">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-649">
         - PdfFile</span></span><br><span data-ttu-id="42176-650">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-650">
         - Selection</span></span><br><span data-ttu-id="42176-651">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-651">
         - Settings</span></span><br><span data-ttu-id="42176-652">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-652">
         - TableBindings</span></span><br><span data-ttu-id="42176-653">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-653">
         - TableCoercion</span></span><br><span data-ttu-id="42176-654">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-654">
         - TextBindings</span></span><br><span data-ttu-id="42176-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-655">
         - TextCoercion</span></span><br><span data-ttu-id="42176-656">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-656">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-657">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-657">Office 2019 for Mac</span></span><br><span data-ttu-id="42176-658">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-658">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-659">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-659">- TaskPane</span></span><br><span data-ttu-id="42176-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="42176-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="42176-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="42176-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="42176-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="42176-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="42176-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-665">- BindingEvents</span></span><br><span data-ttu-id="42176-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-666">
         - CompressedFile</span></span><br><span data-ttu-id="42176-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-668">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-669">
         - File</span></span><br><span data-ttu-id="42176-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-671">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-671">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-672">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-672">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-673">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-673">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-674">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-674">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-675">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-675">
         - PdfFile</span></span><br><span data-ttu-id="42176-676">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-676">
         - Selection</span></span><br><span data-ttu-id="42176-677">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-677">
         - Settings</span></span><br><span data-ttu-id="42176-678">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-678">
         - TableBindings</span></span><br><span data-ttu-id="42176-679">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-679">
         - TableCoercion</span></span><br><span data-ttu-id="42176-680">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-680">
         - TextBindings</span></span><br><span data-ttu-id="42176-681">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-681">
         - TextCoercion</span></span><br><span data-ttu-id="42176-682">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-682">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-683">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-683">Office 2016 for Mac</span></span><br><span data-ttu-id="42176-684">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-684">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-685">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="42176-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="42176-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="42176-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="42176-688">- BindingEvents</span></span><br><span data-ttu-id="42176-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-689">
         - CompressedFile</span></span><br><span data-ttu-id="42176-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="42176-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="42176-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-691">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-692">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-692">
         - File</span></span><br><span data-ttu-id="42176-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-694">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-694">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-695">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="42176-695">
         - MatrixBindings</span></span><br><span data-ttu-id="42176-696">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-696">
         - MatrixCoercion</span></span><br><span data-ttu-id="42176-697">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-697">
         - OoxmlCoercion</span></span><br><span data-ttu-id="42176-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-698">
         - PdfFile</span></span><br><span data-ttu-id="42176-699">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-699">
         - Selection</span></span><br><span data-ttu-id="42176-700">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-700">
         - Settings</span></span><br><span data-ttu-id="42176-701">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="42176-701">
         - TableBindings</span></span><br><span data-ttu-id="42176-702">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-702">
         - TableCoercion</span></span><br><span data-ttu-id="42176-703">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="42176-703">
         - TextBindings</span></span><br><span data-ttu-id="42176-704">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-704">
         - TextCoercion</span></span><br><span data-ttu-id="42176-705">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="42176-705">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="42176-706">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="42176-706">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="42176-707">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="42176-707">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="42176-708">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-708">Platform</span></span></th>
    <th><span data-ttu-id="42176-709">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-709">Extension points</span></span></th>
    <th><span data-ttu-id="42176-710">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-710">API requirement sets</span></span></th>
    <th><span data-ttu-id="42176-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-712">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-712">Office Online</span></span></td>
    <td> <span data-ttu-id="42176-713">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-713">- Content</span></span><br><span data-ttu-id="42176-714">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-714">
         - TaskPane</span></span><br><span data-ttu-id="42176-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-717">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-717">- ActiveView</span></span><br><span data-ttu-id="42176-718">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-718">
         - CompressedFile</span></span><br><span data-ttu-id="42176-719">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-719">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-720">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-720">
         - File</span></span><br><span data-ttu-id="42176-721">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-721">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-722">
         - PdfFile</span></span><br><span data-ttu-id="42176-723">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-723">
         - Selection</span></span><br><span data-ttu-id="42176-724">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-724">
         - Settings</span></span><br><span data-ttu-id="42176-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-725">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-726">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-726">Office apps on Windows</span></span><br><span data-ttu-id="42176-727">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-727">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-728">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-728">- Content</span></span><br><span data-ttu-id="42176-729">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-729">
         - TaskPane</span></span><br><span data-ttu-id="42176-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-732">- ActiveView</span></span><br><span data-ttu-id="42176-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-733">
         - CompressedFile</span></span><br><span data-ttu-id="42176-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-734">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-735">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-735">
         - File</span></span><br><span data-ttu-id="42176-736">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-736">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-737">
         - PdfFile</span></span><br><span data-ttu-id="42176-738">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-738">
         - Selection</span></span><br><span data-ttu-id="42176-739">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-739">
         - Settings</span></span><br><span data-ttu-id="42176-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-740">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-741">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-741">Office 2019 for Windows</span></span><br><span data-ttu-id="42176-742">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-742">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-743">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-743">- Content</span></span><br><span data-ttu-id="42176-744">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-744">
         - TaskPane</span></span><br><span data-ttu-id="42176-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-747">- ActiveView</span></span><br><span data-ttu-id="42176-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-748">
         - CompressedFile</span></span><br><span data-ttu-id="42176-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-749">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-750">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-750">
         - File</span></span><br><span data-ttu-id="42176-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-751">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-752">
         - PdfFile</span></span><br><span data-ttu-id="42176-753">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-753">
         - Selection</span></span><br><span data-ttu-id="42176-754">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-754">
         - Settings</span></span><br><span data-ttu-id="42176-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-756">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-756">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="42176-757">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-757">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-758">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-758">- Content</span></span><br><span data-ttu-id="42176-759">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-759">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="42176-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="42176-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-761">- ActiveView</span></span><br><span data-ttu-id="42176-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-762">
         - CompressedFile</span></span><br><span data-ttu-id="42176-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-763">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-764">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-764">
         - File</span></span><br><span data-ttu-id="42176-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-765">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-766">
         - PdfFile</span></span><br><span data-ttu-id="42176-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-767">
         - Selection</span></span><br><span data-ttu-id="42176-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-768">
         - Settings</span></span><br><span data-ttu-id="42176-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-770">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-770">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="42176-771">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-772">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-772">- Content</span></span><br><span data-ttu-id="42176-773">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-773">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="42176-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="42176-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="42176-775">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-775">- ActiveView</span></span><br><span data-ttu-id="42176-776">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-776">
         - CompressedFile</span></span><br><span data-ttu-id="42176-777">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-777">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-778">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-778">
         - File</span></span><br><span data-ttu-id="42176-779">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-779">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-780">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-780">
         - PdfFile</span></span><br><span data-ttu-id="42176-781">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-781">
         - Selection</span></span><br><span data-ttu-id="42176-782">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-782">
         - Settings</span></span><br><span data-ttu-id="42176-783">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-783">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-784">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="42176-784">Office for iPad</span></span><br><span data-ttu-id="42176-785">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-785">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-786">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-786">- Content</span></span><br><span data-ttu-id="42176-787">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-787">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-789">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-789">- ActiveView</span></span><br><span data-ttu-id="42176-790">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-790">
         - CompressedFile</span></span><br><span data-ttu-id="42176-791">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-791">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-792">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-792">
         - File</span></span><br><span data-ttu-id="42176-793">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-793">
         - PdfFile</span></span><br><span data-ttu-id="42176-794">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-794">
         - Selection</span></span><br><span data-ttu-id="42176-795">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-795">
         - Settings</span></span><br><span data-ttu-id="42176-796">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-796">
         - TextCoercion</span></span><br><span data-ttu-id="42176-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-797">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-798">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-798">Office for Mac</span></span><br><span data-ttu-id="42176-799">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="42176-799">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="42176-800">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-800">- Content</span></span><br><span data-ttu-id="42176-801">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-801">
         - TaskPane</span></span><br><span data-ttu-id="42176-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-804">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-804">- ActiveView</span></span><br><span data-ttu-id="42176-805">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-805">
         - CompressedFile</span></span><br><span data-ttu-id="42176-806">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-806">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-807">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-807">
         - File</span></span><br><span data-ttu-id="42176-808">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-808">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-809">
         - PdfFile</span></span><br><span data-ttu-id="42176-810">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-810">
         - Selection</span></span><br><span data-ttu-id="42176-811">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-811">
         - Settings</span></span><br><span data-ttu-id="42176-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-813">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-813">Office 2019 for Mac</span></span><br><span data-ttu-id="42176-814">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-814">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-815">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-815">- Content</span></span><br><span data-ttu-id="42176-816">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-816">
         - TaskPane</span></span><br><span data-ttu-id="42176-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-819">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-819">- ActiveView</span></span><br><span data-ttu-id="42176-820">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-820">
         - CompressedFile</span></span><br><span data-ttu-id="42176-821">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-821">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-822">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-822">
         - File</span></span><br><span data-ttu-id="42176-823">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-823">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-824">
         - PdfFile</span></span><br><span data-ttu-id="42176-825">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-825">
         - Selection</span></span><br><span data-ttu-id="42176-826">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-826">
         - Settings</span></span><br><span data-ttu-id="42176-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-828">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-828">Office 2016 for Mac</span></span><br><span data-ttu-id="42176-829">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-829">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-830">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-830">- Content</span></span><br><span data-ttu-id="42176-831">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-831">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="42176-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="42176-833">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="42176-833">- ActiveView</span></span><br><span data-ttu-id="42176-834">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="42176-834">
         - CompressedFile</span></span><br><span data-ttu-id="42176-835">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-835">
         - DocumentEvents</span></span><br><span data-ttu-id="42176-836">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="42176-836">
         - File</span></span><br><span data-ttu-id="42176-837">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-837">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-838">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="42176-838">
         - PdfFile</span></span><br><span data-ttu-id="42176-839">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-839">
         - Selection</span></span><br><span data-ttu-id="42176-840">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-840">
         - Settings</span></span><br><span data-ttu-id="42176-841">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-841">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="42176-842">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="42176-842">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="42176-843">OneNote</span><span class="sxs-lookup"><span data-stu-id="42176-843">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="42176-844">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-844">Platform</span></span></th>
    <th><span data-ttu-id="42176-845">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-845">Extension points</span></span></th>
    <th><span data-ttu-id="42176-846">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-846">API requirement sets</span></span></th>
    <th><span data-ttu-id="42176-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-848">Office Online</span><span class="sxs-lookup"><span data-stu-id="42176-848">Office Online</span></span></td>
    <td> <span data-ttu-id="42176-849">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="42176-849">- Content</span></span><br><span data-ttu-id="42176-850">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-850">
         - TaskPane</span></span><br><span data-ttu-id="42176-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="42176-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="42176-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="42176-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-854">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="42176-854">- DocumentEvents</span></span><br><span data-ttu-id="42176-855">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-855">
         - HtmlCoercion</span></span><br><span data-ttu-id="42176-856">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-856">
         - ImageCoercion</span></span><br><span data-ttu-id="42176-857">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="42176-857">
         - Settings</span></span><br><span data-ttu-id="42176-858">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-858">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="42176-859">Project</span><span class="sxs-lookup"><span data-stu-id="42176-859">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="42176-860">Plataforma</span><span class="sxs-lookup"><span data-stu-id="42176-860">Platform</span></span></th>
    <th><span data-ttu-id="42176-861">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="42176-861">Extension points</span></span></th>
    <th><span data-ttu-id="42176-862">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="42176-862">API requirement sets</span></span></th>
    <th><span data-ttu-id="42176-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="42176-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-864">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-864">Office 2019 for Windows</span></span><br><span data-ttu-id="42176-865">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-866">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-866">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-868">- Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-868">- Selection</span></span><br><span data-ttu-id="42176-869">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-869">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-870">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-870">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="42176-871">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-871">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-872">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-872">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-874">- Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-874">- Selection</span></span><br><span data-ttu-id="42176-875">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-875">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="42176-876">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="42176-876">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="42176-877">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="42176-877">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="42176-878">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="42176-878">- TaskPane</span></span></td>
    <td> <span data-ttu-id="42176-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="42176-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="42176-880">- Seleção</span><span class="sxs-lookup"><span data-stu-id="42176-880">- Selection</span></span><br><span data-ttu-id="42176-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="42176-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="42176-882">Confira também</span><span class="sxs-lookup"><span data-stu-id="42176-882">See also</span></span>

- [<span data-ttu-id="42176-883">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="42176-883">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="42176-884">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="42176-884">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="42176-885">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="42176-885">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="42176-886">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="42176-886">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="42176-887">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="42176-887">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="42176-888">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="42176-888">Update history for Office 365 ProPlus releases</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="42176-889">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="42176-889">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="42176-890">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="42176-890">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="42176-891">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="42176-891">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="42176-892">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="42176-892">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="42176-893">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="42176-893">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
