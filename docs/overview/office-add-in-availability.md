---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint, OneNote e Project.
ms.date: 05/23/2019
localization_priority: Priority
ms.openlocfilehash: 6fb1f0db839910e91d7a5215f8e21f5b33ff2165
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432191"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="12f17-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="12f17-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="12f17-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="12f17-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="12f17-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="12f17-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="12f17-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="12f17-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="12f17-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="12f17-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="12f17-108">Excel</span><span class="sxs-lookup"><span data-stu-id="12f17-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="12f17-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="12f17-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="12f17-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="12f17-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-113">Office Online</span></span></td>
    <td> <span data-ttu-id="12f17-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-114">- TaskPane</span></span><br><span data-ttu-id="12f17-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-115">
        - Content</span></span><br><span data-ttu-id="12f17-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-116">
        - Custom Functions</span></span><br><span data-ttu-id="12f17-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="12f17-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12f17-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12f17-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12f17-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-128">
        - BindingEvents</span></span><br><span data-ttu-id="12f17-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-129">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-130">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-131">
        - File</span></span><br><span data-ttu-id="12f17-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-132">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-134">
        - Selection</span></span><br><span data-ttu-id="12f17-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-135">
        - Settings</span></span><br><span data-ttu-id="12f17-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-136">
        - TableBindings</span></span><br><span data-ttu-id="12f17-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-137">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-138">
        - TextBindings</span></span><br><span data-ttu-id="12f17-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-140">Office on Windows</span></span><br><span data-ttu-id="12f17-141">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-142">- TaskPane</span></span><br><span data-ttu-id="12f17-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-143">
        - Content</span></span><br><span data-ttu-id="12f17-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-144">
        - Custom Functions</span></span><br><span data-ttu-id="12f17-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="12f17-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="12f17-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12f17-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12f17-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-156">
        - BindingEvents</span></span><br><span data-ttu-id="12f17-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-157">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-158">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-159">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-159">
        - File</span></span><br><span data-ttu-id="12f17-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-160">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-162">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-162">
        - Selection</span></span><br><span data-ttu-id="12f17-163">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-163">
        - Settings</span></span><br><span data-ttu-id="12f17-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-164">
        - TableBindings</span></span><br><span data-ttu-id="12f17-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-165">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-166">
        - TextBindings</span></span><br><span data-ttu-id="12f17-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-168">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-168">Office 2019 on Windows</span></span><br><span data-ttu-id="12f17-169">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12f17-170">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-170">- TaskPane</span></span><br><span data-ttu-id="12f17-171">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-171">
        - Content</span></span><br><span data-ttu-id="12f17-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12f17-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-182">- BindingEvents</span></span><br><span data-ttu-id="12f17-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-183">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-184">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-185">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-185">
        - File</span></span><br><span data-ttu-id="12f17-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-186">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-187">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-189">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-189">
        - Selection</span></span><br><span data-ttu-id="12f17-190">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-190">
        - Settings</span></span><br><span data-ttu-id="12f17-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-191">
        - TableBindings</span></span><br><span data-ttu-id="12f17-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-192">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-193">
        - TextBindings</span></span><br><span data-ttu-id="12f17-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-195">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-195">Office 2016 on Windows</span></span><br><span data-ttu-id="12f17-196">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12f17-197">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-197">- TaskPane</span></span><br><span data-ttu-id="12f17-198">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-198">
        - Content</span></span></td>
    <td><span data-ttu-id="12f17-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="12f17-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-201">- BindingEvents</span></span><br><span data-ttu-id="12f17-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-202">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-203">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-204">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-204">
        - File</span></span><br><span data-ttu-id="12f17-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-205">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-206">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-208">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-208">
        - Selection</span></span><br><span data-ttu-id="12f17-209">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-209">
        - Settings</span></span><br><span data-ttu-id="12f17-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-210">
        - TableBindings</span></span><br><span data-ttu-id="12f17-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-211">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-212">
        - TextBindings</span></span><br><span data-ttu-id="12f17-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-214">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-214">Office 2013 on Windows</span></span><br><span data-ttu-id="12f17-215">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12f17-216">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-216">
        - TaskPane</span></span><br><span data-ttu-id="12f17-217">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="12f17-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12f17-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="12f17-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-219">
        - BindingEvents</span></span><br><span data-ttu-id="12f17-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-220">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-221">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-222">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-222">
        - File</span></span><br><span data-ttu-id="12f17-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-223">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-224">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-226">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-226">
        - Selection</span></span><br><span data-ttu-id="12f17-227">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-227">
        - Settings</span></span><br><span data-ttu-id="12f17-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-228">
        - TableBindings</span></span><br><span data-ttu-id="12f17-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-229">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-230">
        - TextBindings</span></span><br><span data-ttu-id="12f17-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-232">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="12f17-232">Office for iPad</span></span><br><span data-ttu-id="12f17-233">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12f17-234">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-234">- TaskPane</span></span><br><span data-ttu-id="12f17-235">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-235">
        - Content</span></span><br><span data-ttu-id="12f17-236">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12f17-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12f17-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12f17-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-247">- BindingEvents</span></span><br><span data-ttu-id="12f17-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-248">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-249">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-249">
        - File</span></span><br><span data-ttu-id="12f17-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-250">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-251">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-253">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-253">
        - Selection</span></span><br><span data-ttu-id="12f17-254">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-254">
        - Settings</span></span><br><span data-ttu-id="12f17-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-255">
        - TableBindings</span></span><br><span data-ttu-id="12f17-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-256">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-257">
        - TextBindings</span></span><br><span data-ttu-id="12f17-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-259">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-259">Office for Mac</span></span><br><span data-ttu-id="12f17-260">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12f17-261">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-261">- TaskPane</span></span><br><span data-ttu-id="12f17-262">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-262">
        - Content</span></span><br><span data-ttu-id="12f17-263">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-263">
        - Custom Functions</span></span><br><span data-ttu-id="12f17-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12f17-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="12f17-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="12f17-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-275">- BindingEvents</span></span><br><span data-ttu-id="12f17-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-276">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-277">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-278">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-278">
        - File</span></span><br><span data-ttu-id="12f17-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-279">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-280">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-282">
        - PdfFile</span></span><br><span data-ttu-id="12f17-283">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-283">
        - Selection</span></span><br><span data-ttu-id="12f17-284">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-284">
        - Settings</span></span><br><span data-ttu-id="12f17-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-285">
        - TableBindings</span></span><br><span data-ttu-id="12f17-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-286">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-287">
        - TextBindings</span></span><br><span data-ttu-id="12f17-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-289">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-289">Office 2019 for Mac</span></span><br><span data-ttu-id="12f17-290">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12f17-291">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-291">- TaskPane</span></span><br><span data-ttu-id="12f17-292">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-292">
        - Content</span></span><br><span data-ttu-id="12f17-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="12f17-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="12f17-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="12f17-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="12f17-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="12f17-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="12f17-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="12f17-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="12f17-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="12f17-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="12f17-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-303">- BindingEvents</span></span><br><span data-ttu-id="12f17-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-304">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-305">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-306">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-306">
        - File</span></span><br><span data-ttu-id="12f17-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-307">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-308">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-310">
        - PdfFile</span></span><br><span data-ttu-id="12f17-311">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-311">
        - Selection</span></span><br><span data-ttu-id="12f17-312">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-312">
        - Settings</span></span><br><span data-ttu-id="12f17-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-313">
        - TableBindings</span></span><br><span data-ttu-id="12f17-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-314">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-315">
        - TextBindings</span></span><br><span data-ttu-id="12f17-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-317">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-317">Office 2016 for Mac</span></span><br><span data-ttu-id="12f17-318">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="12f17-319">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-319">- TaskPane</span></span><br><span data-ttu-id="12f17-320">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-320">
        - Content</span></span></td>
    <td><span data-ttu-id="12f17-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="12f17-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="12f17-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-323">- BindingEvents</span></span><br><span data-ttu-id="12f17-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-324">
        - CompressedFile</span></span><br><span data-ttu-id="12f17-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-325">
        - DocumentEvents</span></span><br><span data-ttu-id="12f17-326">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-326">
        - File</span></span><br><span data-ttu-id="12f17-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-327">
        - ImageCoercion</span></span><br><span data-ttu-id="12f17-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-328">
        - MatrixBindings</span></span><br><span data-ttu-id="12f17-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="12f17-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-330">
        - PdfFile</span></span><br><span data-ttu-id="12f17-331">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-331">
        - Selection</span></span><br><span data-ttu-id="12f17-332">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-332">
        - Settings</span></span><br><span data-ttu-id="12f17-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-333">
        - TableBindings</span></span><br><span data-ttu-id="12f17-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-334">
        - TableCoercion</span></span><br><span data-ttu-id="12f17-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-335">
        - TextBindings</span></span><br><span data-ttu-id="12f17-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12f17-337">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="12f17-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="12f17-338">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="12f17-339">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="12f17-340">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="12f17-341">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="12f17-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-343">Office Online</span></span></td>
    <td><span data-ttu-id="12f17-344">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12f17-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-346">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-346">Office on Windows</span></span><br><span data-ttu-id="12f17-347">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12f17-348">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12f17-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-350">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="12f17-350">Office for iPad</span></span><br><span data-ttu-id="12f17-351">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12f17-352">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12f17-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-354">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-354">Office for Mac</span></span><br><span data-ttu-id="12f17-355">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="12f17-356">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="12f17-356">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="12f17-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="12f17-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="12f17-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12f17-359">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-359">Platform</span></span></th>
    <th><span data-ttu-id="12f17-360">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-360">Extension points</span></span></th>
    <th><span data-ttu-id="12f17-361">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="12f17-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-363">Office Online</span></span></td>
    <td> <span data-ttu-id="12f17-364">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-364">- Mail Read</span></span><br><span data-ttu-id="12f17-365">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-365">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12f17-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12f17-374">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-375">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-375">Office on Windows</span></span><br><span data-ttu-id="12f17-376">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-377">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-377">- Mail Read</span></span><br><span data-ttu-id="12f17-378">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-378">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12f17-380">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="12f17-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12f17-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12f17-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12f17-388">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-389">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-389">Office 2019 on Windows</span></span><br><span data-ttu-id="12f17-390">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-391">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-391">- Mail Read</span></span><br><span data-ttu-id="12f17-392">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-392">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12f17-394">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="12f17-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12f17-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12f17-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12f17-402">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-403">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-403">Office 2016 on Windows</span></span><br><span data-ttu-id="12f17-404">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-405">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-405">- Mail Read</span></span><br><span data-ttu-id="12f17-406">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-406">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="12f17-408">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="12f17-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="12f17-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12f17-413">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-414">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-414">Office 2013 on Windows</span></span><br><span data-ttu-id="12f17-415">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-416">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-416">- Mail Read</span></span><br><span data-ttu-id="12f17-417">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="12f17-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="12f17-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="12f17-422">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-423">Office para iOS</span><span class="sxs-lookup"><span data-stu-id="12f17-423">Office for iOS</span></span><br><span data-ttu-id="12f17-424">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-425">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-425">- Mail Read</span></span><br><span data-ttu-id="12f17-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12f17-432">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-433">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-433">Office for Mac</span></span><br><span data-ttu-id="12f17-434">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-435">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-435">- Mail Read</span></span><br><span data-ttu-id="12f17-436">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-436">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="12f17-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="12f17-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="12f17-445">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-446">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-446">Office 2019 for Mac</span></span><br><span data-ttu-id="12f17-447">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-448">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-448">- Mail Read</span></span><br><span data-ttu-id="12f17-449">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-449">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12f17-457">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-458">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-458">Office 2016 for Mac</span></span><br><span data-ttu-id="12f17-459">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-460">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-460">- Mail Read</span></span><br><span data-ttu-id="12f17-461">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="12f17-461">
      - Mail Compose</span></span><br><span data-ttu-id="12f17-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="12f17-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="12f17-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="12f17-469">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-470">Office para Android</span><span class="sxs-lookup"><span data-stu-id="12f17-470">Office for Android</span></span><br><span data-ttu-id="12f17-471">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-471">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-472">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="12f17-472">- Mail Read</span></span><br><span data-ttu-id="12f17-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="12f17-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="12f17-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="12f17-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="12f17-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="12f17-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="12f17-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="12f17-479">Não disponível</span><span class="sxs-lookup"><span data-stu-id="12f17-479">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="12f17-480">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="12f17-480">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="12f17-481">Word</span><span class="sxs-lookup"><span data-stu-id="12f17-481">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12f17-482">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-482">Platform</span></span></th>
    <th><span data-ttu-id="12f17-483">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-483">Extension points</span></span></th>
    <th><span data-ttu-id="12f17-484">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-484">API requirement sets</span></span></th>
    <th><span data-ttu-id="12f17-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-486">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-486">Office Online</span></span></td>
    <td> <span data-ttu-id="12f17-487">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-487">- TaskPane</span></span><br><span data-ttu-id="12f17-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-493">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-493">- BindingEvents</span></span><br><span data-ttu-id="12f17-494">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-494">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-495">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-495">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-496">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-496">
         - File</span></span><br><span data-ttu-id="12f17-497">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-497">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-498">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-498">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-499">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-499">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-500">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-500">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-501">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-501">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-502">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-502">
         - PdfFile</span></span><br><span data-ttu-id="12f17-503">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-503">
         - Selection</span></span><br><span data-ttu-id="12f17-504">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-504">
         - Settings</span></span><br><span data-ttu-id="12f17-505">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-505">
         - TableBindings</span></span><br><span data-ttu-id="12f17-506">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-506">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-507">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-507">
         - TextBindings</span></span><br><span data-ttu-id="12f17-508">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-508">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-509">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-509">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-510">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-510">Office on Windows</span></span><br><span data-ttu-id="12f17-511">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-511">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-512">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-512">- TaskPane</span></span><br><span data-ttu-id="12f17-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-518">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-518">- BindingEvents</span></span><br><span data-ttu-id="12f17-519">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-519">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-520">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-520">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-521">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-521">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-522">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-522">
         - File</span></span><br><span data-ttu-id="12f17-523">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-523">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-524">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-524">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-525">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-525">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-526">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-526">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-527">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-527">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-528">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-528">
         - PdfFile</span></span><br><span data-ttu-id="12f17-529">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-529">
         - Selection</span></span><br><span data-ttu-id="12f17-530">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-530">
         - Settings</span></span><br><span data-ttu-id="12f17-531">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-531">
         - TableBindings</span></span><br><span data-ttu-id="12f17-532">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-532">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-533">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-533">
         - TextBindings</span></span><br><span data-ttu-id="12f17-534">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-534">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-535">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-535">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-536">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-536">Office 2019 on Windows</span></span><br><span data-ttu-id="12f17-537">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-537">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-538">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-538">- TaskPane</span></span><br><span data-ttu-id="12f17-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-544">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-544">- BindingEvents</span></span><br><span data-ttu-id="12f17-545">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-545">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-546">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-546">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-547">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-547">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-548">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-548">
         - File</span></span><br><span data-ttu-id="12f17-549">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-549">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-550">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-550">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-551">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-551">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-552">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-552">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-553">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-553">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-554">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-554">
         - PdfFile</span></span><br><span data-ttu-id="12f17-555">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-555">
         - Selection</span></span><br><span data-ttu-id="12f17-556">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-556">
         - Settings</span></span><br><span data-ttu-id="12f17-557">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-557">
         - TableBindings</span></span><br><span data-ttu-id="12f17-558">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-558">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-559">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-559">
         - TextBindings</span></span><br><span data-ttu-id="12f17-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-560">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-561">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-561">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-562">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-562">Office 2016 on Windows</span></span><br><span data-ttu-id="12f17-563">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-563">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-564">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-564">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="12f17-567">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-567">- BindingEvents</span></span><br><span data-ttu-id="12f17-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-568">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-569">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-569">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-570">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-570">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-571">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-571">
         - File</span></span><br><span data-ttu-id="12f17-572">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-572">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-573">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-574">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-577">
         - PdfFile</span></span><br><span data-ttu-id="12f17-578">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-578">
         - Selection</span></span><br><span data-ttu-id="12f17-579">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-579">
         - Settings</span></span><br><span data-ttu-id="12f17-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-580">
         - TableBindings</span></span><br><span data-ttu-id="12f17-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-581">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-582">
         - TextBindings</span></span><br><span data-ttu-id="12f17-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-583">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-585">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-585">Office 2013 on Windows</span></span><br><span data-ttu-id="12f17-586">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-587">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12f17-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12f17-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-589">- BindingEvents</span></span><br><span data-ttu-id="12f17-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-590">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-592">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-593">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-593">
         - File</span></span><br><span data-ttu-id="12f17-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-595">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-596">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-599">
         - PdfFile</span></span><br><span data-ttu-id="12f17-600">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-600">
         - Selection</span></span><br><span data-ttu-id="12f17-601">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-601">
         - Settings</span></span><br><span data-ttu-id="12f17-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-602">
         - TableBindings</span></span><br><span data-ttu-id="12f17-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-603">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-604">
         - TextBindings</span></span><br><span data-ttu-id="12f17-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-605">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-606">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-607">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="12f17-607">Office for iPad</span></span><br><span data-ttu-id="12f17-608">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-608">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-609">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12f17-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12f17-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-614">- BindingEvents</span></span><br><span data-ttu-id="12f17-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-615">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-617">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-618">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-618">
         - File</span></span><br><span data-ttu-id="12f17-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-620">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-621">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-624">
         - PdfFile</span></span><br><span data-ttu-id="12f17-625">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-625">
         - Selection</span></span><br><span data-ttu-id="12f17-626">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-626">
         - Settings</span></span><br><span data-ttu-id="12f17-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-627">
         - TableBindings</span></span><br><span data-ttu-id="12f17-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-628">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-629">
         - TextBindings</span></span><br><span data-ttu-id="12f17-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-630">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-632">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-632">Office for Mac</span></span><br><span data-ttu-id="12f17-633">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-633">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-634">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-634">- TaskPane</span></span><br><span data-ttu-id="12f17-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12f17-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12f17-640">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-640">- BindingEvents</span></span><br><span data-ttu-id="12f17-641">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-641">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-642">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-642">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-643">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-643">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-644">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-644">
         - File</span></span><br><span data-ttu-id="12f17-645">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-645">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-646">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-646">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-647">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-647">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-648">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-648">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-649">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-649">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-650">
         - PdfFile</span></span><br><span data-ttu-id="12f17-651">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-651">
         - Selection</span></span><br><span data-ttu-id="12f17-652">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-652">
         - Settings</span></span><br><span data-ttu-id="12f17-653">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-653">
         - TableBindings</span></span><br><span data-ttu-id="12f17-654">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-654">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-655">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-655">
         - TextBindings</span></span><br><span data-ttu-id="12f17-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-656">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-657">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-657">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-658">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-658">Office 2019 for Mac</span></span><br><span data-ttu-id="12f17-659">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-659">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-660">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-660">- TaskPane</span></span><br><span data-ttu-id="12f17-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="12f17-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="12f17-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="12f17-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="12f17-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="12f17-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="12f17-666">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-666">- BindingEvents</span></span><br><span data-ttu-id="12f17-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-667">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-668">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-668">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-669">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-669">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-670">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-670">
         - File</span></span><br><span data-ttu-id="12f17-671">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-671">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-672">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-672">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-673">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-673">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-674">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-674">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-675">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-675">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-676">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-676">
         - PdfFile</span></span><br><span data-ttu-id="12f17-677">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-677">
         - Selection</span></span><br><span data-ttu-id="12f17-678">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-678">
         - Settings</span></span><br><span data-ttu-id="12f17-679">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-679">
         - TableBindings</span></span><br><span data-ttu-id="12f17-680">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-680">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-681">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-681">
         - TextBindings</span></span><br><span data-ttu-id="12f17-682">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-682">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-683">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-683">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-684">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-684">Office 2016 for Mac</span></span><br><span data-ttu-id="12f17-685">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-685">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-686">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="12f17-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="12f17-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="12f17-689">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-689">- BindingEvents</span></span><br><span data-ttu-id="12f17-690">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-690">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-691">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="12f17-691">
         - CustomXmlParts</span></span><br><span data-ttu-id="12f17-692">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-692">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-693">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-693">
         - File</span></span><br><span data-ttu-id="12f17-694">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-694">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-695">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-695">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-696">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-696">
         - MatrixBindings</span></span><br><span data-ttu-id="12f17-697">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-697">
         - MatrixCoercion</span></span><br><span data-ttu-id="12f17-698">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-698">
         - OoxmlCoercion</span></span><br><span data-ttu-id="12f17-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-699">
         - PdfFile</span></span><br><span data-ttu-id="12f17-700">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-700">
         - Selection</span></span><br><span data-ttu-id="12f17-701">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-701">
         - Settings</span></span><br><span data-ttu-id="12f17-702">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-702">
         - TableBindings</span></span><br><span data-ttu-id="12f17-703">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-703">
         - TableCoercion</span></span><br><span data-ttu-id="12f17-704">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="12f17-704">
         - TextBindings</span></span><br><span data-ttu-id="12f17-705">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-705">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-706">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="12f17-706">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="12f17-707">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="12f17-707">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="12f17-708">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="12f17-708">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12f17-709">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-709">Platform</span></span></th>
    <th><span data-ttu-id="12f17-710">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-710">Extension points</span></span></th>
    <th><span data-ttu-id="12f17-711">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-711">API requirement sets</span></span></th>
    <th><span data-ttu-id="12f17-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-713">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-713">Office Online</span></span></td>
    <td> <span data-ttu-id="12f17-714">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-714">- Content</span></span><br><span data-ttu-id="12f17-715">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-715">
         - TaskPane</span></span><br><span data-ttu-id="12f17-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-718">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-718">- ActiveView</span></span><br><span data-ttu-id="12f17-719">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-719">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-720">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-720">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-721">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-721">
         - File</span></span><br><span data-ttu-id="12f17-722">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-722">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-723">
         - PdfFile</span></span><br><span data-ttu-id="12f17-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-724">
         - Selection</span></span><br><span data-ttu-id="12f17-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-725">
         - Settings</span></span><br><span data-ttu-id="12f17-726">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-726">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-727">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-727">Office on Windows</span></span><br><span data-ttu-id="12f17-728">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-728">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-729">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-729">- Content</span></span><br><span data-ttu-id="12f17-730">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-730">
         - TaskPane</span></span><br><span data-ttu-id="12f17-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-733">- ActiveView</span></span><br><span data-ttu-id="12f17-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-734">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-735">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-736">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-736">
         - File</span></span><br><span data-ttu-id="12f17-737">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-737">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-738">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-738">
         - PdfFile</span></span><br><span data-ttu-id="12f17-739">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-739">
         - Selection</span></span><br><span data-ttu-id="12f17-740">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-740">
         - Settings</span></span><br><span data-ttu-id="12f17-741">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-741">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-742">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-742">Office 2019 on Windows</span></span><br><span data-ttu-id="12f17-743">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-743">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-744">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-744">- Content</span></span><br><span data-ttu-id="12f17-745">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-745">
         - TaskPane</span></span><br><span data-ttu-id="12f17-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-748">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-748">- ActiveView</span></span><br><span data-ttu-id="12f17-749">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-749">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-750">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-750">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-751">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-751">
         - File</span></span><br><span data-ttu-id="12f17-752">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-752">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-753">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-753">
         - PdfFile</span></span><br><span data-ttu-id="12f17-754">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-754">
         - Selection</span></span><br><span data-ttu-id="12f17-755">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-755">
         - Settings</span></span><br><span data-ttu-id="12f17-756">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-756">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-757">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-757">Office 2016 on Windows</span></span><br><span data-ttu-id="12f17-758">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-758">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-759">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-759">- Content</span></span><br><span data-ttu-id="12f17-760">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-760">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12f17-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12f17-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-762">- ActiveView</span></span><br><span data-ttu-id="12f17-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-763">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-764">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-765">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-765">
         - File</span></span><br><span data-ttu-id="12f17-766">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-766">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-767">
         - PdfFile</span></span><br><span data-ttu-id="12f17-768">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-768">
         - Selection</span></span><br><span data-ttu-id="12f17-769">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-769">
         - Settings</span></span><br><span data-ttu-id="12f17-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-771">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-771">Office 2013 on Windows</span></span><br><span data-ttu-id="12f17-772">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-772">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-773">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-773">- Content</span></span><br><span data-ttu-id="12f17-774">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-774">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="12f17-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12f17-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12f17-776">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-776">- ActiveView</span></span><br><span data-ttu-id="12f17-777">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-777">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-778">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-778">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-779">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-779">
         - File</span></span><br><span data-ttu-id="12f17-780">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-780">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-781">
         - PdfFile</span></span><br><span data-ttu-id="12f17-782">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-782">
         - Selection</span></span><br><span data-ttu-id="12f17-783">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-783">
         - Settings</span></span><br><span data-ttu-id="12f17-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-785">Office para iPad</span><span class="sxs-lookup"><span data-stu-id="12f17-785">Office for iPad</span></span><br><span data-ttu-id="12f17-786">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-786">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-787">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-787">- Content</span></span><br><span data-ttu-id="12f17-788">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-790">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-790">- ActiveView</span></span><br><span data-ttu-id="12f17-791">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-791">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-792">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-792">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-793">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-793">
         - File</span></span><br><span data-ttu-id="12f17-794">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-794">
         - PdfFile</span></span><br><span data-ttu-id="12f17-795">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-795">
         - Selection</span></span><br><span data-ttu-id="12f17-796">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-796">
         - Settings</span></span><br><span data-ttu-id="12f17-797">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-797">
         - TextCoercion</span></span><br><span data-ttu-id="12f17-798">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-798">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-799">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-799">Office for Mac</span></span><br><span data-ttu-id="12f17-800">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="12f17-800">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="12f17-801">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-801">- Content</span></span><br><span data-ttu-id="12f17-802">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-802">
         - TaskPane</span></span><br><span data-ttu-id="12f17-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-805">- ActiveView</span></span><br><span data-ttu-id="12f17-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-806">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-807">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-808">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-808">
         - File</span></span><br><span data-ttu-id="12f17-809">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-809">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-810">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-810">
         - PdfFile</span></span><br><span data-ttu-id="12f17-811">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-811">
         - Selection</span></span><br><span data-ttu-id="12f17-812">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-812">
         - Settings</span></span><br><span data-ttu-id="12f17-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-814">Office 2019 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-814">Office 2019 for Mac</span></span><br><span data-ttu-id="12f17-815">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-815">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-816">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-816">- Content</span></span><br><span data-ttu-id="12f17-817">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-817">
         - TaskPane</span></span><br><span data-ttu-id="12f17-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-820">- ActiveView</span></span><br><span data-ttu-id="12f17-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-821">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-822">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-823">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-823">
         - File</span></span><br><span data-ttu-id="12f17-824">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-824">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-825">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-825">
         - PdfFile</span></span><br><span data-ttu-id="12f17-826">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-826">
         - Selection</span></span><br><span data-ttu-id="12f17-827">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-827">
         - Settings</span></span><br><span data-ttu-id="12f17-828">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-828">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-829">Office 2016 para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-829">Office 2016 for Mac</span></span><br><span data-ttu-id="12f17-830">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-830">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-831">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-831">- Content</span></span><br><span data-ttu-id="12f17-832">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-832">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="12f17-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="12f17-834">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="12f17-834">- ActiveView</span></span><br><span data-ttu-id="12f17-835">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="12f17-835">
         - CompressedFile</span></span><br><span data-ttu-id="12f17-836">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-836">
         - DocumentEvents</span></span><br><span data-ttu-id="12f17-837">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="12f17-837">
         - File</span></span><br><span data-ttu-id="12f17-838">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-838">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-839">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="12f17-839">
         - PdfFile</span></span><br><span data-ttu-id="12f17-840">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-840">
         - Selection</span></span><br><span data-ttu-id="12f17-841">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-841">
         - Settings</span></span><br><span data-ttu-id="12f17-842">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-842">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="12f17-843">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="12f17-843">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="12f17-844">OneNote</span><span class="sxs-lookup"><span data-stu-id="12f17-844">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12f17-845">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-845">Platform</span></span></th>
    <th><span data-ttu-id="12f17-846">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-846">Extension points</span></span></th>
    <th><span data-ttu-id="12f17-847">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-847">API requirement sets</span></span></th>
    <th><span data-ttu-id="12f17-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-849">Office Online</span><span class="sxs-lookup"><span data-stu-id="12f17-849">Office Online</span></span></td>
    <td> <span data-ttu-id="12f17-850">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="12f17-850">- Content</span></span><br><span data-ttu-id="12f17-851">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-851">
         - TaskPane</span></span><br><span data-ttu-id="12f17-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="12f17-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="12f17-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="12f17-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-855">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="12f17-855">- DocumentEvents</span></span><br><span data-ttu-id="12f17-856">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-856">
         - HtmlCoercion</span></span><br><span data-ttu-id="12f17-857">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-857">
         - ImageCoercion</span></span><br><span data-ttu-id="12f17-858">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="12f17-858">
         - Settings</span></span><br><span data-ttu-id="12f17-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-859">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="12f17-860">Project</span><span class="sxs-lookup"><span data-stu-id="12f17-860">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="12f17-861">Plataforma</span><span class="sxs-lookup"><span data-stu-id="12f17-861">Platform</span></span></th>
    <th><span data-ttu-id="12f17-862">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="12f17-862">Extension points</span></span></th>
    <th><span data-ttu-id="12f17-863">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="12f17-863">API requirement sets</span></span></th>
    <th><span data-ttu-id="12f17-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="12f17-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-865">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-865">Office 2019 on Windows</span></span><br><span data-ttu-id="12f17-866">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-866">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-867">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-867">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-869">- Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-869">- Selection</span></span><br><span data-ttu-id="12f17-870">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-870">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-871">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-871">Office 2016 on Windows</span></span><br><span data-ttu-id="12f17-872">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-872">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-873">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-873">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-875">- Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-875">- Selection</span></span><br><span data-ttu-id="12f17-876">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-876">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="12f17-877">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="12f17-877">Office 2013 on Windows</span></span><br><span data-ttu-id="12f17-878">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="12f17-878">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="12f17-879">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="12f17-879">- TaskPane</span></span></td>
    <td> <span data-ttu-id="12f17-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="12f17-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="12f17-881">- Seleção</span><span class="sxs-lookup"><span data-stu-id="12f17-881">- Selection</span></span><br><span data-ttu-id="12f17-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="12f17-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="12f17-883">Confira também</span><span class="sxs-lookup"><span data-stu-id="12f17-883">See also</span></span>

- [<span data-ttu-id="12f17-884">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="12f17-884">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="12f17-885">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="12f17-885">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="12f17-886">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="12f17-886">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="12f17-887">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="12f17-887">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="12f17-888">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="12f17-888">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="12f17-889">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="12f17-889">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="12f17-890">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="12f17-890">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="12f17-891">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="12f17-891">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="12f17-892">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12f17-892">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="12f17-893">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="12f17-893">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="12f17-894">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="12f17-894">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
