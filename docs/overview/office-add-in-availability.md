---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: 956ee6b8a9e990a3d6d942ee4a65a1e9275ea025
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851366"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="44578-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="44578-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="44578-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="44578-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="44578-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="44578-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="44578-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="44578-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="44578-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="44578-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="44578-108">Excel</span><span class="sxs-lookup"><span data-stu-id="44578-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="44578-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="44578-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="44578-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="44578-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="44578-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-114">- TaskPane</span></span><br><span data-ttu-id="44578-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-115">
        - Content</span></span><br><span data-ttu-id="44578-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-116">
        - Custom Functions</span></span><br><span data-ttu-id="44578-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="44578-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="44578-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="44578-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="44578-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="44578-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="44578-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="44578-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="44578-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-130">
        - BindingEvents</span></span><br><span data-ttu-id="44578-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-131">
        - CompressedFile</span></span><br><span data-ttu-id="44578-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-132">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-133">
        - File</span></span><br><span data-ttu-id="44578-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-134">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-136">
        - Selection</span></span><br><span data-ttu-id="44578-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-137">
        - Settings</span></span><br><span data-ttu-id="44578-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-138">
        - TableBindings</span></span><br><span data-ttu-id="44578-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-139">
        - TableCoercion</span></span><br><span data-ttu-id="44578-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-140">
        - TextBindings</span></span><br><span data-ttu-id="44578-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-142">Office on Windows</span></span><br><span data-ttu-id="44578-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-144">- TaskPane</span></span><br><span data-ttu-id="44578-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-145">
        - Content</span></span><br><span data-ttu-id="44578-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-146">
        - Custom Functions</span></span><br><span data-ttu-id="44578-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="44578-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="44578-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="44578-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="44578-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="44578-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="44578-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="44578-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-161">
        - BindingEvents</span></span><br><span data-ttu-id="44578-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-162">
        - CompressedFile</span></span><br><span data-ttu-id="44578-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-163">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-164">
        - File</span></span><br><span data-ttu-id="44578-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-165">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-167">
        - Selection</span></span><br><span data-ttu-id="44578-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-168">
        - Settings</span></span><br><span data-ttu-id="44578-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-169">
        - TableBindings</span></span><br><span data-ttu-id="44578-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-170">
        - TableCoercion</span></span><br><span data-ttu-id="44578-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-171">
        - TextBindings</span></span><br><span data-ttu-id="44578-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-173">Office 2019 on Windows</span></span><br><span data-ttu-id="44578-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="44578-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-175">- TaskPane</span></span><br><span data-ttu-id="44578-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-176">
        - Content</span></span><br><span data-ttu-id="44578-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="44578-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-188">- BindingEvents</span></span><br><span data-ttu-id="44578-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-189">
        - CompressedFile</span></span><br><span data-ttu-id="44578-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-190">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-191">
        - File</span></span><br><span data-ttu-id="44578-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-192">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-194">
        - Selection</span></span><br><span data-ttu-id="44578-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-195">
        - Settings</span></span><br><span data-ttu-id="44578-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-196">
        - TableBindings</span></span><br><span data-ttu-id="44578-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-197">
        - TableCoercion</span></span><br><span data-ttu-id="44578-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-198">
        - TextBindings</span></span><br><span data-ttu-id="44578-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-200">Office 2016 on Windows</span></span><br><span data-ttu-id="44578-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="44578-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-202">- TaskPane</span></span><br><span data-ttu-id="44578-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-203">
        - Content</span></span></td>
    <td><span data-ttu-id="44578-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="44578-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-207">- BindingEvents</span></span><br><span data-ttu-id="44578-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-208">
        - CompressedFile</span></span><br><span data-ttu-id="44578-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-209">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-210">
        - File</span></span><br><span data-ttu-id="44578-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-211">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-213">
        - Selection</span></span><br><span data-ttu-id="44578-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-214">
        - Settings</span></span><br><span data-ttu-id="44578-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-215">
        - TableBindings</span></span><br><span data-ttu-id="44578-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-216">
        - TableCoercion</span></span><br><span data-ttu-id="44578-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-217">
        - TextBindings</span></span><br><span data-ttu-id="44578-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-219">Office 2013 on Windows</span></span><br><span data-ttu-id="44578-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="44578-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-221">
        - TaskPane</span></span><br><span data-ttu-id="44578-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="44578-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="44578-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="44578-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-225">
        - BindingEvents</span></span><br><span data-ttu-id="44578-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-226">
        - CompressedFile</span></span><br><span data-ttu-id="44578-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-227">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-228">
        - File</span></span><br><span data-ttu-id="44578-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-229">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-231">
        - Selection</span></span><br><span data-ttu-id="44578-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-232">
        - Settings</span></span><br><span data-ttu-id="44578-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-233">
        - TableBindings</span></span><br><span data-ttu-id="44578-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-234">
        - TableCoercion</span></span><br><span data-ttu-id="44578-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-235">
        - TextBindings</span></span><br><span data-ttu-id="44578-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="44578-237">Office on iPad</span></span><br><span data-ttu-id="44578-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="44578-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-239">- TaskPane</span></span><br><span data-ttu-id="44578-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-240">
        - Content</span></span></td>
    <td><span data-ttu-id="44578-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="44578-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="44578-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="44578-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="44578-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-253">- BindingEvents</span></span><br><span data-ttu-id="44578-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-254">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-255">
        - File</span></span><br><span data-ttu-id="44578-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-256">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-258">
        - Selection</span></span><br><span data-ttu-id="44578-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-259">
        - Settings</span></span><br><span data-ttu-id="44578-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-260">
        - TableBindings</span></span><br><span data-ttu-id="44578-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-261">
        - TableCoercion</span></span><br><span data-ttu-id="44578-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-262">
        - TextBindings</span></span><br><span data-ttu-id="44578-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-264">Office on Mac</span></span><br><span data-ttu-id="44578-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="44578-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-266">- TaskPane</span></span><br><span data-ttu-id="44578-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-267">
        - Content</span></span><br><span data-ttu-id="44578-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-268">
        - Custom Functions</span></span><br><span data-ttu-id="44578-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="44578-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="44578-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="44578-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="44578-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="44578-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="44578-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-283">- BindingEvents</span></span><br><span data-ttu-id="44578-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-284">
        - CompressedFile</span></span><br><span data-ttu-id="44578-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-285">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-286">
        - File</span></span><br><span data-ttu-id="44578-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-287">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-289">
        - PdfFile</span></span><br><span data-ttu-id="44578-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-290">
        - Selection</span></span><br><span data-ttu-id="44578-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-291">
        - Settings</span></span><br><span data-ttu-id="44578-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-292">
        - TableBindings</span></span><br><span data-ttu-id="44578-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-293">
        - TableCoercion</span></span><br><span data-ttu-id="44578-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-294">
        - TextBindings</span></span><br><span data-ttu-id="44578-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-296">Office 2019 on Mac</span></span><br><span data-ttu-id="44578-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="44578-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-298">- TaskPane</span></span><br><span data-ttu-id="44578-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-299">
        - Content</span></span><br><span data-ttu-id="44578-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="44578-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="44578-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="44578-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="44578-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="44578-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="44578-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="44578-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="44578-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-311">- BindingEvents</span></span><br><span data-ttu-id="44578-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-312">
        - CompressedFile</span></span><br><span data-ttu-id="44578-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-313">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-314">
        - File</span></span><br><span data-ttu-id="44578-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-315">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-317">
        - PdfFile</span></span><br><span data-ttu-id="44578-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-318">
        - Selection</span></span><br><span data-ttu-id="44578-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-319">
        - Settings</span></span><br><span data-ttu-id="44578-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-320">
        - TableBindings</span></span><br><span data-ttu-id="44578-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-321">
        - TableCoercion</span></span><br><span data-ttu-id="44578-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-322">
        - TextBindings</span></span><br><span data-ttu-id="44578-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-324">Office 2016 on Mac</span></span><br><span data-ttu-id="44578-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="44578-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-326">- TaskPane</span></span><br><span data-ttu-id="44578-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-327">
        - Content</span></span></td>
    <td><span data-ttu-id="44578-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="44578-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="44578-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="44578-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-331">- BindingEvents</span></span><br><span data-ttu-id="44578-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-332">
        - CompressedFile</span></span><br><span data-ttu-id="44578-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-333">
        - DocumentEvents</span></span><br><span data-ttu-id="44578-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-334">
        - File</span></span><br><span data-ttu-id="44578-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-335">
        - MatrixBindings</span></span><br><span data-ttu-id="44578-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="44578-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-337">
        - PdfFile</span></span><br><span data-ttu-id="44578-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-338">
        - Selection</span></span><br><span data-ttu-id="44578-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-339">
        - Settings</span></span><br><span data-ttu-id="44578-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-340">
        - TableBindings</span></span><br><span data-ttu-id="44578-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-341">
        - TableCoercion</span></span><br><span data-ttu-id="44578-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-342">
        - TextBindings</span></span><br><span data-ttu-id="44578-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="44578-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="44578-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="44578-345">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="44578-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="44578-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="44578-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="44578-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-350">Office on the web</span></span></td>
    <td><span data-ttu-id="44578-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="44578-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-353">Office on Windows</span></span><br><span data-ttu-id="44578-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="44578-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="44578-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="44578-357">Office for Mac</span></span><br><span data-ttu-id="44578-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="44578-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="44578-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="44578-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="44578-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="44578-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="44578-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-362">Platform</span></span></th>
    <th><span data-ttu-id="44578-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-363">Extension points</span></span></th>
    <th><span data-ttu-id="44578-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="44578-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-366">Office on the web</span></span><br><span data-ttu-id="44578-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="44578-367">(modern)</span></span></td>
    <td> <span data-ttu-id="44578-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-368">- Mail Read</span></span><br><span data-ttu-id="44578-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-369">
      - Mail Compose</span></span><br><span data-ttu-id="44578-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="44578-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="44578-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="44578-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-380">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-380">Office on the web</span></span><br><span data-ttu-id="44578-381">(clássico)</span><span class="sxs-lookup"><span data-stu-id="44578-381">(classic)</span></span></td>
    <td> <span data-ttu-id="44578-382">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-382">- Mail Read</span></span><br><span data-ttu-id="44578-383">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-383">
      - Mail Compose</span></span><br><span data-ttu-id="44578-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="44578-391">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-392">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-392">Office on Windows</span></span><br><span data-ttu-id="44578-393">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-394">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-394">- Mail Read</span></span><br><span data-ttu-id="44578-395">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-395">
      - Mail Compose</span></span><br><span data-ttu-id="44578-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="44578-397">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="44578-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="44578-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="44578-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="44578-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="44578-406">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-407">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-407">Office 2019 on Windows</span></span><br><span data-ttu-id="44578-408">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-409">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-409">- Mail Read</span></span><br><span data-ttu-id="44578-410">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-410">
      - Mail Compose</span></span><br><span data-ttu-id="44578-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="44578-412">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="44578-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="44578-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="44578-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="44578-420">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-421">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-421">Office 2016 on Windows</span></span><br><span data-ttu-id="44578-422">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-423">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-423">- Mail Read</span></span><br><span data-ttu-id="44578-424">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-424">
      - Mail Compose</span></span><br><span data-ttu-id="44578-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="44578-426">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="44578-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="44578-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="44578-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-432">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-432">Office 2013 on Windows</span></span><br><span data-ttu-id="44578-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-434">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-434">- Mail Read</span></span><br><span data-ttu-id="44578-435">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="44578-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="44578-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="44578-440">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-441">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="44578-441">Office on iOS</span></span><br><span data-ttu-id="44578-442">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-443">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-443">- Mail Read</span></span><br><span data-ttu-id="44578-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="44578-450">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-451">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-451">Office on Mac</span></span><br><span data-ttu-id="44578-452">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-453">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-453">- Mail Read</span></span><br><span data-ttu-id="44578-454">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-454">
      - Mail Compose</span></span><br><span data-ttu-id="44578-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="44578-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="44578-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="44578-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="44578-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="44578-464">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-465">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-465">Office 2019 on Mac</span></span><br><span data-ttu-id="44578-466">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-467">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-467">- Mail Read</span></span><br><span data-ttu-id="44578-468">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-468">
      - Mail Compose</span></span><br><span data-ttu-id="44578-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="44578-476">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-477">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-477">Office 2016 on Mac</span></span><br><span data-ttu-id="44578-478">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-479">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-479">- Mail Read</span></span><br><span data-ttu-id="44578-480">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="44578-480">
      - Mail Compose</span></span><br><span data-ttu-id="44578-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="44578-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="44578-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="44578-488">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-489">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="44578-489">Office on Android</span></span><br><span data-ttu-id="44578-490">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-491">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="44578-491">- Mail Read</span></span><br><span data-ttu-id="44578-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="44578-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="44578-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="44578-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="44578-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="44578-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="44578-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="44578-498">Não disponível</span><span class="sxs-lookup"><span data-stu-id="44578-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="44578-499">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="44578-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44578-500">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="44578-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="44578-501">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="44578-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="44578-502">Word</span><span class="sxs-lookup"><span data-stu-id="44578-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="44578-503">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-503">Platform</span></span></th>
    <th><span data-ttu-id="44578-504">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-504">Extension points</span></span></th>
    <th><span data-ttu-id="44578-505">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="44578-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-507">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="44578-508">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-508">- TaskPane</span></span><br><span data-ttu-id="44578-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="44578-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-516">- BindingEvents</span></span><br><span data-ttu-id="44578-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-518">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-519">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-519">
         - File</span></span><br><span data-ttu-id="44578-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-521">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-524">
         - PdfFile</span></span><br><span data-ttu-id="44578-525">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-525">
         - Selection</span></span><br><span data-ttu-id="44578-526">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-526">
         - Settings</span></span><br><span data-ttu-id="44578-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-527">
         - TableBindings</span></span><br><span data-ttu-id="44578-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-528">
         - TableCoercion</span></span><br><span data-ttu-id="44578-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-529">
         - TextBindings</span></span><br><span data-ttu-id="44578-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-530">
         - TextCoercion</span></span><br><span data-ttu-id="44578-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-532">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-532">Office on Windows</span></span><br><span data-ttu-id="44578-533">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-534">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-534">- TaskPane</span></span><br><span data-ttu-id="44578-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="44578-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-542">- BindingEvents</span></span><br><span data-ttu-id="44578-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-543">
         - CompressedFile</span></span><br><span data-ttu-id="44578-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-545">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-546">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-546">
         - File</span></span><br><span data-ttu-id="44578-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-548">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-551">
         - PdfFile</span></span><br><span data-ttu-id="44578-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-552">
         - Selection</span></span><br><span data-ttu-id="44578-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-553">
         - Settings</span></span><br><span data-ttu-id="44578-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-554">
         - TableBindings</span></span><br><span data-ttu-id="44578-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-555">
         - TableCoercion</span></span><br><span data-ttu-id="44578-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-556">
         - TextBindings</span></span><br><span data-ttu-id="44578-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-557">
         - TextCoercion</span></span><br><span data-ttu-id="44578-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-559">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-559">Office 2019 on Windows</span></span><br><span data-ttu-id="44578-560">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-561">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-561">- TaskPane</span></span><br><span data-ttu-id="44578-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-568">- BindingEvents</span></span><br><span data-ttu-id="44578-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-569">
         - CompressedFile</span></span><br><span data-ttu-id="44578-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-571">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-572">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-572">
         - File</span></span><br><span data-ttu-id="44578-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-574">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-577">
         - PdfFile</span></span><br><span data-ttu-id="44578-578">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-578">
         - Selection</span></span><br><span data-ttu-id="44578-579">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-579">
         - Settings</span></span><br><span data-ttu-id="44578-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-580">
         - TableBindings</span></span><br><span data-ttu-id="44578-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-581">
         - TableCoercion</span></span><br><span data-ttu-id="44578-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-582">
         - TextBindings</span></span><br><span data-ttu-id="44578-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-583">
         - TextCoercion</span></span><br><span data-ttu-id="44578-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-585">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-585">Office 2016 on Windows</span></span><br><span data-ttu-id="44578-586">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-587">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="44578-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-591">- BindingEvents</span></span><br><span data-ttu-id="44578-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-592">
         - CompressedFile</span></span><br><span data-ttu-id="44578-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-594">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-595">
         - File</span></span><br><span data-ttu-id="44578-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-597">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-600">
         - PdfFile</span></span><br><span data-ttu-id="44578-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-601">
         - Selection</span></span><br><span data-ttu-id="44578-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-602">
         - Settings</span></span><br><span data-ttu-id="44578-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-603">
         - TableBindings</span></span><br><span data-ttu-id="44578-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-604">
         - TableCoercion</span></span><br><span data-ttu-id="44578-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-605">
         - TextBindings</span></span><br><span data-ttu-id="44578-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-606">
         - TextCoercion</span></span><br><span data-ttu-id="44578-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-608">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-608">Office 2013 on Windows</span></span><br><span data-ttu-id="44578-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="44578-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="44578-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-613">- BindingEvents</span></span><br><span data-ttu-id="44578-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-614">
         - CompressedFile</span></span><br><span data-ttu-id="44578-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-616">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-617">
         - File</span></span><br><span data-ttu-id="44578-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-619">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-622">
         - PdfFile</span></span><br><span data-ttu-id="44578-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-623">
         - Selection</span></span><br><span data-ttu-id="44578-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-624">
         - Settings</span></span><br><span data-ttu-id="44578-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-625">
         - TableBindings</span></span><br><span data-ttu-id="44578-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-626">
         - TableCoercion</span></span><br><span data-ttu-id="44578-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-627">
         - TextBindings</span></span><br><span data-ttu-id="44578-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-628">
         - TextCoercion</span></span><br><span data-ttu-id="44578-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-630">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="44578-630">Office on iPad</span></span><br><span data-ttu-id="44578-631">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="44578-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-638">- BindingEvents</span></span><br><span data-ttu-id="44578-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-639">
         - CompressedFile</span></span><br><span data-ttu-id="44578-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-641">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-642">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-642">
         - File</span></span><br><span data-ttu-id="44578-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-644">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-647">
         - PdfFile</span></span><br><span data-ttu-id="44578-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-648">
         - Selection</span></span><br><span data-ttu-id="44578-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-649">
         - Settings</span></span><br><span data-ttu-id="44578-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-650">
         - TableBindings</span></span><br><span data-ttu-id="44578-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-651">
         - TableCoercion</span></span><br><span data-ttu-id="44578-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-652">
         - TextBindings</span></span><br><span data-ttu-id="44578-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-653">
         - TextCoercion</span></span><br><span data-ttu-id="44578-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-655">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-655">Office on Mac</span></span><br><span data-ttu-id="44578-656">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-657">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-657">- TaskPane</span></span><br><span data-ttu-id="44578-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="44578-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-665">- BindingEvents</span></span><br><span data-ttu-id="44578-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-666">
         - CompressedFile</span></span><br><span data-ttu-id="44578-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-668">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-669">
         - File</span></span><br><span data-ttu-id="44578-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-671">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-674">
         - PdfFile</span></span><br><span data-ttu-id="44578-675">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-675">
         - Selection</span></span><br><span data-ttu-id="44578-676">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-676">
         - Settings</span></span><br><span data-ttu-id="44578-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-677">
         - TableBindings</span></span><br><span data-ttu-id="44578-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-678">
         - TableCoercion</span></span><br><span data-ttu-id="44578-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-679">
         - TextBindings</span></span><br><span data-ttu-id="44578-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-680">
         - TextCoercion</span></span><br><span data-ttu-id="44578-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-682">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-682">Office 2019 on Mac</span></span><br><span data-ttu-id="44578-683">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-684">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-684">- TaskPane</span></span><br><span data-ttu-id="44578-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="44578-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="44578-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="44578-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="44578-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-691">- BindingEvents</span></span><br><span data-ttu-id="44578-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-692">
         - CompressedFile</span></span><br><span data-ttu-id="44578-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-694">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-695">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-695">
         - File</span></span><br><span data-ttu-id="44578-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-697">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-700">
         - PdfFile</span></span><br><span data-ttu-id="44578-701">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-701">
         - Selection</span></span><br><span data-ttu-id="44578-702">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-702">
         - Settings</span></span><br><span data-ttu-id="44578-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-703">
         - TableBindings</span></span><br><span data-ttu-id="44578-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-704">
         - TableCoercion</span></span><br><span data-ttu-id="44578-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-705">
         - TextBindings</span></span><br><span data-ttu-id="44578-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-706">
         - TextCoercion</span></span><br><span data-ttu-id="44578-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-708">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-708">Office 2016 on Mac</span></span><br><span data-ttu-id="44578-709">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-710">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="44578-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="44578-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="44578-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="44578-714">- BindingEvents</span></span><br><span data-ttu-id="44578-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-715">
         - CompressedFile</span></span><br><span data-ttu-id="44578-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="44578-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="44578-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-717">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-718">
         - File</span></span><br><span data-ttu-id="44578-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="44578-720">
         - MatrixBindings</span></span><br><span data-ttu-id="44578-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="44578-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="44578-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-723">
         - PdfFile</span></span><br><span data-ttu-id="44578-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-724">
         - Selection</span></span><br><span data-ttu-id="44578-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-725">
         - Settings</span></span><br><span data-ttu-id="44578-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="44578-726">
         - TableBindings</span></span><br><span data-ttu-id="44578-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-727">
         - TableCoercion</span></span><br><span data-ttu-id="44578-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="44578-728">
         - TextBindings</span></span><br><span data-ttu-id="44578-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-729">
         - TextCoercion</span></span><br><span data-ttu-id="44578-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="44578-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="44578-731">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="44578-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="44578-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="44578-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="44578-733">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-733">Platform</span></span></th>
    <th><span data-ttu-id="44578-734">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-734">Extension points</span></span></th>
    <th><span data-ttu-id="44578-735">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="44578-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-737">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="44578-738">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-738">- Content</span></span><br><span data-ttu-id="44578-739">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-739">
         - TaskPane</span></span><br><span data-ttu-id="44578-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="44578-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="44578-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-745">- ActiveView</span></span><br><span data-ttu-id="44578-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-746">
         - CompressedFile</span></span><br><span data-ttu-id="44578-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-747">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-748">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-748">
         - File</span></span><br><span data-ttu-id="44578-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-749">
         - PdfFile</span></span><br><span data-ttu-id="44578-750">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-750">
         - Selection</span></span><br><span data-ttu-id="44578-751">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-751">
         - Settings</span></span><br><span data-ttu-id="44578-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-753">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-753">Office on Windows</span></span><br><span data-ttu-id="44578-754">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-755">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-755">- Content</span></span><br><span data-ttu-id="44578-756">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-756">
         - TaskPane</span></span><br><span data-ttu-id="44578-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="44578-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="44578-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-762">- ActiveView</span></span><br><span data-ttu-id="44578-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-763">
         - CompressedFile</span></span><br><span data-ttu-id="44578-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-764">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-765">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-765">
         - File</span></span><br><span data-ttu-id="44578-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-766">
         - PdfFile</span></span><br><span data-ttu-id="44578-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-767">
         - Selection</span></span><br><span data-ttu-id="44578-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-768">
         - Settings</span></span><br><span data-ttu-id="44578-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-770">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-770">Office 2019 on Windows</span></span><br><span data-ttu-id="44578-771">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-772">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-772">- Content</span></span><br><span data-ttu-id="44578-773">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-773">
         - TaskPane</span></span><br><span data-ttu-id="44578-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-777">- ActiveView</span></span><br><span data-ttu-id="44578-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-778">
         - CompressedFile</span></span><br><span data-ttu-id="44578-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-779">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-780">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-780">
         - File</span></span><br><span data-ttu-id="44578-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-781">
         - PdfFile</span></span><br><span data-ttu-id="44578-782">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-782">
         - Selection</span></span><br><span data-ttu-id="44578-783">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-783">
         - Settings</span></span><br><span data-ttu-id="44578-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-785">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-785">Office 2016 on Windows</span></span><br><span data-ttu-id="44578-786">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-787">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-787">- Content</span></span><br><span data-ttu-id="44578-788">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="44578-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="44578-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-791">- ActiveView</span></span><br><span data-ttu-id="44578-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-792">
         - CompressedFile</span></span><br><span data-ttu-id="44578-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-793">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-794">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-794">
         - File</span></span><br><span data-ttu-id="44578-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-795">
         - PdfFile</span></span><br><span data-ttu-id="44578-796">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-796">
         - Selection</span></span><br><span data-ttu-id="44578-797">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-797">
         - Settings</span></span><br><span data-ttu-id="44578-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-799">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-799">Office 2013 on Windows</span></span><br><span data-ttu-id="44578-800">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-801">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-801">- Content</span></span><br><span data-ttu-id="44578-802">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="44578-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="44578-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="44578-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-805">- ActiveView</span></span><br><span data-ttu-id="44578-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-806">
         - CompressedFile</span></span><br><span data-ttu-id="44578-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-807">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-808">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-808">
         - File</span></span><br><span data-ttu-id="44578-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-809">
         - PdfFile</span></span><br><span data-ttu-id="44578-810">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-810">
         - Selection</span></span><br><span data-ttu-id="44578-811">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-811">
         - Settings</span></span><br><span data-ttu-id="44578-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-813">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="44578-813">Office on iPad</span></span><br><span data-ttu-id="44578-814">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-815">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-815">- Content</span></span><br><span data-ttu-id="44578-816">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="44578-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-820">- ActiveView</span></span><br><span data-ttu-id="44578-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-821">
         - CompressedFile</span></span><br><span data-ttu-id="44578-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-822">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-823">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-823">
         - File</span></span><br><span data-ttu-id="44578-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-824">
         - PdfFile</span></span><br><span data-ttu-id="44578-825">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-825">
         - Selection</span></span><br><span data-ttu-id="44578-826">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-826">
         - Settings</span></span><br><span data-ttu-id="44578-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-828">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-828">Office on Mac</span></span><br><span data-ttu-id="44578-829">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="44578-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="44578-830">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-830">- Content</span></span><br><span data-ttu-id="44578-831">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-831">
         - TaskPane</span></span><br><span data-ttu-id="44578-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="44578-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="44578-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="44578-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="44578-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-837">- ActiveView</span></span><br><span data-ttu-id="44578-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-838">
         - CompressedFile</span></span><br><span data-ttu-id="44578-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-839">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-840">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-840">
         - File</span></span><br><span data-ttu-id="44578-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-841">
         - PdfFile</span></span><br><span data-ttu-id="44578-842">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-842">
         - Selection</span></span><br><span data-ttu-id="44578-843">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-843">
         - Settings</span></span><br><span data-ttu-id="44578-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-845">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-845">Office 2019 on Mac</span></span><br><span data-ttu-id="44578-846">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-847">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-847">- Content</span></span><br><span data-ttu-id="44578-848">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-848">
         - TaskPane</span></span><br><span data-ttu-id="44578-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-852">- ActiveView</span></span><br><span data-ttu-id="44578-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-853">
         - CompressedFile</span></span><br><span data-ttu-id="44578-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-854">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-855">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-855">
         - File</span></span><br><span data-ttu-id="44578-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-856">
         - PdfFile</span></span><br><span data-ttu-id="44578-857">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-857">
         - Selection</span></span><br><span data-ttu-id="44578-858">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-858">
         - Settings</span></span><br><span data-ttu-id="44578-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-860">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="44578-860">Office 2016 on Mac</span></span><br><span data-ttu-id="44578-861">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-862">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-862">- Content</span></span><br><span data-ttu-id="44578-863">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="44578-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="44578-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="44578-866">- ActiveView</span></span><br><span data-ttu-id="44578-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="44578-867">
         - CompressedFile</span></span><br><span data-ttu-id="44578-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-868">
         - DocumentEvents</span></span><br><span data-ttu-id="44578-869">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="44578-869">
         - File</span></span><br><span data-ttu-id="44578-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="44578-870">
         - PdfFile</span></span><br><span data-ttu-id="44578-871">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-871">
         - Selection</span></span><br><span data-ttu-id="44578-872">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-872">
         - Settings</span></span><br><span data-ttu-id="44578-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="44578-874">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="44578-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="44578-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="44578-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="44578-876">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-876">Platform</span></span></th>
    <th><span data-ttu-id="44578-877">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-877">Extension points</span></span></th>
    <th><span data-ttu-id="44578-878">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="44578-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-880">Office na Web</span><span class="sxs-lookup"><span data-stu-id="44578-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="44578-881">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="44578-881">- Content</span></span><br><span data-ttu-id="44578-882">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-882">
         - TaskPane</span></span><br><span data-ttu-id="44578-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="44578-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="44578-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="44578-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="44578-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="44578-887">- DocumentEvents</span></span><br><span data-ttu-id="44578-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="44578-889">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="44578-889">
         - Settings</span></span><br><span data-ttu-id="44578-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="44578-891">Project</span><span class="sxs-lookup"><span data-stu-id="44578-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="44578-892">Plataforma</span><span class="sxs-lookup"><span data-stu-id="44578-892">Platform</span></span></th>
    <th><span data-ttu-id="44578-893">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="44578-893">Extension points</span></span></th>
    <th><span data-ttu-id="44578-894">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="44578-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="44578-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="44578-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-896">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-896">Office 2019 on Windows</span></span><br><span data-ttu-id="44578-897">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-898">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-900">- Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-900">- Selection</span></span><br><span data-ttu-id="44578-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-902">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-902">Office 2016 on Windows</span></span><br><span data-ttu-id="44578-903">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-904">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-906">- Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-906">- Selection</span></span><br><span data-ttu-id="44578-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="44578-908">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="44578-908">Office 2013 on Windows</span></span><br><span data-ttu-id="44578-909">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="44578-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="44578-910">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="44578-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="44578-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="44578-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="44578-912">- Seleção</span><span class="sxs-lookup"><span data-stu-id="44578-912">- Selection</span></span><br><span data-ttu-id="44578-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="44578-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="44578-914">Confira também</span><span class="sxs-lookup"><span data-stu-id="44578-914">See also</span></span>

- [<span data-ttu-id="44578-915">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="44578-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="44578-916">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="44578-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="44578-917">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="44578-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="44578-918">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="44578-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="44578-919">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="44578-919">API reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="44578-920">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="44578-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="44578-921">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="44578-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="44578-922">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="44578-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="44578-923">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="44578-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="44578-924">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="44578-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="44578-925">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="44578-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="44578-926">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="44578-926">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)