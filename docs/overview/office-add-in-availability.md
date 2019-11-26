---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: ecb906e595c08b973b5146416a5317d59547ed39
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757482"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b8742-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8742-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b8742-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="b8742-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b8742-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="b8742-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b8742-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="b8742-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b8742-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b8742-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b8742-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b8742-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b8742-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b8742-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b8742-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b8742-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8742-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-114">- TaskPane</span></span><br><span data-ttu-id="b8742-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-115">
        - Content</span></span><br><span data-ttu-id="b8742-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-116">
        - Custom Functions</span></span><br><span data-ttu-id="b8742-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b8742-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b8742-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8742-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8742-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8742-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8742-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="b8742-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="b8742-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-130">
        - BindingEvents</span></span><br><span data-ttu-id="b8742-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-131">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-132">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-133">
        - File</span></span><br><span data-ttu-id="b8742-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-134">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-136">
        - Selection</span></span><br><span data-ttu-id="b8742-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-137">
        - Settings</span></span><br><span data-ttu-id="b8742-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-138">
        - TableBindings</span></span><br><span data-ttu-id="b8742-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-139">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-140">
        - TextBindings</span></span><br><span data-ttu-id="b8742-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-142">Office on Windows</span></span><br><span data-ttu-id="b8742-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-144">- TaskPane</span></span><br><span data-ttu-id="b8742-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-145">
        - Content</span></span><br><span data-ttu-id="b8742-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-146">
        - Custom Functions</span></span><br><span data-ttu-id="b8742-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b8742-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b8742-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8742-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8742-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8742-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8742-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b8742-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-161">
        - BindingEvents</span></span><br><span data-ttu-id="b8742-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-162">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-163">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-164">
        - File</span></span><br><span data-ttu-id="b8742-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-165">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-167">
        - Selection</span></span><br><span data-ttu-id="b8742-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-168">
        - Settings</span></span><br><span data-ttu-id="b8742-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-169">
        - TableBindings</span></span><br><span data-ttu-id="b8742-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-170">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-171">
        - TextBindings</span></span><br><span data-ttu-id="b8742-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-173">Office 2019 on Windows</span></span><br><span data-ttu-id="b8742-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8742-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-175">- TaskPane</span></span><br><span data-ttu-id="b8742-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-176">
        - Content</span></span><br><span data-ttu-id="b8742-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8742-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-188">- BindingEvents</span></span><br><span data-ttu-id="b8742-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-189">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-190">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-191">
        - File</span></span><br><span data-ttu-id="b8742-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-192">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-194">
        - Selection</span></span><br><span data-ttu-id="b8742-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-195">
        - Settings</span></span><br><span data-ttu-id="b8742-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-196">
        - TableBindings</span></span><br><span data-ttu-id="b8742-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-197">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-198">
        - TextBindings</span></span><br><span data-ttu-id="b8742-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-200">Office 2016 on Windows</span></span><br><span data-ttu-id="b8742-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8742-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-202">- TaskPane</span></span><br><span data-ttu-id="b8742-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-203">
        - Content</span></span></td>
    <td><span data-ttu-id="b8742-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8742-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-207">- BindingEvents</span></span><br><span data-ttu-id="b8742-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-208">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-209">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-210">
        - File</span></span><br><span data-ttu-id="b8742-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-211">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-213">
        - Selection</span></span><br><span data-ttu-id="b8742-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-214">
        - Settings</span></span><br><span data-ttu-id="b8742-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-215">
        - TableBindings</span></span><br><span data-ttu-id="b8742-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-216">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-217">
        - TextBindings</span></span><br><span data-ttu-id="b8742-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-219">Office 2013 on Windows</span></span><br><span data-ttu-id="b8742-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8742-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-221">
        - TaskPane</span></span><br><span data-ttu-id="b8742-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b8742-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8742-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8742-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-225">
        - BindingEvents</span></span><br><span data-ttu-id="b8742-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-226">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-227">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-228">
        - File</span></span><br><span data-ttu-id="b8742-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-229">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-231">
        - Selection</span></span><br><span data-ttu-id="b8742-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-232">
        - Settings</span></span><br><span data-ttu-id="b8742-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-233">
        - TableBindings</span></span><br><span data-ttu-id="b8742-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-234">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-235">
        - TextBindings</span></span><br><span data-ttu-id="b8742-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b8742-237">Office on iPad</span></span><br><span data-ttu-id="b8742-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8742-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-239">- TaskPane</span></span><br><span data-ttu-id="b8742-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-240">
        - Content</span></span></td>
    <td><span data-ttu-id="b8742-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8742-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8742-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8742-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8742-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-253">- BindingEvents</span></span><br><span data-ttu-id="b8742-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-254">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-255">
        - File</span></span><br><span data-ttu-id="b8742-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-256">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-258">
        - Selection</span></span><br><span data-ttu-id="b8742-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-259">
        - Settings</span></span><br><span data-ttu-id="b8742-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-260">
        - TableBindings</span></span><br><span data-ttu-id="b8742-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-261">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-262">
        - TextBindings</span></span><br><span data-ttu-id="b8742-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-264">Office on Mac</span></span><br><span data-ttu-id="b8742-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8742-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-266">- TaskPane</span></span><br><span data-ttu-id="b8742-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-267">
        - Content</span></span><br><span data-ttu-id="b8742-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-268">
        - Custom Functions</span></span><br><span data-ttu-id="b8742-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8742-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b8742-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b8742-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="b8742-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="b8742-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b8742-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-283">- BindingEvents</span></span><br><span data-ttu-id="b8742-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-284">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-285">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-286">
        - File</span></span><br><span data-ttu-id="b8742-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-287">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-289">
        - PdfFile</span></span><br><span data-ttu-id="b8742-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-290">
        - Selection</span></span><br><span data-ttu-id="b8742-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-291">
        - Settings</span></span><br><span data-ttu-id="b8742-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-292">
        - TableBindings</span></span><br><span data-ttu-id="b8742-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-293">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-294">
        - TextBindings</span></span><br><span data-ttu-id="b8742-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-296">Office 2019 on Mac</span></span><br><span data-ttu-id="b8742-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8742-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-298">- TaskPane</span></span><br><span data-ttu-id="b8742-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-299">
        - Content</span></span><br><span data-ttu-id="b8742-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b8742-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b8742-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b8742-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b8742-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b8742-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b8742-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b8742-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b8742-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-311">- BindingEvents</span></span><br><span data-ttu-id="b8742-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-312">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-313">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-314">
        - File</span></span><br><span data-ttu-id="b8742-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-315">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-317">
        - PdfFile</span></span><br><span data-ttu-id="b8742-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-318">
        - Selection</span></span><br><span data-ttu-id="b8742-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-319">
        - Settings</span></span><br><span data-ttu-id="b8742-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-320">
        - TableBindings</span></span><br><span data-ttu-id="b8742-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-321">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-322">
        - TextBindings</span></span><br><span data-ttu-id="b8742-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-324">Office 2016 on Mac</span></span><br><span data-ttu-id="b8742-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b8742-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-326">- TaskPane</span></span><br><span data-ttu-id="b8742-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-327">
        - Content</span></span></td>
    <td><span data-ttu-id="b8742-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b8742-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8742-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b8742-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-331">- BindingEvents</span></span><br><span data-ttu-id="b8742-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-332">
        - CompressedFile</span></span><br><span data-ttu-id="b8742-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-333">
        - DocumentEvents</span></span><br><span data-ttu-id="b8742-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-334">
        - File</span></span><br><span data-ttu-id="b8742-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-335">
        - MatrixBindings</span></span><br><span data-ttu-id="b8742-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="b8742-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-337">
        - PdfFile</span></span><br><span data-ttu-id="b8742-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-338">
        - Selection</span></span><br><span data-ttu-id="b8742-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-339">
        - Settings</span></span><br><span data-ttu-id="b8742-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-340">
        - TableBindings</span></span><br><span data-ttu-id="b8742-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-341">
        - TableCoercion</span></span><br><span data-ttu-id="b8742-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-342">
        - TextBindings</span></span><br><span data-ttu-id="b8742-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8742-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b8742-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="b8742-345">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b8742-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b8742-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b8742-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b8742-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-350">Office on the web</span></span></td>
    <td><span data-ttu-id="b8742-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8742-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-353">Office on Windows</span></span><br><span data-ttu-id="b8742-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b8742-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8742-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-357">Office for Mac</span></span><br><span data-ttu-id="b8742-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b8742-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b8742-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b8742-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b8742-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="b8742-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8742-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-362">Platform</span></span></th>
    <th><span data-ttu-id="b8742-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-363">Extension points</span></span></th>
    <th><span data-ttu-id="b8742-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8742-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-366">Office on the web</span></span><br><span data-ttu-id="b8742-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="b8742-367">(modern)</span></span></td>
    <td> <span data-ttu-id="b8742-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-368">- Mail Read</span></span><br><span data-ttu-id="b8742-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-369">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8742-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8742-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8742-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-380">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-380">Office on the web</span></span><br><span data-ttu-id="b8742-381">(clássico)</span><span class="sxs-lookup"><span data-stu-id="b8742-381">(classic)</span></span></td>
    <td> <span data-ttu-id="b8742-382">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-382">- Mail Read</span></span><br><span data-ttu-id="b8742-383">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-383">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8742-391">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-392">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-392">Office on Windows</span></span><br><span data-ttu-id="b8742-393">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-394">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-394">- Mail Read</span></span><br><span data-ttu-id="b8742-395">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-395">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8742-397">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b8742-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b8742-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8742-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8742-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8742-406">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-407">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-407">Office 2019 on Windows</span></span><br><span data-ttu-id="b8742-408">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-409">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-409">- Mail Read</span></span><br><span data-ttu-id="b8742-410">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-410">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8742-412">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b8742-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b8742-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8742-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b8742-420">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-421">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-421">Office 2016 on Windows</span></span><br><span data-ttu-id="b8742-422">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-423">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-423">- Mail Read</span></span><br><span data-ttu-id="b8742-424">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-424">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b8742-426">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b8742-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b8742-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b8742-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-432">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-432">Office 2013 on Windows</span></span><br><span data-ttu-id="b8742-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-434">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-434">- Mail Read</span></span><br><span data-ttu-id="b8742-435">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="b8742-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b8742-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b8742-440">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-441">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="b8742-441">Office on iOS</span></span><br><span data-ttu-id="b8742-442">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-443">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-443">- Mail Read</span></span><br><span data-ttu-id="b8742-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8742-450">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-451">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-451">Office on Mac</span></span><br><span data-ttu-id="b8742-452">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-453">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-453">- Mail Read</span></span><br><span data-ttu-id="b8742-454">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-454">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b8742-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b8742-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b8742-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b8742-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b8742-464">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-465">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-465">Office 2019 on Mac</span></span><br><span data-ttu-id="b8742-466">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-467">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-467">- Mail Read</span></span><br><span data-ttu-id="b8742-468">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-468">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8742-476">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-477">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-477">Office 2016 on Mac</span></span><br><span data-ttu-id="b8742-478">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-479">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-479">- Mail Read</span></span><br><span data-ttu-id="b8742-480">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b8742-480">
      - Mail Compose</span></span><br><span data-ttu-id="b8742-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b8742-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b8742-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b8742-488">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-489">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="b8742-489">Office on Android</span></span><br><span data-ttu-id="b8742-490">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-491">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b8742-491">- Mail Read</span></span><br><span data-ttu-id="b8742-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b8742-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b8742-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b8742-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b8742-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b8742-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b8742-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b8742-498">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b8742-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8742-499">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b8742-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b8742-500">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="b8742-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b8742-501">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="b8742-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b8742-502">Word</span><span class="sxs-lookup"><span data-stu-id="b8742-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8742-503">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-503">Platform</span></span></th>
    <th><span data-ttu-id="b8742-504">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-504">Extension points</span></span></th>
    <th><span data-ttu-id="b8742-505">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8742-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-507">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8742-508">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-508">- TaskPane</span></span><br><span data-ttu-id="b8742-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8742-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-516">- BindingEvents</span></span><br><span data-ttu-id="b8742-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-518">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-519">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-519">
         - File</span></span><br><span data-ttu-id="b8742-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-521">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-524">
         - PdfFile</span></span><br><span data-ttu-id="b8742-525">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-525">
         - Selection</span></span><br><span data-ttu-id="b8742-526">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-526">
         - Settings</span></span><br><span data-ttu-id="b8742-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-527">
         - TableBindings</span></span><br><span data-ttu-id="b8742-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-528">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-529">
         - TextBindings</span></span><br><span data-ttu-id="b8742-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-530">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-532">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-532">Office on Windows</span></span><br><span data-ttu-id="b8742-533">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-534">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-534">- TaskPane</span></span><br><span data-ttu-id="b8742-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8742-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-542">- BindingEvents</span></span><br><span data-ttu-id="b8742-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-543">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-545">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-546">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-546">
         - File</span></span><br><span data-ttu-id="b8742-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-548">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-551">
         - PdfFile</span></span><br><span data-ttu-id="b8742-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-552">
         - Selection</span></span><br><span data-ttu-id="b8742-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-553">
         - Settings</span></span><br><span data-ttu-id="b8742-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-554">
         - TableBindings</span></span><br><span data-ttu-id="b8742-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-555">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-556">
         - TextBindings</span></span><br><span data-ttu-id="b8742-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-557">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-559">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-559">Office 2019 on Windows</span></span><br><span data-ttu-id="b8742-560">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-561">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-561">- TaskPane</span></span><br><span data-ttu-id="b8742-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-568">- BindingEvents</span></span><br><span data-ttu-id="b8742-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-569">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-571">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-572">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-572">
         - File</span></span><br><span data-ttu-id="b8742-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-574">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-577">
         - PdfFile</span></span><br><span data-ttu-id="b8742-578">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-578">
         - Selection</span></span><br><span data-ttu-id="b8742-579">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-579">
         - Settings</span></span><br><span data-ttu-id="b8742-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-580">
         - TableBindings</span></span><br><span data-ttu-id="b8742-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-581">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-582">
         - TextBindings</span></span><br><span data-ttu-id="b8742-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-583">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-585">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-585">Office 2016 on Windows</span></span><br><span data-ttu-id="b8742-586">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-587">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8742-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-591">- BindingEvents</span></span><br><span data-ttu-id="b8742-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-592">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-594">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-595">
         - File</span></span><br><span data-ttu-id="b8742-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-597">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-600">
         - PdfFile</span></span><br><span data-ttu-id="b8742-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-601">
         - Selection</span></span><br><span data-ttu-id="b8742-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-602">
         - Settings</span></span><br><span data-ttu-id="b8742-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-603">
         - TableBindings</span></span><br><span data-ttu-id="b8742-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-604">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-605">
         - TextBindings</span></span><br><span data-ttu-id="b8742-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-606">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-608">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-608">Office 2013 on Windows</span></span><br><span data-ttu-id="b8742-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8742-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8742-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-613">- BindingEvents</span></span><br><span data-ttu-id="b8742-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-614">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-616">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-617">
         - File</span></span><br><span data-ttu-id="b8742-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-619">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-622">
         - PdfFile</span></span><br><span data-ttu-id="b8742-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-623">
         - Selection</span></span><br><span data-ttu-id="b8742-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-624">
         - Settings</span></span><br><span data-ttu-id="b8742-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-625">
         - TableBindings</span></span><br><span data-ttu-id="b8742-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-626">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-627">
         - TextBindings</span></span><br><span data-ttu-id="b8742-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-628">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-630">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b8742-630">Office on iPad</span></span><br><span data-ttu-id="b8742-631">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8742-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-638">- BindingEvents</span></span><br><span data-ttu-id="b8742-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-639">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-641">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-642">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-642">
         - File</span></span><br><span data-ttu-id="b8742-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-644">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-647">
         - PdfFile</span></span><br><span data-ttu-id="b8742-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-648">
         - Selection</span></span><br><span data-ttu-id="b8742-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-649">
         - Settings</span></span><br><span data-ttu-id="b8742-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-650">
         - TableBindings</span></span><br><span data-ttu-id="b8742-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-651">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-652">
         - TextBindings</span></span><br><span data-ttu-id="b8742-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-653">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-655">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-655">Office on Mac</span></span><br><span data-ttu-id="b8742-656">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-657">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-657">- TaskPane</span></span><br><span data-ttu-id="b8742-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8742-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-665">- BindingEvents</span></span><br><span data-ttu-id="b8742-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-666">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-668">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-669">
         - File</span></span><br><span data-ttu-id="b8742-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-671">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-674">
         - PdfFile</span></span><br><span data-ttu-id="b8742-675">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-675">
         - Selection</span></span><br><span data-ttu-id="b8742-676">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-676">
         - Settings</span></span><br><span data-ttu-id="b8742-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-677">
         - TableBindings</span></span><br><span data-ttu-id="b8742-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-678">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-679">
         - TextBindings</span></span><br><span data-ttu-id="b8742-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-680">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-682">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-682">Office 2019 on Mac</span></span><br><span data-ttu-id="b8742-683">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-684">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-684">- TaskPane</span></span><br><span data-ttu-id="b8742-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b8742-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b8742-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b8742-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b8742-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-691">- BindingEvents</span></span><br><span data-ttu-id="b8742-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-692">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-694">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-695">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-695">
         - File</span></span><br><span data-ttu-id="b8742-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-697">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-700">
         - PdfFile</span></span><br><span data-ttu-id="b8742-701">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-701">
         - Selection</span></span><br><span data-ttu-id="b8742-702">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-702">
         - Settings</span></span><br><span data-ttu-id="b8742-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-703">
         - TableBindings</span></span><br><span data-ttu-id="b8742-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-704">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-705">
         - TextBindings</span></span><br><span data-ttu-id="b8742-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-706">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-708">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-708">Office 2016 on Mac</span></span><br><span data-ttu-id="b8742-709">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-710">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b8742-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b8742-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b8742-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-714">- BindingEvents</span></span><br><span data-ttu-id="b8742-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-715">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b8742-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="b8742-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-717">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-718">
         - File</span></span><br><span data-ttu-id="b8742-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-720">
         - MatrixBindings</span></span><br><span data-ttu-id="b8742-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="b8742-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b8742-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-723">
         - PdfFile</span></span><br><span data-ttu-id="b8742-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-724">
         - Selection</span></span><br><span data-ttu-id="b8742-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-725">
         - Settings</span></span><br><span data-ttu-id="b8742-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-726">
         - TableBindings</span></span><br><span data-ttu-id="b8742-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-727">
         - TableCoercion</span></span><br><span data-ttu-id="b8742-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b8742-728">
         - TextBindings</span></span><br><span data-ttu-id="b8742-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-729">
         - TextCoercion</span></span><br><span data-ttu-id="b8742-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b8742-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b8742-731">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b8742-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b8742-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b8742-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8742-733">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-733">Platform</span></span></th>
    <th><span data-ttu-id="b8742-734">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-734">Extension points</span></span></th>
    <th><span data-ttu-id="b8742-735">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8742-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-737">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8742-738">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-738">- Content</span></span><br><span data-ttu-id="b8742-739">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-739">
         - TaskPane</span></span><br><span data-ttu-id="b8742-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8742-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8742-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-745">- ActiveView</span></span><br><span data-ttu-id="b8742-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-746">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-747">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-748">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-748">
         - File</span></span><br><span data-ttu-id="b8742-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-749">
         - PdfFile</span></span><br><span data-ttu-id="b8742-750">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-750">
         - Selection</span></span><br><span data-ttu-id="b8742-751">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-751">
         - Settings</span></span><br><span data-ttu-id="b8742-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-753">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-753">Office on Windows</span></span><br><span data-ttu-id="b8742-754">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-755">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-755">- Content</span></span><br><span data-ttu-id="b8742-756">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-756">
         - TaskPane</span></span><br><span data-ttu-id="b8742-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8742-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8742-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-762">- ActiveView</span></span><br><span data-ttu-id="b8742-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-763">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-764">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-765">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-765">
         - File</span></span><br><span data-ttu-id="b8742-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-766">
         - PdfFile</span></span><br><span data-ttu-id="b8742-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-767">
         - Selection</span></span><br><span data-ttu-id="b8742-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-768">
         - Settings</span></span><br><span data-ttu-id="b8742-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-770">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-770">Office 2019 on Windows</span></span><br><span data-ttu-id="b8742-771">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-772">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-772">- Content</span></span><br><span data-ttu-id="b8742-773">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-773">
         - TaskPane</span></span><br><span data-ttu-id="b8742-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-777">- ActiveView</span></span><br><span data-ttu-id="b8742-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-778">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-779">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-780">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-780">
         - File</span></span><br><span data-ttu-id="b8742-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-781">
         - PdfFile</span></span><br><span data-ttu-id="b8742-782">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-782">
         - Selection</span></span><br><span data-ttu-id="b8742-783">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-783">
         - Settings</span></span><br><span data-ttu-id="b8742-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-785">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-785">Office 2016 on Windows</span></span><br><span data-ttu-id="b8742-786">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-787">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-787">- Content</span></span><br><span data-ttu-id="b8742-788">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8742-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8742-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-791">- ActiveView</span></span><br><span data-ttu-id="b8742-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-792">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-793">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-794">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-794">
         - File</span></span><br><span data-ttu-id="b8742-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-795">
         - PdfFile</span></span><br><span data-ttu-id="b8742-796">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-796">
         - Selection</span></span><br><span data-ttu-id="b8742-797">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-797">
         - Settings</span></span><br><span data-ttu-id="b8742-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-799">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-799">Office 2013 on Windows</span></span><br><span data-ttu-id="b8742-800">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-801">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-801">- Content</span></span><br><span data-ttu-id="b8742-802">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b8742-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8742-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8742-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-805">- ActiveView</span></span><br><span data-ttu-id="b8742-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-806">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-807">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-808">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-808">
         - File</span></span><br><span data-ttu-id="b8742-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-809">
         - PdfFile</span></span><br><span data-ttu-id="b8742-810">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-810">
         - Selection</span></span><br><span data-ttu-id="b8742-811">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-811">
         - Settings</span></span><br><span data-ttu-id="b8742-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-813">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b8742-813">Office on iPad</span></span><br><span data-ttu-id="b8742-814">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-815">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-815">- Content</span></span><br><span data-ttu-id="b8742-816">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8742-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-820">- ActiveView</span></span><br><span data-ttu-id="b8742-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-821">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-822">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-823">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-823">
         - File</span></span><br><span data-ttu-id="b8742-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-824">
         - PdfFile</span></span><br><span data-ttu-id="b8742-825">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-825">
         - Selection</span></span><br><span data-ttu-id="b8742-826">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-826">
         - Settings</span></span><br><span data-ttu-id="b8742-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-828">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-828">Office on Mac</span></span><br><span data-ttu-id="b8742-829">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b8742-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b8742-830">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-830">- Content</span></span><br><span data-ttu-id="b8742-831">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-831">
         - TaskPane</span></span><br><span data-ttu-id="b8742-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b8742-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b8742-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b8742-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b8742-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-837">- ActiveView</span></span><br><span data-ttu-id="b8742-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-838">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-839">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-840">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-840">
         - File</span></span><br><span data-ttu-id="b8742-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-841">
         - PdfFile</span></span><br><span data-ttu-id="b8742-842">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-842">
         - Selection</span></span><br><span data-ttu-id="b8742-843">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-843">
         - Settings</span></span><br><span data-ttu-id="b8742-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-845">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-845">Office 2019 on Mac</span></span><br><span data-ttu-id="b8742-846">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-847">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-847">- Content</span></span><br><span data-ttu-id="b8742-848">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-848">
         - TaskPane</span></span><br><span data-ttu-id="b8742-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-852">- ActiveView</span></span><br><span data-ttu-id="b8742-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-853">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-854">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-855">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-855">
         - File</span></span><br><span data-ttu-id="b8742-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-856">
         - PdfFile</span></span><br><span data-ttu-id="b8742-857">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-857">
         - Selection</span></span><br><span data-ttu-id="b8742-858">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-858">
         - Settings</span></span><br><span data-ttu-id="b8742-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-860">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-860">Office 2016 on Mac</span></span><br><span data-ttu-id="b8742-861">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-862">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-862">- Content</span></span><br><span data-ttu-id="b8742-863">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b8742-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b8742-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b8742-866">- ActiveView</span></span><br><span data-ttu-id="b8742-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b8742-867">
         - CompressedFile</span></span><br><span data-ttu-id="b8742-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-868">
         - DocumentEvents</span></span><br><span data-ttu-id="b8742-869">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b8742-869">
         - File</span></span><br><span data-ttu-id="b8742-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b8742-870">
         - PdfFile</span></span><br><span data-ttu-id="b8742-871">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-871">
         - Selection</span></span><br><span data-ttu-id="b8742-872">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-872">
         - Settings</span></span><br><span data-ttu-id="b8742-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b8742-874">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b8742-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b8742-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="b8742-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8742-876">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-876">Platform</span></span></th>
    <th><span data-ttu-id="b8742-877">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-877">Extension points</span></span></th>
    <th><span data-ttu-id="b8742-878">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8742-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-880">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b8742-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="b8742-881">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b8742-881">- Content</span></span><br><span data-ttu-id="b8742-882">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-882">
         - TaskPane</span></span><br><span data-ttu-id="b8742-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b8742-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b8742-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b8742-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b8742-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b8742-887">- DocumentEvents</span></span><br><span data-ttu-id="b8742-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="b8742-889">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b8742-889">
         - Settings</span></span><br><span data-ttu-id="b8742-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b8742-891">Project</span><span class="sxs-lookup"><span data-stu-id="b8742-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b8742-892">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b8742-892">Platform</span></span></th>
    <th><span data-ttu-id="b8742-893">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b8742-893">Extension points</span></span></th>
    <th><span data-ttu-id="b8742-894">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b8742-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="b8742-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b8742-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-896">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-896">Office 2019 on Windows</span></span><br><span data-ttu-id="b8742-897">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-898">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-900">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-900">- Selection</span></span><br><span data-ttu-id="b8742-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-902">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-902">Office 2016 on Windows</span></span><br><span data-ttu-id="b8742-903">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-904">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-906">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-906">- Selection</span></span><br><span data-ttu-id="b8742-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b8742-908">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b8742-908">Office 2013 on Windows</span></span><br><span data-ttu-id="b8742-909">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b8742-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b8742-910">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b8742-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b8742-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b8742-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b8742-912">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b8742-912">- Selection</span></span><br><span data-ttu-id="b8742-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b8742-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b8742-914">Confira também</span><span class="sxs-lookup"><span data-stu-id="b8742-914">See also</span></span>

- [<span data-ttu-id="b8742-915">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b8742-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b8742-916">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="b8742-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b8742-917">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="b8742-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b8742-918">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="b8742-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b8742-919">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="b8742-919">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b8742-920">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="b8742-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b8742-921">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="b8742-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b8742-922">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="b8742-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b8742-923">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b8742-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b8742-924">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b8742-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b8742-925">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b8742-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
