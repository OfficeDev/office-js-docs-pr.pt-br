---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901945"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="b98d9-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b98d9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="b98d9-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="b98d9-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="b98d9-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="b98d9-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="b98d9-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="b98d9-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="b98d9-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="b98d9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="b98d9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="b98d9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b98d9-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b98d9-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b98d9-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b98d9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="b98d9-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-114">- TaskPane</span></span><br><span data-ttu-id="b98d9-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-115">
        - Content</span></span><br><span data-ttu-id="b98d9-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-116">
        - Custom Functions</span></span><br><span data-ttu-id="b98d9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b98d9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b98d9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b98d9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-128">
        - BindingEvents</span></span><br><span data-ttu-id="b98d9-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-129">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-130">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-131">
        - File</span></span><br><span data-ttu-id="b98d9-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-132">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-134">
        - Selection</span></span><br><span data-ttu-id="b98d9-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-135">
        - Settings</span></span><br><span data-ttu-id="b98d9-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-136">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-137">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-138">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-140">Office on Windows</span></span><br><span data-ttu-id="b98d9-141">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-142">- TaskPane</span></span><br><span data-ttu-id="b98d9-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-143">
        - Content</span></span><br><span data-ttu-id="b98d9-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-144">
        - Custom Functions</span></span><br><span data-ttu-id="b98d9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="b98d9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="b98d9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b98d9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b98d9-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-158">
        - BindingEvents</span></span><br><span data-ttu-id="b98d9-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-159">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-160">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-161">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-161">
        - File</span></span><br><span data-ttu-id="b98d9-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-162">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-164">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-164">
        - Selection</span></span><br><span data-ttu-id="b98d9-165">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-165">
        - Settings</span></span><br><span data-ttu-id="b98d9-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-166">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-167">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-168">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-170">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-170">Office 2019 on Windows</span></span><br><span data-ttu-id="b98d9-171">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b98d9-172">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-172">- TaskPane</span></span><br><span data-ttu-id="b98d9-173">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-173">
        - Content</span></span><br><span data-ttu-id="b98d9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b98d9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-185">- BindingEvents</span></span><br><span data-ttu-id="b98d9-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-186">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-187">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-188">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-188">
        - File</span></span><br><span data-ttu-id="b98d9-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-189">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-191">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-191">
        - Selection</span></span><br><span data-ttu-id="b98d9-192">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-192">
        - Settings</span></span><br><span data-ttu-id="b98d9-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-193">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-194">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-195">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-197">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-197">Office 2016 on Windows</span></span><br><span data-ttu-id="b98d9-198">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b98d9-199">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-199">- TaskPane</span></span><br><span data-ttu-id="b98d9-200">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-200">
        - Content</span></span></td>
    <td><span data-ttu-id="b98d9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b98d9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-204">- BindingEvents</span></span><br><span data-ttu-id="b98d9-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-205">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-206">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-207">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-207">
        - File</span></span><br><span data-ttu-id="b98d9-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-208">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-210">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-210">
        - Selection</span></span><br><span data-ttu-id="b98d9-211">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-211">
        - Settings</span></span><br><span data-ttu-id="b98d9-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-212">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-213">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-214">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-216">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-216">Office 2013 on Windows</span></span><br><span data-ttu-id="b98d9-217">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b98d9-218">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-218">
        - TaskPane</span></span><br><span data-ttu-id="b98d9-219">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="b98d9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b98d9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b98d9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-222">
        - BindingEvents</span></span><br><span data-ttu-id="b98d9-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-223">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-224">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-225">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-225">
        - File</span></span><br><span data-ttu-id="b98d9-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-226">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-228">
        - Selection</span></span><br><span data-ttu-id="b98d9-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-229">
        - Settings</span></span><br><span data-ttu-id="b98d9-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-230">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-231">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-232">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-234">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b98d9-234">Office on iPad</span></span><br><span data-ttu-id="b98d9-235">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b98d9-236">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-236">- TaskPane</span></span><br><span data-ttu-id="b98d9-237">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-237">
        - Content</span></span></td>
    <td><span data-ttu-id="b98d9-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b98d9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-249">- BindingEvents</span></span><br><span data-ttu-id="b98d9-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-250">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-251">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-251">
        - File</span></span><br><span data-ttu-id="b98d9-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-252">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-254">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-254">
        - Selection</span></span><br><span data-ttu-id="b98d9-255">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-255">
        - Settings</span></span><br><span data-ttu-id="b98d9-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-256">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-257">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-258">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-260">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-260">Office on Mac</span></span><br><span data-ttu-id="b98d9-261">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b98d9-262">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-262">- TaskPane</span></span><br><span data-ttu-id="b98d9-263">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-263">
        - Content</span></span><br><span data-ttu-id="b98d9-264">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-264">
        - Custom Functions</span></span><br><span data-ttu-id="b98d9-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b98d9-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="b98d9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="b98d9-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-278">- BindingEvents</span></span><br><span data-ttu-id="b98d9-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-279">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-280">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-281">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-281">
        - File</span></span><br><span data-ttu-id="b98d9-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-282">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-284">
        - PdfFile</span></span><br><span data-ttu-id="b98d9-285">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-285">
        - Selection</span></span><br><span data-ttu-id="b98d9-286">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-286">
        - Settings</span></span><br><span data-ttu-id="b98d9-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-287">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-288">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-289">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-291">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-291">Office 2019 on Mac</span></span><br><span data-ttu-id="b98d9-292">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b98d9-293">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-293">- TaskPane</span></span><br><span data-ttu-id="b98d9-294">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-294">
        - Content</span></span><br><span data-ttu-id="b98d9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="b98d9-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="b98d9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="b98d9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="b98d9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="b98d9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="b98d9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="b98d9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="b98d9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-306">- BindingEvents</span></span><br><span data-ttu-id="b98d9-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-307">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-308">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-309">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-309">
        - File</span></span><br><span data-ttu-id="b98d9-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-310">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-312">
        - PdfFile</span></span><br><span data-ttu-id="b98d9-313">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-313">
        - Selection</span></span><br><span data-ttu-id="b98d9-314">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-314">
        - Settings</span></span><br><span data-ttu-id="b98d9-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-315">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-316">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-317">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-319">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-319">Office 2016 on Mac</span></span><br><span data-ttu-id="b98d9-320">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="b98d9-321">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-321">- TaskPane</span></span><br><span data-ttu-id="b98d9-322">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-322">
        - Content</span></span></td>
    <td><span data-ttu-id="b98d9-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="b98d9-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b98d9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="b98d9-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-326">- BindingEvents</span></span><br><span data-ttu-id="b98d9-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-327">
        - CompressedFile</span></span><br><span data-ttu-id="b98d9-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-328">
        - DocumentEvents</span></span><br><span data-ttu-id="b98d9-329">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-329">
        - File</span></span><br><span data-ttu-id="b98d9-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-330">
        - MatrixBindings</span></span><br><span data-ttu-id="b98d9-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-332">
        - PdfFile</span></span><br><span data-ttu-id="b98d9-333">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-333">
        - Selection</span></span><br><span data-ttu-id="b98d9-334">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-334">
        - Settings</span></span><br><span data-ttu-id="b98d9-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-335">
        - TableBindings</span></span><br><span data-ttu-id="b98d9-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-336">
        - TableCoercion</span></span><br><span data-ttu-id="b98d9-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-337">
        - TextBindings</span></span><br><span data-ttu-id="b98d9-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b98d9-339">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b98d9-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="b98d9-340">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="b98d9-341">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="b98d9-342">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="b98d9-343">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="b98d9-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-345">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-345">Office on the web</span></span></td>
    <td><span data-ttu-id="b98d9-346">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b98d9-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-348">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-348">Office on Windows</span></span><br><span data-ttu-id="b98d9-349">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="b98d9-350">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b98d9-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-352">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-352">Office for Mac</span></span><br><span data-ttu-id="b98d9-353">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="b98d9-354">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="b98d9-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="b98d9-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="b98d9-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="b98d9-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b98d9-357">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-357">Platform</span></span></th>
    <th><span data-ttu-id="b98d9-358">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-358">Extension points</span></span></th>
    <th><span data-ttu-id="b98d9-359">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="b98d9-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-361">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-361">Office on the web</span></span><br><span data-ttu-id="b98d9-362">(moderno)</span><span class="sxs-lookup"><span data-stu-id="b98d9-362">(modern)</span></span></td>
    <td> <span data-ttu-id="b98d9-363">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-363">- Mail Read</span></span><br><span data-ttu-id="b98d9-364">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-364">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b98d9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b98d9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b98d9-374">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-375">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-375">Office on the web</span></span><br><span data-ttu-id="b98d9-376">(clássico)</span><span class="sxs-lookup"><span data-stu-id="b98d9-376">(classic)</span></span></td>
    <td> <span data-ttu-id="b98d9-377">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-377">- Mail Read</span></span><br><span data-ttu-id="b98d9-378">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-378">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b98d9-386">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-387">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-387">Office on Windows</span></span><br><span data-ttu-id="b98d9-388">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-389">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-389">- Mail Read</span></span><br><span data-ttu-id="b98d9-390">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-390">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b98d9-392">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b98d9-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b98d9-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b98d9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b98d9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b98d9-401">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-401">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-402">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-402">Office 2019 on Windows</span></span><br><span data-ttu-id="b98d9-403">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-403">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-404">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-404">- Mail Read</span></span><br><span data-ttu-id="b98d9-405">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-405">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b98d9-407">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b98d9-407">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b98d9-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b98d9-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="b98d9-415">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-416">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-416">Office 2016 on Windows</span></span><br><span data-ttu-id="b98d9-417">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-418">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-418">- Mail Read</span></span><br><span data-ttu-id="b98d9-419">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-419">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="b98d9-421">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="b98d9-421">
      - Modules</span></span></td>
    <td> <span data-ttu-id="b98d9-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b98d9-426">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-427">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-427">Office 2013 on Windows</span></span><br><span data-ttu-id="b98d9-428">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-428">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-429">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-429">- Mail Read</span></span><br><span data-ttu-id="b98d9-430">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-430">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="b98d9-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="b98d9-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="b98d9-435">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-436">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="b98d9-436">Office on iOS</span></span><br><span data-ttu-id="b98d9-437">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-437">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-438">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-438">- Mail Read</span></span><br><span data-ttu-id="b98d9-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b98d9-445">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-446">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-446">Office on Mac</span></span><br><span data-ttu-id="b98d9-447">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-447">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-448">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-448">- Mail Read</span></span><br><span data-ttu-id="b98d9-449">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-449">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="b98d9-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="b98d9-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="b98d9-459">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-460">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-460">Office 2019 on Mac</span></span><br><span data-ttu-id="b98d9-461">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-462">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-462">- Mail Read</span></span><br><span data-ttu-id="b98d9-463">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-463">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b98d9-471">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-472">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-472">Office 2016 on Mac</span></span><br><span data-ttu-id="b98d9-473">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-474">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-474">- Mail Read</span></span><br><span data-ttu-id="b98d9-475">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-475">
      - Mail Compose</span></span><br><span data-ttu-id="b98d9-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="b98d9-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="b98d9-483">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-484">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="b98d9-484">Office on Android</span></span><br><span data-ttu-id="b98d9-485">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-486">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="b98d9-486">- Mail Read</span></span><br><span data-ttu-id="b98d9-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="b98d9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="b98d9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="b98d9-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="b98d9-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="b98d9-493">Não disponível</span><span class="sxs-lookup"><span data-stu-id="b98d9-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="b98d9-494">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b98d9-494">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b98d9-495">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="b98d9-495">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="b98d9-496">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="b98d9-496">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="b98d9-497">Word</span><span class="sxs-lookup"><span data-stu-id="b98d9-497">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b98d9-498">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-498">Platform</span></span></th>
    <th><span data-ttu-id="b98d9-499">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-499">Extension points</span></span></th>
    <th><span data-ttu-id="b98d9-500">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-500">API requirement sets</span></span></th>
    <th><span data-ttu-id="b98d9-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-502">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-502">Office on the web</span></span></td>
    <td> <span data-ttu-id="b98d9-503">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-503">- TaskPane</span></span><br><span data-ttu-id="b98d9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b98d9-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-511">- BindingEvents</span></span><br><span data-ttu-id="b98d9-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-513">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-514">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-514">
         - File</span></span><br><span data-ttu-id="b98d9-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-516">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-519">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-520">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-520">
         - Selection</span></span><br><span data-ttu-id="b98d9-521">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-521">
         - Settings</span></span><br><span data-ttu-id="b98d9-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-522">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-523">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-524">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-525">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-526">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-527">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-527">Office on Windows</span></span><br><span data-ttu-id="b98d9-528">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-528">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-529">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-529">- TaskPane</span></span><br><span data-ttu-id="b98d9-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b98d9-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-537">- BindingEvents</span></span><br><span data-ttu-id="b98d9-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-538">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-540">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-541">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-541">
         - File</span></span><br><span data-ttu-id="b98d9-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-543">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-546">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-547">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-547">
         - Selection</span></span><br><span data-ttu-id="b98d9-548">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-548">
         - Settings</span></span><br><span data-ttu-id="b98d9-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-549">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-550">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-551">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-552">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-553">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-554">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-554">Office 2019 on Windows</span></span><br><span data-ttu-id="b98d9-555">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-555">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-556">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-556">- TaskPane</span></span><br><span data-ttu-id="b98d9-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-563">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-563">- BindingEvents</span></span><br><span data-ttu-id="b98d9-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-564">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-565">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-565">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-566">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-567">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-567">
         - File</span></span><br><span data-ttu-id="b98d9-568">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-568">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-569">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-569">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-570">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-570">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-571">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-571">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-572">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-573">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-573">
         - Selection</span></span><br><span data-ttu-id="b98d9-574">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-574">
         - Settings</span></span><br><span data-ttu-id="b98d9-575">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-575">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-576">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-576">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-577">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-577">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-578">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-578">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-579">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-579">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-580">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-580">Office 2016 on Windows</span></span><br><span data-ttu-id="b98d9-581">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-581">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-582">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-582">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b98d9-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-586">- BindingEvents</span></span><br><span data-ttu-id="b98d9-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-587">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-589">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-590">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-590">
         - File</span></span><br><span data-ttu-id="b98d9-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-592">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-595">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-596">
         - Selection</span></span><br><span data-ttu-id="b98d9-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-597">
         - Settings</span></span><br><span data-ttu-id="b98d9-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-598">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-599">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-600">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-601">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-603">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-603">Office 2013 on Windows</span></span><br><span data-ttu-id="b98d9-604">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-605">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b98d9-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b98d9-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-608">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-608">- BindingEvents</span></span><br><span data-ttu-id="b98d9-609">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-609">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-610">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-610">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-611">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-611">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-612">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-612">
         - File</span></span><br><span data-ttu-id="b98d9-613">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-613">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-614">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-614">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-615">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-615">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-616">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-616">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-617">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-617">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-618">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-618">
         - Selection</span></span><br><span data-ttu-id="b98d9-619">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-619">
         - Settings</span></span><br><span data-ttu-id="b98d9-620">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-620">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-621">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-621">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-622">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-622">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-623">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-624">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-624">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-625">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b98d9-625">Office on iPad</span></span><br><span data-ttu-id="b98d9-626">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-626">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-627">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-627">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b98d9-633">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-633">- BindingEvents</span></span><br><span data-ttu-id="b98d9-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-634">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-635">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-635">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-636">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-636">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-637">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-637">
         - File</span></span><br><span data-ttu-id="b98d9-638">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-638">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-639">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-639">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-640">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-640">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-641">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-641">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-642">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-642">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-643">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-643">
         - Selection</span></span><br><span data-ttu-id="b98d9-644">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-644">
         - Settings</span></span><br><span data-ttu-id="b98d9-645">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-645">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-646">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-646">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-647">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-647">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-648">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-648">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-649">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-649">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-650">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-650">Office on Mac</span></span><br><span data-ttu-id="b98d9-651">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-651">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-652">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-652">- TaskPane</span></span><br><span data-ttu-id="b98d9-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="b98d9-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-660">- BindingEvents</span></span><br><span data-ttu-id="b98d9-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-661">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-663">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-664">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-664">
         - File</span></span><br><span data-ttu-id="b98d9-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-666">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-669">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-670">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-670">
         - Selection</span></span><br><span data-ttu-id="b98d9-671">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-671">
         - Settings</span></span><br><span data-ttu-id="b98d9-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-672">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-673">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-674">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-675">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-677">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-677">Office 2019 on Mac</span></span><br><span data-ttu-id="b98d9-678">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-678">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-679">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-679">- TaskPane</span></span><br><span data-ttu-id="b98d9-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="b98d9-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="b98d9-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="b98d9-686">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-686">- BindingEvents</span></span><br><span data-ttu-id="b98d9-687">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-687">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-688">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-688">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-689">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-689">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-690">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-690">
         - File</span></span><br><span data-ttu-id="b98d9-691">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-691">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-692">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-692">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-693">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-693">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-694">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-694">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-695">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-695">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-696">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-696">
         - Selection</span></span><br><span data-ttu-id="b98d9-697">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-697">
         - Settings</span></span><br><span data-ttu-id="b98d9-698">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-698">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-699">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-699">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-700">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-700">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-701">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-702">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-702">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-703">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-703">Office 2016 on Mac</span></span><br><span data-ttu-id="b98d9-704">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-704">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-705">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-705">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="b98d9-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="b98d9-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="b98d9-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-709">- BindingEvents</span></span><br><span data-ttu-id="b98d9-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-710">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="b98d9-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="b98d9-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-712">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-713">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-713">
         - File</span></span><br><span data-ttu-id="b98d9-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-715">
         - MatrixBindings</span></span><br><span data-ttu-id="b98d9-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="b98d9-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="b98d9-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-718">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-719">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-719">
         - Selection</span></span><br><span data-ttu-id="b98d9-720">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-720">
         - Settings</span></span><br><span data-ttu-id="b98d9-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-721">
         - TableBindings</span></span><br><span data-ttu-id="b98d9-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-722">
         - TableCoercion</span></span><br><span data-ttu-id="b98d9-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="b98d9-723">
         - TextBindings</span></span><br><span data-ttu-id="b98d9-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-724">
         - TextCoercion</span></span><br><span data-ttu-id="b98d9-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-725">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="b98d9-726">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b98d9-726">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="b98d9-727">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="b98d9-727">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b98d9-728">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-728">Platform</span></span></th>
    <th><span data-ttu-id="b98d9-729">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-729">Extension points</span></span></th>
    <th><span data-ttu-id="b98d9-730">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-730">API requirement sets</span></span></th>
    <th><span data-ttu-id="b98d9-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-732">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-732">Office on the web</span></span></td>
    <td> <span data-ttu-id="b98d9-733">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-733">- Content</span></span><br><span data-ttu-id="b98d9-734">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-734">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b98d9-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b98d9-740">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-740">- ActiveView</span></span><br><span data-ttu-id="b98d9-741">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-741">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-742">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-742">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-743">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-743">
         - File</span></span><br><span data-ttu-id="b98d9-744">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-744">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-745">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-745">
         - Selection</span></span><br><span data-ttu-id="b98d9-746">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-746">
         - Settings</span></span><br><span data-ttu-id="b98d9-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-747">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-748">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-748">Office on Windows</span></span><br><span data-ttu-id="b98d9-749">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-749">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-750">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-750">- Content</span></span><br><span data-ttu-id="b98d9-751">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-751">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b98d9-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b98d9-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-757">- ActiveView</span></span><br><span data-ttu-id="b98d9-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-758">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-759">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-760">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-760">
         - File</span></span><br><span data-ttu-id="b98d9-761">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-761">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-762">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-762">
         - Selection</span></span><br><span data-ttu-id="b98d9-763">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-763">
         - Settings</span></span><br><span data-ttu-id="b98d9-764">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-764">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-765">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-765">Office 2019 on Windows</span></span><br><span data-ttu-id="b98d9-766">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-766">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-767">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-767">- Content</span></span><br><span data-ttu-id="b98d9-768">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-768">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-772">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-772">- ActiveView</span></span><br><span data-ttu-id="b98d9-773">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-773">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-774">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-774">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-775">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-775">
         - File</span></span><br><span data-ttu-id="b98d9-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-776">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-777">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-777">
         - Selection</span></span><br><span data-ttu-id="b98d9-778">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-778">
         - Settings</span></span><br><span data-ttu-id="b98d9-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-780">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-780">Office 2016 on Windows</span></span><br><span data-ttu-id="b98d9-781">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-782">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-782">- Content</span></span><br><span data-ttu-id="b98d9-783">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-783">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b98d9-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b98d9-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-786">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-786">- ActiveView</span></span><br><span data-ttu-id="b98d9-787">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-787">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-788">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-788">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-789">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-789">
         - File</span></span><br><span data-ttu-id="b98d9-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-790">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-791">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-791">
         - Selection</span></span><br><span data-ttu-id="b98d9-792">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-792">
         - Settings</span></span><br><span data-ttu-id="b98d9-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-794">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-794">Office 2013 on Windows</span></span><br><span data-ttu-id="b98d9-795">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-795">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-796">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-796">- Content</span></span><br><span data-ttu-id="b98d9-797">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-797">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="b98d9-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b98d9-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b98d9-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-800">- ActiveView</span></span><br><span data-ttu-id="b98d9-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-801">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-802">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-803">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-803">
         - File</span></span><br><span data-ttu-id="b98d9-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-804">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-805">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-805">
         - Selection</span></span><br><span data-ttu-id="b98d9-806">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-806">
         - Settings</span></span><br><span data-ttu-id="b98d9-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-808">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="b98d9-808">Office on iPad</span></span><br><span data-ttu-id="b98d9-809">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-810">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-810">- Content</span></span><br><span data-ttu-id="b98d9-811">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b98d9-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-815">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-815">- ActiveView</span></span><br><span data-ttu-id="b98d9-816">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-816">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-817">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-817">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-818">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-818">
         - File</span></span><br><span data-ttu-id="b98d9-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-819">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-820">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-820">
         - Selection</span></span><br><span data-ttu-id="b98d9-821">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-821">
         - Settings</span></span><br><span data-ttu-id="b98d9-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-823">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-823">Office on Mac</span></span><br><span data-ttu-id="b98d9-824">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="b98d9-824">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="b98d9-825">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-825">- Content</span></span><br><span data-ttu-id="b98d9-826">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-826">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="b98d9-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="b98d9-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="b98d9-832">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-832">- ActiveView</span></span><br><span data-ttu-id="b98d9-833">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-833">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-834">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-834">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-835">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-835">
         - File</span></span><br><span data-ttu-id="b98d9-836">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-836">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-837">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-837">
         - Selection</span></span><br><span data-ttu-id="b98d9-838">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-838">
         - Settings</span></span><br><span data-ttu-id="b98d9-839">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-839">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-840">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-840">Office 2019 on Mac</span></span><br><span data-ttu-id="b98d9-841">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-841">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-842">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-842">- Content</span></span><br><span data-ttu-id="b98d9-843">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-843">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-847">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-847">- ActiveView</span></span><br><span data-ttu-id="b98d9-848">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-848">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-849">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-849">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-850">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-850">
         - File</span></span><br><span data-ttu-id="b98d9-851">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-851">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-852">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-852">
         - Selection</span></span><br><span data-ttu-id="b98d9-853">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-853">
         - Settings</span></span><br><span data-ttu-id="b98d9-854">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-854">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-855">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-855">Office 2016 on Mac</span></span><br><span data-ttu-id="b98d9-856">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-856">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-857">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-857">- Content</span></span><br><span data-ttu-id="b98d9-858">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-858">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="b98d9-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="b98d9-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-861">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="b98d9-861">- ActiveView</span></span><br><span data-ttu-id="b98d9-862">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-862">
         - CompressedFile</span></span><br><span data-ttu-id="b98d9-863">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-863">
         - DocumentEvents</span></span><br><span data-ttu-id="b98d9-864">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="b98d9-864">
         - File</span></span><br><span data-ttu-id="b98d9-865">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="b98d9-865">
         - PdfFile</span></span><br><span data-ttu-id="b98d9-866">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-866">
         - Selection</span></span><br><span data-ttu-id="b98d9-867">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-867">
         - Settings</span></span><br><span data-ttu-id="b98d9-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="b98d9-869">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="b98d9-869">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="b98d9-870">OneNote</span><span class="sxs-lookup"><span data-stu-id="b98d9-870">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b98d9-871">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-871">Platform</span></span></th>
    <th><span data-ttu-id="b98d9-872">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-872">Extension points</span></span></th>
    <th><span data-ttu-id="b98d9-873">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-873">API requirement sets</span></span></th>
    <th><span data-ttu-id="b98d9-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-875">Office na Web</span><span class="sxs-lookup"><span data-stu-id="b98d9-875">Office on the web</span></span></td>
    <td> <span data-ttu-id="b98d9-876">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="b98d9-876">- Content</span></span><br><span data-ttu-id="b98d9-877">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-877">
         - TaskPane</span></span><br><span data-ttu-id="b98d9-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="b98d9-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="b98d9-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="b98d9-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-882">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="b98d9-882">- DocumentEvents</span></span><br><span data-ttu-id="b98d9-883">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-883">
         - HtmlCoercion</span></span><br><span data-ttu-id="b98d9-884">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="b98d9-884">
         - Settings</span></span><br><span data-ttu-id="b98d9-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-885">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="b98d9-886">Project</span><span class="sxs-lookup"><span data-stu-id="b98d9-886">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="b98d9-887">Plataforma</span><span class="sxs-lookup"><span data-stu-id="b98d9-887">Platform</span></span></th>
    <th><span data-ttu-id="b98d9-888">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="b98d9-888">Extension points</span></span></th>
    <th><span data-ttu-id="b98d9-889">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-889">API requirement sets</span></span></th>
    <th><span data-ttu-id="b98d9-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="b98d9-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-891">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-891">Office 2019 on Windows</span></span><br><span data-ttu-id="b98d9-892">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-893">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-895">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-895">- Selection</span></span><br><span data-ttu-id="b98d9-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-897">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-897">Office 2016 on Windows</span></span><br><span data-ttu-id="b98d9-898">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-899">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-901">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-901">- Selection</span></span><br><span data-ttu-id="b98d9-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-902">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="b98d9-903">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="b98d9-903">Office 2013 on Windows</span></span><br><span data-ttu-id="b98d9-904">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="b98d9-904">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="b98d9-905">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="b98d9-905">- TaskPane</span></span></td>
    <td> <span data-ttu-id="b98d9-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="b98d9-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="b98d9-907">- Seleção</span><span class="sxs-lookup"><span data-stu-id="b98d9-907">- Selection</span></span><br><span data-ttu-id="b98d9-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="b98d9-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="b98d9-909">Confira também</span><span class="sxs-lookup"><span data-stu-id="b98d9-909">See also</span></span>

- [<span data-ttu-id="b98d9-910">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="b98d9-910">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="b98d9-911">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="b98d9-911">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="b98d9-912">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="b98d9-912">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="b98d9-913">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="b98d9-913">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="b98d9-914">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="b98d9-914">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="b98d9-915">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="b98d9-915">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="b98d9-916">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="b98d9-916">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="b98d9-917">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="b98d9-917">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="b98d9-918">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b98d9-918">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="b98d9-919">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="b98d9-919">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="b98d9-920">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="b98d9-920">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
