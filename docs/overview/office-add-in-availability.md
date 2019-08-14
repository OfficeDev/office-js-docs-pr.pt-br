---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: 1e368fe21a1bcdb2a7f44c88ce8e881605fa96f2
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395649"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5eb81-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5eb81-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5eb81-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="5eb81-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="5eb81-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="5eb81-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="5eb81-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="5eb81-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="5eb81-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="5eb81-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="5eb81-108">Excel</span><span class="sxs-lookup"><span data-stu-id="5eb81-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5eb81-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5eb81-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5eb81-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5eb81-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="5eb81-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-114">- TaskPane</span></span><br><span data-ttu-id="5eb81-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-115">
        - Content</span></span><br><span data-ttu-id="5eb81-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-116">
        - Custom Functions</span></span><br><span data-ttu-id="5eb81-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="5eb81-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5eb81-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5eb81-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-128">
        - BindingEvents</span></span><br><span data-ttu-id="5eb81-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-129">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-130">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-131">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-131">
        - File</span></span><br><span data-ttu-id="5eb81-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-132">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-134">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-134">
        - Selection</span></span><br><span data-ttu-id="5eb81-135">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-135">
        - Settings</span></span><br><span data-ttu-id="5eb81-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-136">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-137">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-138">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-140">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-140">Office on Windows</span></span><br><span data-ttu-id="5eb81-141">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-142">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-142">- TaskPane</span></span><br><span data-ttu-id="5eb81-143">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-143">
        - Content</span></span><br><span data-ttu-id="5eb81-144">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-144">
        - Custom Functions</span></span><br><span data-ttu-id="5eb81-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="5eb81-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5eb81-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5eb81-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="5eb81-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-158">
        - BindingEvents</span></span><br><span data-ttu-id="5eb81-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-159">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-160">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-161">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-161">
        - File</span></span><br><span data-ttu-id="5eb81-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-162">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-164">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-164">
        - Selection</span></span><br><span data-ttu-id="5eb81-165">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-165">
        - Settings</span></span><br><span data-ttu-id="5eb81-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-166">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-167">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-168">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-170">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-170">Office 2019 on Windows</span></span><br><span data-ttu-id="5eb81-171">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5eb81-172">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-172">- TaskPane</span></span><br><span data-ttu-id="5eb81-173">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-173">
        - Content</span></span><br><span data-ttu-id="5eb81-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5eb81-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-185">- BindingEvents</span></span><br><span data-ttu-id="5eb81-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-186">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-187">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-188">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-188">
        - File</span></span><br><span data-ttu-id="5eb81-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-189">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-191">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-191">
        - Selection</span></span><br><span data-ttu-id="5eb81-192">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-192">
        - Settings</span></span><br><span data-ttu-id="5eb81-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-193">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-194">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-195">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-197">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-197">Office 2016 on Windows</span></span><br><span data-ttu-id="5eb81-198">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5eb81-199">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-199">- TaskPane</span></span><br><span data-ttu-id="5eb81-200">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-200">
        - Content</span></span></td>
    <td><span data-ttu-id="5eb81-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5eb81-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-204">- BindingEvents</span></span><br><span data-ttu-id="5eb81-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-205">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-206">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-207">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-207">
        - File</span></span><br><span data-ttu-id="5eb81-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-208">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-210">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-210">
        - Selection</span></span><br><span data-ttu-id="5eb81-211">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-211">
        - Settings</span></span><br><span data-ttu-id="5eb81-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-212">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-213">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-214">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-216">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-216">Office 2013 on Windows</span></span><br><span data-ttu-id="5eb81-217">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5eb81-218">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-218">
        - TaskPane</span></span><br><span data-ttu-id="5eb81-219">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5eb81-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5eb81-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5eb81-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-222">
        - BindingEvents</span></span><br><span data-ttu-id="5eb81-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-223">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-224">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-225">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-225">
        - File</span></span><br><span data-ttu-id="5eb81-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-226">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-228">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-228">
        - Selection</span></span><br><span data-ttu-id="5eb81-229">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-229">
        - Settings</span></span><br><span data-ttu-id="5eb81-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-230">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-231">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-232">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-234">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="5eb81-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="5eb81-235">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5eb81-236">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-236">- TaskPane</span></span><br><span data-ttu-id="5eb81-237">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-237">
        - Content</span></span><br><span data-ttu-id="5eb81-238">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-238">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5eb81-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5eb81-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-250">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-250">- BindingEvents</span></span><br><span data-ttu-id="5eb81-251">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-251">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-252">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-252">
        - File</span></span><br><span data-ttu-id="5eb81-253">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-253">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-254">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-254">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-255">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-255">
        - Selection</span></span><br><span data-ttu-id="5eb81-256">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-256">
        - Settings</span></span><br><span data-ttu-id="5eb81-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-257">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-258">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-259">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-260">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-261">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-261">Office apps on Mac</span></span><br><span data-ttu-id="5eb81-262">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-262">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5eb81-263">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-263">- TaskPane</span></span><br><span data-ttu-id="5eb81-264">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-264">
        - Content</span></span><br><span data-ttu-id="5eb81-265">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-265">
        - Custom Functions</span></span><br><span data-ttu-id="5eb81-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5eb81-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5eb81-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="5eb81-279">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-279">- BindingEvents</span></span><br><span data-ttu-id="5eb81-280">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-280">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-281">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-281">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-282">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-282">
        - File</span></span><br><span data-ttu-id="5eb81-283">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-283">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-284">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-284">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-285">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-285">
        - PdfFile</span></span><br><span data-ttu-id="5eb81-286">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-286">
        - Selection</span></span><br><span data-ttu-id="5eb81-287">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-287">
        - Settings</span></span><br><span data-ttu-id="5eb81-288">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-288">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-289">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-289">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-290">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-290">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-291">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-291">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-292">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-292">Office 2019 for Mac</span></span><br><span data-ttu-id="5eb81-293">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-293">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5eb81-294">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-294">- TaskPane</span></span><br><span data-ttu-id="5eb81-295">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-295">
        - Content</span></span><br><span data-ttu-id="5eb81-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5eb81-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5eb81-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5eb81-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5eb81-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5eb81-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5eb81-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5eb81-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5eb81-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-307">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-307">- BindingEvents</span></span><br><span data-ttu-id="5eb81-308">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-308">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-309">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-309">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-310">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-310">
        - File</span></span><br><span data-ttu-id="5eb81-311">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-311">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-312">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-312">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-313">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-313">
        - PdfFile</span></span><br><span data-ttu-id="5eb81-314">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-314">
        - Selection</span></span><br><span data-ttu-id="5eb81-315">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-315">
        - Settings</span></span><br><span data-ttu-id="5eb81-316">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-316">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-317">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-317">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-318">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-318">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-319">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-319">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-320">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-320">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="5eb81-321">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-321">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5eb81-322">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-322">- TaskPane</span></span><br><span data-ttu-id="5eb81-323">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-323">
        - Content</span></span></td>
    <td><span data-ttu-id="5eb81-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5eb81-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5eb81-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="5eb81-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-327">- BindingEvents</span></span><br><span data-ttu-id="5eb81-328">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-328">
        - CompressedFile</span></span><br><span data-ttu-id="5eb81-329">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-329">
        - DocumentEvents</span></span><br><span data-ttu-id="5eb81-330">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-330">
        - File</span></span><br><span data-ttu-id="5eb81-331">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-331">
        - MatrixBindings</span></span><br><span data-ttu-id="5eb81-332">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-332">
        - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-333">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-333">
        - PdfFile</span></span><br><span data-ttu-id="5eb81-334">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-334">
        - Selection</span></span><br><span data-ttu-id="5eb81-335">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-335">
        - Settings</span></span><br><span data-ttu-id="5eb81-336">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-336">
        - TableBindings</span></span><br><span data-ttu-id="5eb81-337">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-337">
        - TableCoercion</span></span><br><span data-ttu-id="5eb81-338">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-338">
        - TextBindings</span></span><br><span data-ttu-id="5eb81-339">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-339">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5eb81-340">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="5eb81-340">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="5eb81-341">Funções Personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-341">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5eb81-342">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-342">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5eb81-343">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-343">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5eb81-344">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-344">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5eb81-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-346">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-346">Office on the web</span></span></td>
    <td><span data-ttu-id="5eb81-347">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-347">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5eb81-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-349">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-349">Office on Windows</span></span><br><span data-ttu-id="5eb81-350">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-350">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="5eb81-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5eb81-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-353">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-353">Office for Mac</span></span><br><span data-ttu-id="5eb81-354">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-354">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5eb81-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="5eb81-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5eb81-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="5eb81-357">Outlook</span><span class="sxs-lookup"><span data-stu-id="5eb81-357">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5eb81-358">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-358">Platform</span></span></th>
    <th><span data-ttu-id="5eb81-359">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-359">Extension points</span></span></th>
    <th><span data-ttu-id="5eb81-360">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-360">API requirement sets</span></span></th>
    <th><span data-ttu-id="5eb81-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-362">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-362">Office on the web</span></span><br><span data-ttu-id="5eb81-363">(moderno)</span><span class="sxs-lookup"><span data-stu-id="5eb81-363">Modern</span></span></td>
    <td> <span data-ttu-id="5eb81-364">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-364">- Mail Read</span></span><br><span data-ttu-id="5eb81-365">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-365">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5eb81-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5eb81-374">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-375">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-375">Office on the web</span></span><br><span data-ttu-id="5eb81-376">(clássico)</span><span class="sxs-lookup"><span data-stu-id="5eb81-376">Classic.</span></span></td>
    <td> <span data-ttu-id="5eb81-377">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-377">- Mail Read</span></span><br><span data-ttu-id="5eb81-378">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-378">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5eb81-386">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-387">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-387">Office on Windows</span></span><br><span data-ttu-id="5eb81-388">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-389">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-389">- Mail Read</span></span><br><span data-ttu-id="5eb81-390">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-390">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5eb81-392">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="5eb81-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5eb81-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5eb81-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5eb81-400">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-400">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-401">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-401">Office 2019 on Windows</span></span><br><span data-ttu-id="5eb81-402">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-402">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-403">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-403">- Mail Read</span></span><br><span data-ttu-id="5eb81-404">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-404">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5eb81-406">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="5eb81-406">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5eb81-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5eb81-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5eb81-414">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-415">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-415">Office 2016 on Windows</span></span><br><span data-ttu-id="5eb81-416">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-416">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-417">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-417">- Mail Read</span></span><br><span data-ttu-id="5eb81-418">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-418">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5eb81-420">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="5eb81-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5eb81-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5eb81-425">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-426">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-426">Office 2013 on Windows</span></span><br><span data-ttu-id="5eb81-427">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-427">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-428">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-428">- Mail Read</span></span><br><span data-ttu-id="5eb81-429">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-429">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="5eb81-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="5eb81-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5eb81-434">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-434">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-435">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="5eb81-435">Office apps on iOS</span></span><br><span data-ttu-id="5eb81-436">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-436">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-437">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-437">- Mail Read</span></span><br><span data-ttu-id="5eb81-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5eb81-444">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-445">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-445">Office apps on Mac</span></span><br><span data-ttu-id="5eb81-446">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-446">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-447">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-447">- Mail Read</span></span><br><span data-ttu-id="5eb81-448">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-448">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5eb81-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5eb81-457">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-458">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-458">Office 2019 for Mac</span></span><br><span data-ttu-id="5eb81-459">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-460">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-460">- Mail Read</span></span><br><span data-ttu-id="5eb81-461">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-461">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5eb81-469">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-470">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-470">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="5eb81-471">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-471">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-472">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-472">- Mail Read</span></span><br><span data-ttu-id="5eb81-473">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-473">
      - Mail Compose</span></span><br><span data-ttu-id="5eb81-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5eb81-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5eb81-481">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-482">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="5eb81-482">Office apps on Android</span></span><br><span data-ttu-id="5eb81-483">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-483">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-484">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="5eb81-484">- Mail Read</span></span><br><span data-ttu-id="5eb81-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5eb81-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5eb81-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5eb81-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5eb81-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5eb81-491">Não disponível</span><span class="sxs-lookup"><span data-stu-id="5eb81-491">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="5eb81-492">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="5eb81-492">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="5eb81-493">Word</span><span class="sxs-lookup"><span data-stu-id="5eb81-493">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5eb81-494">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-494">Platform</span></span></th>
    <th><span data-ttu-id="5eb81-495">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-495">Extension points</span></span></th>
    <th><span data-ttu-id="5eb81-496">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-496">API requirement sets</span></span></th>
    <th><span data-ttu-id="5eb81-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-498">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-498">Office on the web</span></span></td>
    <td> <span data-ttu-id="5eb81-499">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-499">- TaskPane</span></span><br><span data-ttu-id="5eb81-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5eb81-507">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-507">- BindingEvents</span></span><br><span data-ttu-id="5eb81-508">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-508">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-509">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-509">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-510">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-510">
         - File</span></span><br><span data-ttu-id="5eb81-511">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-511">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-512">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-512">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-513">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-513">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-514">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-514">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-515">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-515">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-516">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-516">
         - Selection</span></span><br><span data-ttu-id="5eb81-517">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-517">
         - Settings</span></span><br><span data-ttu-id="5eb81-518">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-518">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-519">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-519">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-520">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-520">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-521">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-521">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-522">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-522">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-523">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-523">Office on Windows</span></span><br><span data-ttu-id="5eb81-524">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-524">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-525">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-525">- TaskPane</span></span><br><span data-ttu-id="5eb81-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5eb81-533">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-533">- BindingEvents</span></span><br><span data-ttu-id="5eb81-534">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-534">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-536">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-537">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-537">
         - File</span></span><br><span data-ttu-id="5eb81-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-539">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-542">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-543">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-543">
         - Selection</span></span><br><span data-ttu-id="5eb81-544">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-544">
         - Settings</span></span><br><span data-ttu-id="5eb81-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-545">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-546">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-547">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-548">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-549">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-550">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-550">Office 2019 on Windows</span></span><br><span data-ttu-id="5eb81-551">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-551">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-552">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-552">- TaskPane</span></span><br><span data-ttu-id="5eb81-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-559">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-559">- BindingEvents</span></span><br><span data-ttu-id="5eb81-560">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-560">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-561">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-561">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-562">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-562">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-563">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-563">
         - File</span></span><br><span data-ttu-id="5eb81-564">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-564">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-565">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-565">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-566">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-566">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-567">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-567">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-568">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-569">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-569">
         - Selection</span></span><br><span data-ttu-id="5eb81-570">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-570">
         - Settings</span></span><br><span data-ttu-id="5eb81-571">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-571">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-572">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-572">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-573">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-573">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-574">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-574">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-575">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-575">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-576">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-576">Office 2016 on Windows</span></span><br><span data-ttu-id="5eb81-577">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-577">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-578">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-578">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5eb81-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-582">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-582">- BindingEvents</span></span><br><span data-ttu-id="5eb81-583">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-583">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-584">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-584">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-585">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-586">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-586">
         - File</span></span><br><span data-ttu-id="5eb81-587">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-587">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-588">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-588">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-589">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-589">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-590">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-590">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-591">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-591">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-592">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-592">
         - Selection</span></span><br><span data-ttu-id="5eb81-593">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-593">
         - Settings</span></span><br><span data-ttu-id="5eb81-594">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-594">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-595">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-595">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-596">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-596">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-597">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-598">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-598">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-599">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-599">Office 2013 on Windows</span></span><br><span data-ttu-id="5eb81-600">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-600">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-601">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-601">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5eb81-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5eb81-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-604">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-604">- BindingEvents</span></span><br><span data-ttu-id="5eb81-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-605">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-606">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-606">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-607">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-608">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-608">
         - File</span></span><br><span data-ttu-id="5eb81-609">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-609">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-610">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-610">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-611">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-611">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-612">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-612">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-613">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-613">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-614">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-614">
         - Selection</span></span><br><span data-ttu-id="5eb81-615">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-615">
         - Settings</span></span><br><span data-ttu-id="5eb81-616">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-616">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-617">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-617">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-618">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-618">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-619">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-619">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-620">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-620">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-621">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="5eb81-621">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="5eb81-622">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-622">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-623">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-623">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="5eb81-629">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-629">- BindingEvents</span></span><br><span data-ttu-id="5eb81-630">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-630">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-631">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-631">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-632">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-633">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-633">
         - File</span></span><br><span data-ttu-id="5eb81-634">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-634">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-635">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-635">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-636">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-636">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-637">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-637">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-638">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-639">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-639">
         - Selection</span></span><br><span data-ttu-id="5eb81-640">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-640">
         - Settings</span></span><br><span data-ttu-id="5eb81-641">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-641">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-642">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-642">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-643">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-643">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-644">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-644">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-645">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-645">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-646">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-646">Office apps on Mac</span></span><br><span data-ttu-id="5eb81-647">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-647">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-648">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-648">- TaskPane</span></span><br><span data-ttu-id="5eb81-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="5eb81-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-656">- BindingEvents</span></span><br><span data-ttu-id="5eb81-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-657">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-659">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-660">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-660">
         - File</span></span><br><span data-ttu-id="5eb81-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-662">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-665">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-666">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-666">
         - Selection</span></span><br><span data-ttu-id="5eb81-667">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-667">
         - Settings</span></span><br><span data-ttu-id="5eb81-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-668">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-669">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-670">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-671">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-673">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-673">Office 2019 for Mac</span></span><br><span data-ttu-id="5eb81-674">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-674">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-675">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-675">- TaskPane</span></span><br><span data-ttu-id="5eb81-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="5eb81-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="5eb81-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="5eb81-682">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-682">- BindingEvents</span></span><br><span data-ttu-id="5eb81-683">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-683">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-684">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-684">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-685">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-685">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-686">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-686">
         - File</span></span><br><span data-ttu-id="5eb81-687">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-687">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-688">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-688">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-689">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-689">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-690">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-690">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-691">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-691">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-692">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-692">
         - Selection</span></span><br><span data-ttu-id="5eb81-693">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-693">
         - Settings</span></span><br><span data-ttu-id="5eb81-694">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-694">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-695">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-695">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-696">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-696">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-697">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-697">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-698">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-698">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-699">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-699">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="5eb81-700">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-700">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-701">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-701">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="5eb81-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5eb81-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="5eb81-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-705">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-705">- BindingEvents</span></span><br><span data-ttu-id="5eb81-706">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-706">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-707">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5eb81-707">
         - CustomXmlParts</span></span><br><span data-ttu-id="5eb81-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-708">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-709">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-709">
         - File</span></span><br><span data-ttu-id="5eb81-710">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-710">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-711">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-711">
         - MatrixBindings</span></span><br><span data-ttu-id="5eb81-712">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-712">
         - MatrixCoercion</span></span><br><span data-ttu-id="5eb81-713">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-713">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5eb81-714">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-714">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-715">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-715">
         - Selection</span></span><br><span data-ttu-id="5eb81-716">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-716">
         - Settings</span></span><br><span data-ttu-id="5eb81-717">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-717">
         - TableBindings</span></span><br><span data-ttu-id="5eb81-718">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-718">
         - TableCoercion</span></span><br><span data-ttu-id="5eb81-719">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5eb81-719">
         - TextBindings</span></span><br><span data-ttu-id="5eb81-720">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-720">
         - TextCoercion</span></span><br><span data-ttu-id="5eb81-721">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-721">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="5eb81-722">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="5eb81-722">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5eb81-723">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5eb81-723">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5eb81-724">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-724">Platform</span></span></th>
    <th><span data-ttu-id="5eb81-725">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-725">Extension points</span></span></th>
    <th><span data-ttu-id="5eb81-726">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-726">API requirement sets</span></span></th>
    <th><span data-ttu-id="5eb81-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-728">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-728">Office on the web</span></span></td>
    <td> <span data-ttu-id="5eb81-729">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-729">- Content</span></span><br><span data-ttu-id="5eb81-730">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-730">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5eb81-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5eb81-736">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-736">- ActiveView</span></span><br><span data-ttu-id="5eb81-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-737">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-738">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-738">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-739">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-739">
         - File</span></span><br><span data-ttu-id="5eb81-740">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-740">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-741">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-741">
         - Selection</span></span><br><span data-ttu-id="5eb81-742">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-742">
         - Settings</span></span><br><span data-ttu-id="5eb81-743">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-743">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-744">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-744">Office on Windows</span></span><br><span data-ttu-id="5eb81-745">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-745">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-746">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-746">- Content</span></span><br><span data-ttu-id="5eb81-747">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-747">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5eb81-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5eb81-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-753">- ActiveView</span></span><br><span data-ttu-id="5eb81-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-754">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-755">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-756">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-756">
         - File</span></span><br><span data-ttu-id="5eb81-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-757">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-758">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-758">
         - Selection</span></span><br><span data-ttu-id="5eb81-759">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-759">
         - Settings</span></span><br><span data-ttu-id="5eb81-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-761">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-761">Office 2019 on Windows</span></span><br><span data-ttu-id="5eb81-762">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-763">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-763">- Content</span></span><br><span data-ttu-id="5eb81-764">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-764">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-768">- ActiveView</span></span><br><span data-ttu-id="5eb81-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-769">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-770">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-771">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-771">
         - File</span></span><br><span data-ttu-id="5eb81-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-772">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-773">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-773">
         - Selection</span></span><br><span data-ttu-id="5eb81-774">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-774">
         - Settings</span></span><br><span data-ttu-id="5eb81-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-776">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-776">Office 2016 on Windows</span></span><br><span data-ttu-id="5eb81-777">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-778">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-778">- Content</span></span><br><span data-ttu-id="5eb81-779">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5eb81-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5eb81-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-782">- ActiveView</span></span><br><span data-ttu-id="5eb81-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-783">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-784">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-785">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-785">
         - File</span></span><br><span data-ttu-id="5eb81-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-786">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-787">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-787">
         - Selection</span></span><br><span data-ttu-id="5eb81-788">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-788">
         - Settings</span></span><br><span data-ttu-id="5eb81-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-790">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-790">Office 2013 on Windows</span></span><br><span data-ttu-id="5eb81-791">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-792">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-792">- Content</span></span><br><span data-ttu-id="5eb81-793">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="5eb81-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5eb81-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5eb81-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-796">- ActiveView</span></span><br><span data-ttu-id="5eb81-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-797">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-798">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-799">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-799">
         - File</span></span><br><span data-ttu-id="5eb81-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-800">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-801">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-801">
         - Selection</span></span><br><span data-ttu-id="5eb81-802">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-802">
         - Settings</span></span><br><span data-ttu-id="5eb81-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-804">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="5eb81-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="5eb81-805">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-806">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-806">- Content</span></span><br><span data-ttu-id="5eb81-807">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5eb81-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-811">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-811">- ActiveView</span></span><br><span data-ttu-id="5eb81-812">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-812">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-813">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-813">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-814">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-814">
         - File</span></span><br><span data-ttu-id="5eb81-815">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-815">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-816">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-816">
         - Selection</span></span><br><span data-ttu-id="5eb81-817">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-817">
         - Settings</span></span><br><span data-ttu-id="5eb81-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-818">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-819">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-819">Office apps on Mac</span></span><br><span data-ttu-id="5eb81-820">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="5eb81-820">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="5eb81-821">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-821">- Content</span></span><br><span data-ttu-id="5eb81-822">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-822">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="5eb81-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="5eb81-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="5eb81-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-828">- ActiveView</span></span><br><span data-ttu-id="5eb81-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-829">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-830">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-831">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-831">
         - File</span></span><br><span data-ttu-id="5eb81-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-832">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-833">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-833">
         - Selection</span></span><br><span data-ttu-id="5eb81-834">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-834">
         - Settings</span></span><br><span data-ttu-id="5eb81-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-836">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-836">Office 2019 for Mac</span></span><br><span data-ttu-id="5eb81-837">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-837">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-838">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-838">- Content</span></span><br><span data-ttu-id="5eb81-839">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-839">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-843">- ActiveView</span></span><br><span data-ttu-id="5eb81-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-844">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-845">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-846">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-846">
         - File</span></span><br><span data-ttu-id="5eb81-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-847">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-848">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-848">
         - Selection</span></span><br><span data-ttu-id="5eb81-849">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-849">
         - Settings</span></span><br><span data-ttu-id="5eb81-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-851">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-851">If you're running Office 2016 on a Mac:</span></span><br><span data-ttu-id="5eb81-852">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-852">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-853">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-853">- Content</span></span><br><span data-ttu-id="5eb81-854">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-854">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5eb81-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="5eb81-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-857">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5eb81-857">- ActiveView</span></span><br><span data-ttu-id="5eb81-858">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-858">
         - CompressedFile</span></span><br><span data-ttu-id="5eb81-859">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-859">
         - DocumentEvents</span></span><br><span data-ttu-id="5eb81-860">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="5eb81-860">
         - File</span></span><br><span data-ttu-id="5eb81-861">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5eb81-861">
         - PdfFile</span></span><br><span data-ttu-id="5eb81-862">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-862">
         - Selection</span></span><br><span data-ttu-id="5eb81-863">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-863">
         - Settings</span></span><br><span data-ttu-id="5eb81-864">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-864">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5eb81-865">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="5eb81-865">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="5eb81-866">OneNote</span><span class="sxs-lookup"><span data-stu-id="5eb81-866">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5eb81-867">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-867">Platform</span></span></th>
    <th><span data-ttu-id="5eb81-868">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-868">Extension points</span></span></th>
    <th><span data-ttu-id="5eb81-869">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-869">API requirement sets</span></span></th>
    <th><span data-ttu-id="5eb81-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-871">Office na Web</span><span class="sxs-lookup"><span data-stu-id="5eb81-871">Office on the web</span></span></td>
    <td> <span data-ttu-id="5eb81-872">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="5eb81-872">- Content</span></span><br><span data-ttu-id="5eb81-873">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-873">
         - TaskPane</span></span><br><span data-ttu-id="5eb81-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5eb81-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5eb81-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="5eb81-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-878">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5eb81-878">- DocumentEvents</span></span><br><span data-ttu-id="5eb81-879">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-879">
         - HtmlCoercion</span></span><br><span data-ttu-id="5eb81-880">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="5eb81-880">
         - Settings</span></span><br><span data-ttu-id="5eb81-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="5eb81-882">Project</span><span class="sxs-lookup"><span data-stu-id="5eb81-882">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5eb81-883">Plataforma</span><span class="sxs-lookup"><span data-stu-id="5eb81-883">Platform</span></span></th>
    <th><span data-ttu-id="5eb81-884">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="5eb81-884">Extension points</span></span></th>
    <th><span data-ttu-id="5eb81-885">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-885">API requirement sets</span></span></th>
    <th><span data-ttu-id="5eb81-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="5eb81-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-887">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-887">Office 2019 on Windows</span></span><br><span data-ttu-id="5eb81-888">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-888">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-889">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-889">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-891">- Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-891">- Selection</span></span><br><span data-ttu-id="5eb81-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-892">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-893">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-893">Office 2016 on Windows</span></span><br><span data-ttu-id="5eb81-894">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-894">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-895">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-895">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-897">- Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-897">- Selection</span></span><br><span data-ttu-id="5eb81-898">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-898">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5eb81-899">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="5eb81-899">Office 2013 on Windows</span></span><br><span data-ttu-id="5eb81-900">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="5eb81-900">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5eb81-901">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="5eb81-901">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5eb81-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5eb81-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5eb81-903">- Seleção</span><span class="sxs-lookup"><span data-stu-id="5eb81-903">- Selection</span></span><br><span data-ttu-id="5eb81-904">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5eb81-904">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5eb81-905">Confira também</span><span class="sxs-lookup"><span data-stu-id="5eb81-905">See also</span></span>

- [<span data-ttu-id="5eb81-906">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="5eb81-906">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5eb81-907">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="5eb81-907">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="5eb81-908">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="5eb81-908">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5eb81-909">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="5eb81-909">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5eb81-910">Referência da API JavaScript do Office</span><span class="sxs-lookup"><span data-stu-id="5eb81-910">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="5eb81-911">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="5eb81-911">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="5eb81-912">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="5eb81-912">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="5eb81-913">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="5eb81-913">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="5eb81-914">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="5eb81-914">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="5eb81-915">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="5eb81-915">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="5eb81-916">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="5eb81-916">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
