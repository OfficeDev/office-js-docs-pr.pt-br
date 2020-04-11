---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 04/07/2020
localization_priority: Priority
ms.openlocfilehash: 823fd53e71c71f4a845f9a7b5c6177ad3f14745f
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185614"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ac881-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ac881-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ac881-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="ac881-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="ac881-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="ac881-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ac881-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="ac881-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ac881-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="ac881-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ac881-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ac881-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ac881-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ac881-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ac881-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ac881-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac881-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-114">- TaskPane</span></span><br><span data-ttu-id="ac881-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-115">
        - Content</span></span><br><span data-ttu-id="ac881-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-116">
        - Custom Functions</span></span><br><span data-ttu-id="ac881-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac881-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac881-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac881-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac881-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac881-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac881-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ac881-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ac881-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-130">
        - BindingEvents</span></span><br><span data-ttu-id="ac881-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-131">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-132">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-133">
        - File</span></span><br><span data-ttu-id="ac881-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-134">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-136">
        - Selection</span></span><br><span data-ttu-id="ac881-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-137">
        - Settings</span></span><br><span data-ttu-id="ac881-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-138">
        - TableBindings</span></span><br><span data-ttu-id="ac881-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-139">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-140">
        - TextBindings</span></span><br><span data-ttu-id="ac881-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-142">Office on Windows</span></span><br><span data-ttu-id="ac881-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-144">- TaskPane</span></span><br><span data-ttu-id="ac881-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-145">
        - Content</span></span><br><span data-ttu-id="ac881-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-146">
        - Custom Functions</span></span><br><span data-ttu-id="ac881-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="ac881-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ac881-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac881-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac881-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac881-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac881-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ac881-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-161">
        - BindingEvents</span></span><br><span data-ttu-id="ac881-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-162">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-163">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-164">
        - File</span></span><br><span data-ttu-id="ac881-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-165">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-167">
        - Selection</span></span><br><span data-ttu-id="ac881-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-168">
        - Settings</span></span><br><span data-ttu-id="ac881-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-169">
        - TableBindings</span></span><br><span data-ttu-id="ac881-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-170">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-171">
        - TextBindings</span></span><br><span data-ttu-id="ac881-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-173">Office 2019 on Windows</span></span><br><span data-ttu-id="ac881-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac881-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-175">- TaskPane</span></span><br><span data-ttu-id="ac881-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-176">
        - Content</span></span><br><span data-ttu-id="ac881-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac881-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-188">- BindingEvents</span></span><br><span data-ttu-id="ac881-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-189">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-190">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-191">
        - File</span></span><br><span data-ttu-id="ac881-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-192">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-194">
        - Selection</span></span><br><span data-ttu-id="ac881-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-195">
        - Settings</span></span><br><span data-ttu-id="ac881-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-196">
        - TableBindings</span></span><br><span data-ttu-id="ac881-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-197">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-198">
        - TextBindings</span></span><br><span data-ttu-id="ac881-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-200">Office 2016 on Windows</span></span><br><span data-ttu-id="ac881-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac881-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-202">- TaskPane</span></span><br><span data-ttu-id="ac881-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-203">
        - Content</span></span></td>
    <td><span data-ttu-id="ac881-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac881-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-207">- BindingEvents</span></span><br><span data-ttu-id="ac881-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-208">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-209">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-210">
        - File</span></span><br><span data-ttu-id="ac881-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-211">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-213">
        - Selection</span></span><br><span data-ttu-id="ac881-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-214">
        - Settings</span></span><br><span data-ttu-id="ac881-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-215">
        - TableBindings</span></span><br><span data-ttu-id="ac881-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-216">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-217">
        - TextBindings</span></span><br><span data-ttu-id="ac881-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-219">Office 2013 on Windows</span></span><br><span data-ttu-id="ac881-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac881-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-221">
        - TaskPane</span></span><br><span data-ttu-id="ac881-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ac881-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac881-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac881-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-225">
        - BindingEvents</span></span><br><span data-ttu-id="ac881-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-226">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-227">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-228">
        - File</span></span><br><span data-ttu-id="ac881-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-229">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-231">
        - Selection</span></span><br><span data-ttu-id="ac881-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-232">
        - Settings</span></span><br><span data-ttu-id="ac881-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-233">
        - TableBindings</span></span><br><span data-ttu-id="ac881-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-234">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-235">
        - TextBindings</span></span><br><span data-ttu-id="ac881-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac881-237">Office on iPad</span></span><br><span data-ttu-id="ac881-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac881-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-239">- TaskPane</span></span><br><span data-ttu-id="ac881-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-240">
        - Content</span></span></td>
    <td><span data-ttu-id="ac881-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac881-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac881-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac881-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac881-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-253">- BindingEvents</span></span><br><span data-ttu-id="ac881-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-254">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-255">
        - File</span></span><br><span data-ttu-id="ac881-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-256">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-258">
        - Selection</span></span><br><span data-ttu-id="ac881-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-259">
        - Settings</span></span><br><span data-ttu-id="ac881-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-260">
        - TableBindings</span></span><br><span data-ttu-id="ac881-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-261">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-262">
        - TextBindings</span></span><br><span data-ttu-id="ac881-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-264">Office on Mac</span></span><br><span data-ttu-id="ac881-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac881-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-266">- TaskPane</span></span><br><span data-ttu-id="ac881-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-267">
        - Content</span></span><br><span data-ttu-id="ac881-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-268">
        - Custom Functions</span></span><br><span data-ttu-id="ac881-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac881-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ac881-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ac881-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ac881-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ac881-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ac881-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-283">- BindingEvents</span></span><br><span data-ttu-id="ac881-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-284">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-285">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-286">
        - File</span></span><br><span data-ttu-id="ac881-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-287">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-289">
        - PdfFile</span></span><br><span data-ttu-id="ac881-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-290">
        - Selection</span></span><br><span data-ttu-id="ac881-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-291">
        - Settings</span></span><br><span data-ttu-id="ac881-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-292">
        - TableBindings</span></span><br><span data-ttu-id="ac881-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-293">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-294">
        - TextBindings</span></span><br><span data-ttu-id="ac881-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-296">Office 2019 on Mac</span></span><br><span data-ttu-id="ac881-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac881-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-298">- TaskPane</span></span><br><span data-ttu-id="ac881-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-299">
        - Content</span></span><br><span data-ttu-id="ac881-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ac881-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ac881-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ac881-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ac881-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ac881-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ac881-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ac881-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ac881-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-311">- BindingEvents</span></span><br><span data-ttu-id="ac881-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-312">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-313">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-314">
        - File</span></span><br><span data-ttu-id="ac881-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-315">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-317">
        - PdfFile</span></span><br><span data-ttu-id="ac881-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-318">
        - Selection</span></span><br><span data-ttu-id="ac881-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-319">
        - Settings</span></span><br><span data-ttu-id="ac881-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-320">
        - TableBindings</span></span><br><span data-ttu-id="ac881-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-321">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-322">
        - TextBindings</span></span><br><span data-ttu-id="ac881-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-324">Office 2016 on Mac</span></span><br><span data-ttu-id="ac881-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ac881-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-326">- TaskPane</span></span><br><span data-ttu-id="ac881-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-327">
        - Content</span></span></td>
    <td><span data-ttu-id="ac881-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ac881-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac881-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ac881-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-331">- BindingEvents</span></span><br><span data-ttu-id="ac881-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-332">
        - CompressedFile</span></span><br><span data-ttu-id="ac881-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-333">
        - DocumentEvents</span></span><br><span data-ttu-id="ac881-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-334">
        - File</span></span><br><span data-ttu-id="ac881-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-335">
        - MatrixBindings</span></span><br><span data-ttu-id="ac881-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="ac881-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-337">
        - PdfFile</span></span><br><span data-ttu-id="ac881-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-338">
        - Selection</span></span><br><span data-ttu-id="ac881-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-339">
        - Settings</span></span><br><span data-ttu-id="ac881-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-340">
        - TableBindings</span></span><br><span data-ttu-id="ac881-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-341">
        - TableCoercion</span></span><br><span data-ttu-id="ac881-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-342">
        - TextBindings</span></span><br><span data-ttu-id="ac881-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac881-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac881-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ac881-345">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="ac881-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ac881-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ac881-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ac881-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ac881-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-350">Office on the web</span></span></td>
    <td><span data-ttu-id="ac881-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac881-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-353">Office on Windows</span></span><br><span data-ttu-id="ac881-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ac881-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac881-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-357">Office for Mac</span></span><br><span data-ttu-id="ac881-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ac881-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="ac881-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ac881-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ac881-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="ac881-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac881-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-362">Platform</span></span></th>
    <th><span data-ttu-id="ac881-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-363">Extension points</span></span></th>
    <th><span data-ttu-id="ac881-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac881-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-366">Office on the web</span></span><br><span data-ttu-id="ac881-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="ac881-367">(modern)</span></span></td>
    <td> <span data-ttu-id="ac881-368">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-368">- Message Read</span></span><br><span data-ttu-id="ac881-369">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-369">
      - Message Compose</span></span><br><span data-ttu-id="ac881-370">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-370">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-371">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-371">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac881-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac881-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac881-381">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-382">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-382">Office on the web</span></span><br><span data-ttu-id="ac881-383">(clássico)</span><span class="sxs-lookup"><span data-stu-id="ac881-383">(classic)</span></span></td>
    <td> <span data-ttu-id="ac881-384">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-384">- Message Read</span></span><br><span data-ttu-id="ac881-385">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-385">
      - Message Compose</span></span><br><span data-ttu-id="ac881-386">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-386">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-387">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-387">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac881-395">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-396">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-396">Office on Windows</span></span><br><span data-ttu-id="ac881-397">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-398">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-398">- Message Read</span></span><br><span data-ttu-id="ac881-399">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-399">
      - Message Compose</span></span><br><span data-ttu-id="ac881-400">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-400">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-401">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-401">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac881-403">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="ac881-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac881-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac881-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac881-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac881-412">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-413">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-413">Office 2019 on Windows</span></span><br><span data-ttu-id="ac881-414">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-415">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-415">- Message Read</span></span><br><span data-ttu-id="ac881-416">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-416">
      - Message Compose</span></span><br><span data-ttu-id="ac881-417">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-417">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-418">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-418">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac881-420">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="ac881-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac881-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac881-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ac881-428">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-429">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-429">Office 2016 on Windows</span></span><br><span data-ttu-id="ac881-430">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-431">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-431">- Message Read</span></span><br><span data-ttu-id="ac881-432">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-432">
      - Message Compose</span></span><br><span data-ttu-id="ac881-433">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-433">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-434">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-434">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ac881-436">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="ac881-436">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ac881-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ac881-441">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-442">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-442">Office 2013 on Windows</span></span><br><span data-ttu-id="ac881-443">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-444">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-444">- Message Read</span></span><br><span data-ttu-id="ac881-445">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-445">
      - Message Compose</span></span><br><span data-ttu-id="ac881-446">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-446">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-447">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-447">
      - Appointment Organizer (Compose)</span></span><br>
    <td> <span data-ttu-id="ac881-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ac881-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ac881-452">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-453">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="ac881-453">Office on iOS</span></span><br><span data-ttu-id="ac881-454">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-455">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-455">- Message Read</span></span><br><span data-ttu-id="ac881-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ac881-462">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-463">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-463">Office on Mac</span></span><br><span data-ttu-id="ac881-464">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-465">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-465">- Message Read</span></span><br><span data-ttu-id="ac881-466">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-466">
      - Message Compose</span></span><br><span data-ttu-id="ac881-467">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-467">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-468">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-468">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ac881-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ac881-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ac881-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ac881-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ac881-478">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-479">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-479">Office 2019 on Mac</span></span><br><span data-ttu-id="ac881-480">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-481">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-481">- Message Read</span></span><br><span data-ttu-id="ac881-482">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-482">
      - Message Compose</span></span><br><span data-ttu-id="ac881-483">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-483">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-484">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-484">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac881-492">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-493">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-493">Office 2016 on Mac</span></span><br><span data-ttu-id="ac881-494">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-495">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-495">- Message Read</span></span><br><span data-ttu-id="ac881-496">
      - Composição da mensagem</span><span class="sxs-lookup"><span data-stu-id="ac881-496">
      - Message Compose</span></span><br><span data-ttu-id="ac881-497">
      - Participante do compromisso (Leitura)</span><span class="sxs-lookup"><span data-stu-id="ac881-497">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ac881-498">
      - Organizador de compromissos (Redigir)</span><span class="sxs-lookup"><span data-stu-id="ac881-498">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ac881-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ac881-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ac881-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ac881-506">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-507">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="ac881-507">Office on Android</span></span><br><span data-ttu-id="ac881-508">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-509">- Mensagem lida</span><span class="sxs-lookup"><span data-stu-id="ac881-509">- Message Read</span></span><br><span data-ttu-id="ac881-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ac881-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ac881-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ac881-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ac881-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ac881-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ac881-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ac881-516">Não disponível</span><span class="sxs-lookup"><span data-stu-id="ac881-516">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac881-517">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac881-517">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ac881-518">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="ac881-518">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ac881-519">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="ac881-519">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ac881-520">Word</span><span class="sxs-lookup"><span data-stu-id="ac881-520">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac881-521">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-521">Platform</span></span></th>
    <th><span data-ttu-id="ac881-522">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-522">Extension points</span></span></th>
    <th><span data-ttu-id="ac881-523">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-523">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac881-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-525">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-525">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac881-526">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-526">- TaskPane</span></span><br><span data-ttu-id="ac881-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac881-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-534">- BindingEvents</span></span><br><span data-ttu-id="ac881-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-536">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-537">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-537">
         - File</span></span><br><span data-ttu-id="ac881-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-539">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-542">
         - PdfFile</span></span><br><span data-ttu-id="ac881-543">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-543">
         - Selection</span></span><br><span data-ttu-id="ac881-544">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-544">
         - Settings</span></span><br><span data-ttu-id="ac881-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-545">
         - TableBindings</span></span><br><span data-ttu-id="ac881-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-546">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-547">
         - TextBindings</span></span><br><span data-ttu-id="ac881-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-548">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-549">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-550">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-550">Office on Windows</span></span><br><span data-ttu-id="ac881-551">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-551">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-552">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-552">- TaskPane</span></span><br><span data-ttu-id="ac881-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac881-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-560">- BindingEvents</span></span><br><span data-ttu-id="ac881-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-561">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-563">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-564">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-564">
         - File</span></span><br><span data-ttu-id="ac881-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-566">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-569">
         - PdfFile</span></span><br><span data-ttu-id="ac881-570">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-570">
         - Selection</span></span><br><span data-ttu-id="ac881-571">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-571">
         - Settings</span></span><br><span data-ttu-id="ac881-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-572">
         - TableBindings</span></span><br><span data-ttu-id="ac881-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-573">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-574">
         - TextBindings</span></span><br><span data-ttu-id="ac881-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-575">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-577">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-577">Office 2019 on Windows</span></span><br><span data-ttu-id="ac881-578">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-579">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-579">- TaskPane</span></span><br><span data-ttu-id="ac881-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-586">- BindingEvents</span></span><br><span data-ttu-id="ac881-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-587">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-589">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-590">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-590">
         - File</span></span><br><span data-ttu-id="ac881-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-592">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-595">
         - PdfFile</span></span><br><span data-ttu-id="ac881-596">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-596">
         - Selection</span></span><br><span data-ttu-id="ac881-597">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-597">
         - Settings</span></span><br><span data-ttu-id="ac881-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-598">
         - TableBindings</span></span><br><span data-ttu-id="ac881-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-599">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-600">
         - TextBindings</span></span><br><span data-ttu-id="ac881-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-601">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-603">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-603">Office 2016 on Windows</span></span><br><span data-ttu-id="ac881-604">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-605">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac881-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-609">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-609">- BindingEvents</span></span><br><span data-ttu-id="ac881-610">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-610">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-611">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-611">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-612">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-613">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-613">
         - File</span></span><br><span data-ttu-id="ac881-614">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-614">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-615">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-615">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-616">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-616">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-617">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-617">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-618">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-618">
         - PdfFile</span></span><br><span data-ttu-id="ac881-619">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-619">
         - Selection</span></span><br><span data-ttu-id="ac881-620">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-620">
         - Settings</span></span><br><span data-ttu-id="ac881-621">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-621">
         - TableBindings</span></span><br><span data-ttu-id="ac881-622">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-622">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-623">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-623">
         - TextBindings</span></span><br><span data-ttu-id="ac881-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-624">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-625">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-625">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-626">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-626">Office 2013 on Windows</span></span><br><span data-ttu-id="ac881-627">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-627">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-628">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-628">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac881-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac881-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-631">- BindingEvents</span></span><br><span data-ttu-id="ac881-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-632">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-634">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-635">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-635">
         - File</span></span><br><span data-ttu-id="ac881-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-637">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-640">
         - PdfFile</span></span><br><span data-ttu-id="ac881-641">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-641">
         - Selection</span></span><br><span data-ttu-id="ac881-642">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-642">
         - Settings</span></span><br><span data-ttu-id="ac881-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-643">
         - TableBindings</span></span><br><span data-ttu-id="ac881-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-644">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-645">
         - TextBindings</span></span><br><span data-ttu-id="ac881-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-646">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-647">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-648">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac881-648">Office on iPad</span></span><br><span data-ttu-id="ac881-649">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-650">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-650">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac881-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-656">- BindingEvents</span></span><br><span data-ttu-id="ac881-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-657">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-659">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-660">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-660">
         - File</span></span><br><span data-ttu-id="ac881-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-662">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-665">
         - PdfFile</span></span><br><span data-ttu-id="ac881-666">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-666">
         - Selection</span></span><br><span data-ttu-id="ac881-667">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-667">
         - Settings</span></span><br><span data-ttu-id="ac881-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-668">
         - TableBindings</span></span><br><span data-ttu-id="ac881-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-669">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-670">
         - TextBindings</span></span><br><span data-ttu-id="ac881-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-671">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-673">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-673">Office on Mac</span></span><br><span data-ttu-id="ac881-674">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-674">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-675">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-675">- TaskPane</span></span><br><span data-ttu-id="ac881-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac881-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-683">- BindingEvents</span></span><br><span data-ttu-id="ac881-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-684">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-686">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-687">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-687">
         - File</span></span><br><span data-ttu-id="ac881-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-689">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-692">
         - PdfFile</span></span><br><span data-ttu-id="ac881-693">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-693">
         - Selection</span></span><br><span data-ttu-id="ac881-694">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-694">
         - Settings</span></span><br><span data-ttu-id="ac881-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-695">
         - TableBindings</span></span><br><span data-ttu-id="ac881-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-696">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-697">
         - TextBindings</span></span><br><span data-ttu-id="ac881-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-698">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-700">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-700">Office 2019 on Mac</span></span><br><span data-ttu-id="ac881-701">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-702">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-702">- TaskPane</span></span><br><span data-ttu-id="ac881-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ac881-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ac881-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ac881-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ac881-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-709">- BindingEvents</span></span><br><span data-ttu-id="ac881-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-710">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-712">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-713">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-713">
         - File</span></span><br><span data-ttu-id="ac881-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-715">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-718">
         - PdfFile</span></span><br><span data-ttu-id="ac881-719">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-719">
         - Selection</span></span><br><span data-ttu-id="ac881-720">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-720">
         - Settings</span></span><br><span data-ttu-id="ac881-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-721">
         - TableBindings</span></span><br><span data-ttu-id="ac881-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-722">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-723">
         - TextBindings</span></span><br><span data-ttu-id="ac881-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-724">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-725">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-726">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-726">Office 2016 on Mac</span></span><br><span data-ttu-id="ac881-727">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-727">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-728">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-728">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ac881-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ac881-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ac881-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-732">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-732">- BindingEvents</span></span><br><span data-ttu-id="ac881-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-733">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-734">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ac881-734">
         - CustomXmlParts</span></span><br><span data-ttu-id="ac881-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-735">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-736">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-736">
         - File</span></span><br><span data-ttu-id="ac881-737">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-737">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-738">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-738">
         - MatrixBindings</span></span><br><span data-ttu-id="ac881-739">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-739">
         - MatrixCoercion</span></span><br><span data-ttu-id="ac881-740">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-740">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ac881-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-741">
         - PdfFile</span></span><br><span data-ttu-id="ac881-742">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-742">
         - Selection</span></span><br><span data-ttu-id="ac881-743">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-743">
         - Settings</span></span><br><span data-ttu-id="ac881-744">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-744">
         - TableBindings</span></span><br><span data-ttu-id="ac881-745">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-745">
         - TableCoercion</span></span><br><span data-ttu-id="ac881-746">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ac881-746">
         - TextBindings</span></span><br><span data-ttu-id="ac881-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-747">
         - TextCoercion</span></span><br><span data-ttu-id="ac881-748">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ac881-748">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ac881-749">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac881-749">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ac881-750">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ac881-750">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac881-751">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-751">Platform</span></span></th>
    <th><span data-ttu-id="ac881-752">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-752">Extension points</span></span></th>
    <th><span data-ttu-id="ac881-753">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-753">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac881-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-755">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-755">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac881-756">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-756">- Content</span></span><br><span data-ttu-id="ac881-757">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-757">
         - TaskPane</span></span><br><span data-ttu-id="ac881-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac881-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac881-763">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-763">- ActiveView</span></span><br><span data-ttu-id="ac881-764">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-764">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-765">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-765">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-766">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-766">
         - File</span></span><br><span data-ttu-id="ac881-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-767">
         - PdfFile</span></span><br><span data-ttu-id="ac881-768">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-768">
         - Selection</span></span><br><span data-ttu-id="ac881-769">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-769">
         - Settings</span></span><br><span data-ttu-id="ac881-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-771">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-771">Office on Windows</span></span><br><span data-ttu-id="ac881-772">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-772">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-773">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-773">- Content</span></span><br><span data-ttu-id="ac881-774">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-774">
         - TaskPane</span></span><br><span data-ttu-id="ac881-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac881-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac881-780">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-780">- ActiveView</span></span><br><span data-ttu-id="ac881-781">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-781">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-782">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-782">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-783">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-783">
         - File</span></span><br><span data-ttu-id="ac881-784">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-784">
         - PdfFile</span></span><br><span data-ttu-id="ac881-785">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-785">
         - Selection</span></span><br><span data-ttu-id="ac881-786">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-786">
         - Settings</span></span><br><span data-ttu-id="ac881-787">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-787">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-788">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-788">Office 2019 on Windows</span></span><br><span data-ttu-id="ac881-789">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-789">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-790">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-790">- Content</span></span><br><span data-ttu-id="ac881-791">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-791">
         - TaskPane</span></span><br><span data-ttu-id="ac881-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-795">- ActiveView</span></span><br><span data-ttu-id="ac881-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-796">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-797">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-798">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-798">
         - File</span></span><br><span data-ttu-id="ac881-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-799">
         - PdfFile</span></span><br><span data-ttu-id="ac881-800">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-800">
         - Selection</span></span><br><span data-ttu-id="ac881-801">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-801">
         - Settings</span></span><br><span data-ttu-id="ac881-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-803">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-803">Office 2016 on Windows</span></span><br><span data-ttu-id="ac881-804">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-804">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-805">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-805">- Content</span></span><br><span data-ttu-id="ac881-806">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac881-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac881-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-809">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-809">- ActiveView</span></span><br><span data-ttu-id="ac881-810">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-810">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-811">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-811">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-812">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-812">
         - File</span></span><br><span data-ttu-id="ac881-813">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-813">
         - PdfFile</span></span><br><span data-ttu-id="ac881-814">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-814">
         - Selection</span></span><br><span data-ttu-id="ac881-815">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-815">
         - Settings</span></span><br><span data-ttu-id="ac881-816">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-816">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-817">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-817">Office 2013 on Windows</span></span><br><span data-ttu-id="ac881-818">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-818">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-819">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-819">- Content</span></span><br><span data-ttu-id="ac881-820">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-820">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ac881-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac881-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac881-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-823">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-823">- ActiveView</span></span><br><span data-ttu-id="ac881-824">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-824">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-825">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-825">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-826">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-826">
         - File</span></span><br><span data-ttu-id="ac881-827">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-827">
         - PdfFile</span></span><br><span data-ttu-id="ac881-828">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-828">
         - Selection</span></span><br><span data-ttu-id="ac881-829">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-829">
         - Settings</span></span><br><span data-ttu-id="ac881-830">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-830">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-831">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="ac881-831">Office on iPad</span></span><br><span data-ttu-id="ac881-832">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-832">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-833">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-833">- Content</span></span><br><span data-ttu-id="ac881-834">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-834">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac881-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-838">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-838">- ActiveView</span></span><br><span data-ttu-id="ac881-839">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-839">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-840">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-840">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-841">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-841">
         - File</span></span><br><span data-ttu-id="ac881-842">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-842">
         - PdfFile</span></span><br><span data-ttu-id="ac881-843">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-843">
         - Selection</span></span><br><span data-ttu-id="ac881-844">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-844">
         - Settings</span></span><br><span data-ttu-id="ac881-845">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-845">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-846">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-846">Office on Mac</span></span><br><span data-ttu-id="ac881-847">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="ac881-847">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ac881-848">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-848">- Content</span></span><br><span data-ttu-id="ac881-849">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-849">
         - TaskPane</span></span><br><span data-ttu-id="ac881-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ac881-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ac881-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ac881-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ac881-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-855">- ActiveView</span></span><br><span data-ttu-id="ac881-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-856">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-857">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-858">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-858">
         - File</span></span><br><span data-ttu-id="ac881-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-859">
         - PdfFile</span></span><br><span data-ttu-id="ac881-860">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-860">
         - Selection</span></span><br><span data-ttu-id="ac881-861">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-861">
         - Settings</span></span><br><span data-ttu-id="ac881-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-862">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-863">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-863">Office 2019 on Mac</span></span><br><span data-ttu-id="ac881-864">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-864">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-865">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-865">- Content</span></span><br><span data-ttu-id="ac881-866">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-866">
         - TaskPane</span></span><br><span data-ttu-id="ac881-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-870">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-870">- ActiveView</span></span><br><span data-ttu-id="ac881-871">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-871">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-872">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-872">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-873">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-873">
         - File</span></span><br><span data-ttu-id="ac881-874">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-874">
         - PdfFile</span></span><br><span data-ttu-id="ac881-875">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-875">
         - Selection</span></span><br><span data-ttu-id="ac881-876">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-876">
         - Settings</span></span><br><span data-ttu-id="ac881-877">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-877">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-878">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-878">Office 2016 on Mac</span></span><br><span data-ttu-id="ac881-879">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-879">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-880">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-880">- Content</span></span><br><span data-ttu-id="ac881-881">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-881">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ac881-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ac881-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-884">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ac881-884">- ActiveView</span></span><br><span data-ttu-id="ac881-885">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ac881-885">
         - CompressedFile</span></span><br><span data-ttu-id="ac881-886">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-886">
         - DocumentEvents</span></span><br><span data-ttu-id="ac881-887">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="ac881-887">
         - File</span></span><br><span data-ttu-id="ac881-888">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ac881-888">
         - PdfFile</span></span><br><span data-ttu-id="ac881-889">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-889">
         - Selection</span></span><br><span data-ttu-id="ac881-890">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-890">
         - Settings</span></span><br><span data-ttu-id="ac881-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ac881-892">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="ac881-892">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ac881-893">OneNote</span><span class="sxs-lookup"><span data-stu-id="ac881-893">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac881-894">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-894">Platform</span></span></th>
    <th><span data-ttu-id="ac881-895">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-895">Extension points</span></span></th>
    <th><span data-ttu-id="ac881-896">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-896">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac881-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-898">Office na Web</span><span class="sxs-lookup"><span data-stu-id="ac881-898">Office on the web</span></span></td>
    <td> <span data-ttu-id="ac881-899">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="ac881-899">- Content</span></span><br><span data-ttu-id="ac881-900">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-900">
         - TaskPane</span></span><br><span data-ttu-id="ac881-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="ac881-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ac881-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ac881-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ac881-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-905">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ac881-905">- DocumentEvents</span></span><br><span data-ttu-id="ac881-906">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-906">
         - HtmlCoercion</span></span><br><span data-ttu-id="ac881-907">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="ac881-907">
         - Settings</span></span><br><span data-ttu-id="ac881-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ac881-909">Project</span><span class="sxs-lookup"><span data-stu-id="ac881-909">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ac881-910">Plataforma</span><span class="sxs-lookup"><span data-stu-id="ac881-910">Platform</span></span></th>
    <th><span data-ttu-id="ac881-911">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="ac881-911">Extension points</span></span></th>
    <th><span data-ttu-id="ac881-912">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="ac881-912">API requirement sets</span></span></th>
    <th><span data-ttu-id="ac881-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="ac881-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-914">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-914">Office 2019 on Windows</span></span><br><span data-ttu-id="ac881-915">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-915">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-916">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-916">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-918">- Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-918">- Selection</span></span><br><span data-ttu-id="ac881-919">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-919">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-920">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-920">Office 2016 on Windows</span></span><br><span data-ttu-id="ac881-921">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-921">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-922">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-922">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-924">- Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-924">- Selection</span></span><br><span data-ttu-id="ac881-925">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-925">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ac881-926">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="ac881-926">Office 2013 on Windows</span></span><br><span data-ttu-id="ac881-927">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="ac881-927">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ac881-928">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="ac881-928">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ac881-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ac881-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ac881-930">- Seleção</span><span class="sxs-lookup"><span data-stu-id="ac881-930">- Selection</span></span><br><span data-ttu-id="ac881-931">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ac881-931">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ac881-932">Confira também</span><span class="sxs-lookup"><span data-stu-id="ac881-932">See also</span></span>

- [<span data-ttu-id="ac881-933">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="ac881-933">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ac881-934">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="ac881-934">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ac881-935">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="ac881-935">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ac881-936">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="ac881-936">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ac881-937">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="ac881-937">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ac881-938">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="ac881-938">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ac881-939">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="ac881-939">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ac881-940">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="ac881-940">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ac881-941">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac881-941">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ac881-942">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="ac881-942">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ac881-943">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="ac881-943">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ac881-944">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="ac881-944">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)