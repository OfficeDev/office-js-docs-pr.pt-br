---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554017"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="75487-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="75487-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="75487-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="75487-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="75487-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="75487-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="75487-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="75487-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="75487-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="75487-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="75487-108">Excel</span><span class="sxs-lookup"><span data-stu-id="75487-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="75487-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="75487-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="75487-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="75487-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="75487-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-114">- TaskPane</span></span><br><span data-ttu-id="75487-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-115">
        - Content</span></span><br><span data-ttu-id="75487-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-116">
        - Custom Functions</span></span><br><span data-ttu-id="75487-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="75487-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="75487-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75487-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75487-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="75487-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="75487-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="75487-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="75487-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-130">
        - BindingEvents</span></span><br><span data-ttu-id="75487-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-131">
        - CompressedFile</span></span><br><span data-ttu-id="75487-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-132">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-133">
        - File</span></span><br><span data-ttu-id="75487-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-134">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-136">
        - Selection</span></span><br><span data-ttu-id="75487-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-137">
        - Settings</span></span><br><span data-ttu-id="75487-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-138">
        - TableBindings</span></span><br><span data-ttu-id="75487-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-139">
        - TableCoercion</span></span><br><span data-ttu-id="75487-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-140">
        - TextBindings</span></span><br><span data-ttu-id="75487-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-142">Office on Windows</span></span><br><span data-ttu-id="75487-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-144">- TaskPane</span></span><br><span data-ttu-id="75487-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-145">
        - Content</span></span><br><span data-ttu-id="75487-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-146">
        - Custom Functions</span></span><br><span data-ttu-id="75487-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="75487-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="75487-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75487-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75487-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="75487-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="75487-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="75487-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-161">
        - BindingEvents</span></span><br><span data-ttu-id="75487-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-162">
        - CompressedFile</span></span><br><span data-ttu-id="75487-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-163">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-164">
        - File</span></span><br><span data-ttu-id="75487-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-165">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-167">
        - Selection</span></span><br><span data-ttu-id="75487-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-168">
        - Settings</span></span><br><span data-ttu-id="75487-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-169">
        - TableBindings</span></span><br><span data-ttu-id="75487-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-170">
        - TableCoercion</span></span><br><span data-ttu-id="75487-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-171">
        - TextBindings</span></span><br><span data-ttu-id="75487-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-173">Office 2019 on Windows</span></span><br><span data-ttu-id="75487-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75487-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-175">- TaskPane</span></span><br><span data-ttu-id="75487-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-176">
        - Content</span></span><br><span data-ttu-id="75487-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75487-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-188">- BindingEvents</span></span><br><span data-ttu-id="75487-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-189">
        - CompressedFile</span></span><br><span data-ttu-id="75487-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-190">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-191">
        - File</span></span><br><span data-ttu-id="75487-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-192">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-194">
        - Selection</span></span><br><span data-ttu-id="75487-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-195">
        - Settings</span></span><br><span data-ttu-id="75487-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-196">
        - TableBindings</span></span><br><span data-ttu-id="75487-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-197">
        - TableCoercion</span></span><br><span data-ttu-id="75487-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-198">
        - TextBindings</span></span><br><span data-ttu-id="75487-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-200">Office 2016 on Windows</span></span><br><span data-ttu-id="75487-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75487-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-202">- TaskPane</span></span><br><span data-ttu-id="75487-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-203">
        - Content</span></span></td>
    <td><span data-ttu-id="75487-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="75487-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-207">- BindingEvents</span></span><br><span data-ttu-id="75487-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-208">
        - CompressedFile</span></span><br><span data-ttu-id="75487-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-209">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-210">
        - File</span></span><br><span data-ttu-id="75487-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-211">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-213">
        - Selection</span></span><br><span data-ttu-id="75487-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-214">
        - Settings</span></span><br><span data-ttu-id="75487-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-215">
        - TableBindings</span></span><br><span data-ttu-id="75487-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-216">
        - TableCoercion</span></span><br><span data-ttu-id="75487-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-217">
        - TextBindings</span></span><br><span data-ttu-id="75487-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-219">Office 2013 on Windows</span></span><br><span data-ttu-id="75487-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75487-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-221">
        - TaskPane</span></span><br><span data-ttu-id="75487-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="75487-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75487-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="75487-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-225">
        - BindingEvents</span></span><br><span data-ttu-id="75487-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-226">
        - CompressedFile</span></span><br><span data-ttu-id="75487-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-227">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-228">
        - File</span></span><br><span data-ttu-id="75487-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-229">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-231">
        - Selection</span></span><br><span data-ttu-id="75487-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-232">
        - Settings</span></span><br><span data-ttu-id="75487-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-233">
        - TableBindings</span></span><br><span data-ttu-id="75487-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-234">
        - TableCoercion</span></span><br><span data-ttu-id="75487-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-235">
        - TextBindings</span></span><br><span data-ttu-id="75487-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="75487-237">Office on iPad</span></span><br><span data-ttu-id="75487-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75487-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-239">- TaskPane</span></span><br><span data-ttu-id="75487-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-240">
        - Content</span></span></td>
    <td><span data-ttu-id="75487-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75487-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75487-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="75487-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="75487-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-253">- BindingEvents</span></span><br><span data-ttu-id="75487-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-254">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-255">
        - File</span></span><br><span data-ttu-id="75487-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-256">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-258">
        - Selection</span></span><br><span data-ttu-id="75487-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-259">
        - Settings</span></span><br><span data-ttu-id="75487-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-260">
        - TableBindings</span></span><br><span data-ttu-id="75487-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-261">
        - TableCoercion</span></span><br><span data-ttu-id="75487-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-262">
        - TextBindings</span></span><br><span data-ttu-id="75487-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-264">Office on Mac</span></span><br><span data-ttu-id="75487-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75487-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-266">- TaskPane</span></span><br><span data-ttu-id="75487-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-267">
        - Content</span></span><br><span data-ttu-id="75487-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-268">
        - Custom Functions</span></span><br><span data-ttu-id="75487-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75487-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="75487-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="75487-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="75487-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="75487-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="75487-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-283">- BindingEvents</span></span><br><span data-ttu-id="75487-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-284">
        - CompressedFile</span></span><br><span data-ttu-id="75487-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-285">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-286">
        - File</span></span><br><span data-ttu-id="75487-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-287">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-289">
        - PdfFile</span></span><br><span data-ttu-id="75487-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-290">
        - Selection</span></span><br><span data-ttu-id="75487-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-291">
        - Settings</span></span><br><span data-ttu-id="75487-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-292">
        - TableBindings</span></span><br><span data-ttu-id="75487-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-293">
        - TableCoercion</span></span><br><span data-ttu-id="75487-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-294">
        - TextBindings</span></span><br><span data-ttu-id="75487-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-296">Office 2019 on Mac</span></span><br><span data-ttu-id="75487-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75487-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-298">- TaskPane</span></span><br><span data-ttu-id="75487-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-299">
        - Content</span></span><br><span data-ttu-id="75487-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="75487-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="75487-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="75487-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="75487-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="75487-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="75487-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="75487-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="75487-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-311">- BindingEvents</span></span><br><span data-ttu-id="75487-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-312">
        - CompressedFile</span></span><br><span data-ttu-id="75487-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-313">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-314">
        - File</span></span><br><span data-ttu-id="75487-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-315">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-317">
        - PdfFile</span></span><br><span data-ttu-id="75487-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-318">
        - Selection</span></span><br><span data-ttu-id="75487-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-319">
        - Settings</span></span><br><span data-ttu-id="75487-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-320">
        - TableBindings</span></span><br><span data-ttu-id="75487-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-321">
        - TableCoercion</span></span><br><span data-ttu-id="75487-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-322">
        - TextBindings</span></span><br><span data-ttu-id="75487-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-324">Office 2016 on Mac</span></span><br><span data-ttu-id="75487-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="75487-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-326">- TaskPane</span></span><br><span data-ttu-id="75487-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-327">
        - Content</span></span></td>
    <td><span data-ttu-id="75487-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="75487-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="75487-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="75487-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-331">- BindingEvents</span></span><br><span data-ttu-id="75487-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-332">
        - CompressedFile</span></span><br><span data-ttu-id="75487-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-333">
        - DocumentEvents</span></span><br><span data-ttu-id="75487-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-334">
        - File</span></span><br><span data-ttu-id="75487-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-335">
        - MatrixBindings</span></span><br><span data-ttu-id="75487-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="75487-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-337">
        - PdfFile</span></span><br><span data-ttu-id="75487-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-338">
        - Selection</span></span><br><span data-ttu-id="75487-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-339">
        - Settings</span></span><br><span data-ttu-id="75487-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-340">
        - TableBindings</span></span><br><span data-ttu-id="75487-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-341">
        - TableCoercion</span></span><br><span data-ttu-id="75487-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-342">
        - TextBindings</span></span><br><span data-ttu-id="75487-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="75487-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="75487-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="75487-345">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="75487-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="75487-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="75487-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="75487-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="75487-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-350">Office on the web</span></span></td>
    <td><span data-ttu-id="75487-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75487-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-353">Office on Windows</span></span><br><span data-ttu-id="75487-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="75487-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75487-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="75487-357">Office for Mac</span></span><br><span data-ttu-id="75487-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="75487-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="75487-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="75487-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="75487-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="75487-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75487-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-362">Platform</span></span></th>
    <th><span data-ttu-id="75487-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-363">Extension points</span></span></th>
    <th><span data-ttu-id="75487-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="75487-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-366">Office on the web</span></span><br><span data-ttu-id="75487-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="75487-367">(modern)</span></span></td>
    <td> <span data-ttu-id="75487-368">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-368">- Mail Read</span></span><br><span data-ttu-id="75487-369">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-369">
      - Mail Compose</span></span><br><span data-ttu-id="75487-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75487-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="75487-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="75487-379">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-380">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-380">Office on the web</span></span><br><span data-ttu-id="75487-381">(clássico)</span><span class="sxs-lookup"><span data-stu-id="75487-381">(classic)</span></span></td>
    <td> <span data-ttu-id="75487-382">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-382">- Mail Read</span></span><br><span data-ttu-id="75487-383">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-383">
      - Mail Compose</span></span><br><span data-ttu-id="75487-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75487-391">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-392">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-392">Office on Windows</span></span><br><span data-ttu-id="75487-393">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-394">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-394">- Mail Read</span></span><br><span data-ttu-id="75487-395">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-395">
      - Mail Compose</span></span><br><span data-ttu-id="75487-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75487-397">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="75487-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75487-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75487-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="75487-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="75487-406">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-407">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-407">Office 2019 on Windows</span></span><br><span data-ttu-id="75487-408">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-409">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-409">- Mail Read</span></span><br><span data-ttu-id="75487-410">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-410">
      - Mail Compose</span></span><br><span data-ttu-id="75487-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75487-412">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="75487-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75487-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75487-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="75487-420">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-421">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-421">Office 2016 on Windows</span></span><br><span data-ttu-id="75487-422">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-423">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-423">- Mail Read</span></span><br><span data-ttu-id="75487-424">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-424">
      - Mail Compose</span></span><br><span data-ttu-id="75487-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="75487-426">
      - Módulos</span><span class="sxs-lookup"><span data-stu-id="75487-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="75487-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="75487-431">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-432">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-432">Office 2013 on Windows</span></span><br><span data-ttu-id="75487-433">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-434">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-434">- Mail Read</span></span><br><span data-ttu-id="75487-435">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="75487-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="75487-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="75487-440">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-441">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="75487-441">Office on iOS</span></span><br><span data-ttu-id="75487-442">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-443">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-443">- Mail Read</span></span><br><span data-ttu-id="75487-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="75487-450">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-451">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-451">Office on Mac</span></span><br><span data-ttu-id="75487-452">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-453">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-453">- Mail Read</span></span><br><span data-ttu-id="75487-454">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-454">
      - Mail Compose</span></span><br><span data-ttu-id="75487-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="75487-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="75487-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="75487-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="75487-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="75487-464">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-465">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-465">Office 2019 on Mac</span></span><br><span data-ttu-id="75487-466">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-467">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-467">- Mail Read</span></span><br><span data-ttu-id="75487-468">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-468">
      - Mail Compose</span></span><br><span data-ttu-id="75487-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75487-476">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-477">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-477">Office 2016 on Mac</span></span><br><span data-ttu-id="75487-478">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-479">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-479">- Mail Read</span></span><br><span data-ttu-id="75487-480">
      - Composição de email</span><span class="sxs-lookup"><span data-stu-id="75487-480">
      - Mail Compose</span></span><br><span data-ttu-id="75487-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="75487-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="75487-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="75487-488">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-489">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="75487-489">Office on Android</span></span><br><span data-ttu-id="75487-490">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-491">- Leitura de email</span><span class="sxs-lookup"><span data-stu-id="75487-491">- Mail Read</span></span><br><span data-ttu-id="75487-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="75487-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="75487-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="75487-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="75487-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="75487-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="75487-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="75487-498">Não disponível</span><span class="sxs-lookup"><span data-stu-id="75487-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="75487-499">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="75487-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="75487-500">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="75487-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="75487-501">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="75487-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="75487-502">Word</span><span class="sxs-lookup"><span data-stu-id="75487-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75487-503">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-503">Platform</span></span></th>
    <th><span data-ttu-id="75487-504">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-504">Extension points</span></span></th>
    <th><span data-ttu-id="75487-505">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="75487-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-507">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="75487-508">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-508">- TaskPane</span></span><br><span data-ttu-id="75487-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="75487-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-516">- BindingEvents</span></span><br><span data-ttu-id="75487-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-518">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-519">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-519">
         - File</span></span><br><span data-ttu-id="75487-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-521">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-524">
         - PdfFile</span></span><br><span data-ttu-id="75487-525">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-525">
         - Selection</span></span><br><span data-ttu-id="75487-526">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-526">
         - Settings</span></span><br><span data-ttu-id="75487-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-527">
         - TableBindings</span></span><br><span data-ttu-id="75487-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-528">
         - TableCoercion</span></span><br><span data-ttu-id="75487-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-529">
         - TextBindings</span></span><br><span data-ttu-id="75487-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-530">
         - TextCoercion</span></span><br><span data-ttu-id="75487-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-532">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-532">Office on Windows</span></span><br><span data-ttu-id="75487-533">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-534">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-534">- TaskPane</span></span><br><span data-ttu-id="75487-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="75487-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-542">- BindingEvents</span></span><br><span data-ttu-id="75487-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-543">
         - CompressedFile</span></span><br><span data-ttu-id="75487-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-545">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-546">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-546">
         - File</span></span><br><span data-ttu-id="75487-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-548">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-551">
         - PdfFile</span></span><br><span data-ttu-id="75487-552">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-552">
         - Selection</span></span><br><span data-ttu-id="75487-553">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-553">
         - Settings</span></span><br><span data-ttu-id="75487-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-554">
         - TableBindings</span></span><br><span data-ttu-id="75487-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-555">
         - TableCoercion</span></span><br><span data-ttu-id="75487-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-556">
         - TextBindings</span></span><br><span data-ttu-id="75487-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-557">
         - TextCoercion</span></span><br><span data-ttu-id="75487-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-559">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-559">Office 2019 on Windows</span></span><br><span data-ttu-id="75487-560">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-561">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-561">- TaskPane</span></span><br><span data-ttu-id="75487-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-568">- BindingEvents</span></span><br><span data-ttu-id="75487-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-569">
         - CompressedFile</span></span><br><span data-ttu-id="75487-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-571">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-572">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-572">
         - File</span></span><br><span data-ttu-id="75487-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-574">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-577">
         - PdfFile</span></span><br><span data-ttu-id="75487-578">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-578">
         - Selection</span></span><br><span data-ttu-id="75487-579">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-579">
         - Settings</span></span><br><span data-ttu-id="75487-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-580">
         - TableBindings</span></span><br><span data-ttu-id="75487-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-581">
         - TableCoercion</span></span><br><span data-ttu-id="75487-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-582">
         - TextBindings</span></span><br><span data-ttu-id="75487-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-583">
         - TextCoercion</span></span><br><span data-ttu-id="75487-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-585">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-585">Office 2016 on Windows</span></span><br><span data-ttu-id="75487-586">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-587">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="75487-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-591">- BindingEvents</span></span><br><span data-ttu-id="75487-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-592">
         - CompressedFile</span></span><br><span data-ttu-id="75487-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-594">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-595">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-595">
         - File</span></span><br><span data-ttu-id="75487-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-597">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-600">
         - PdfFile</span></span><br><span data-ttu-id="75487-601">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-601">
         - Selection</span></span><br><span data-ttu-id="75487-602">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-602">
         - Settings</span></span><br><span data-ttu-id="75487-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-603">
         - TableBindings</span></span><br><span data-ttu-id="75487-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-604">
         - TableCoercion</span></span><br><span data-ttu-id="75487-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-605">
         - TextBindings</span></span><br><span data-ttu-id="75487-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-606">
         - TextCoercion</span></span><br><span data-ttu-id="75487-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-608">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-608">Office 2013 on Windows</span></span><br><span data-ttu-id="75487-609">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-610">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75487-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="75487-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-613">- BindingEvents</span></span><br><span data-ttu-id="75487-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-614">
         - CompressedFile</span></span><br><span data-ttu-id="75487-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-616">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-617">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-617">
         - File</span></span><br><span data-ttu-id="75487-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-619">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-622">
         - PdfFile</span></span><br><span data-ttu-id="75487-623">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-623">
         - Selection</span></span><br><span data-ttu-id="75487-624">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-624">
         - Settings</span></span><br><span data-ttu-id="75487-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-625">
         - TableBindings</span></span><br><span data-ttu-id="75487-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-626">
         - TableCoercion</span></span><br><span data-ttu-id="75487-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-627">
         - TextBindings</span></span><br><span data-ttu-id="75487-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-628">
         - TextCoercion</span></span><br><span data-ttu-id="75487-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-630">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="75487-630">Office on iPad</span></span><br><span data-ttu-id="75487-631">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-632">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="75487-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-638">- BindingEvents</span></span><br><span data-ttu-id="75487-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-639">
         - CompressedFile</span></span><br><span data-ttu-id="75487-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-641">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-642">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-642">
         - File</span></span><br><span data-ttu-id="75487-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-644">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-647">
         - PdfFile</span></span><br><span data-ttu-id="75487-648">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-648">
         - Selection</span></span><br><span data-ttu-id="75487-649">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-649">
         - Settings</span></span><br><span data-ttu-id="75487-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-650">
         - TableBindings</span></span><br><span data-ttu-id="75487-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-651">
         - TableCoercion</span></span><br><span data-ttu-id="75487-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-652">
         - TextBindings</span></span><br><span data-ttu-id="75487-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-653">
         - TextCoercion</span></span><br><span data-ttu-id="75487-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-655">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-655">Office on Mac</span></span><br><span data-ttu-id="75487-656">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-657">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-657">- TaskPane</span></span><br><span data-ttu-id="75487-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="75487-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-665">- BindingEvents</span></span><br><span data-ttu-id="75487-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-666">
         - CompressedFile</span></span><br><span data-ttu-id="75487-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-668">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-669">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-669">
         - File</span></span><br><span data-ttu-id="75487-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-671">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-674">
         - PdfFile</span></span><br><span data-ttu-id="75487-675">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-675">
         - Selection</span></span><br><span data-ttu-id="75487-676">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-676">
         - Settings</span></span><br><span data-ttu-id="75487-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-677">
         - TableBindings</span></span><br><span data-ttu-id="75487-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-678">
         - TableCoercion</span></span><br><span data-ttu-id="75487-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-679">
         - TextBindings</span></span><br><span data-ttu-id="75487-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-680">
         - TextCoercion</span></span><br><span data-ttu-id="75487-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-682">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-682">Office 2019 on Mac</span></span><br><span data-ttu-id="75487-683">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-684">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-684">- TaskPane</span></span><br><span data-ttu-id="75487-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="75487-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="75487-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="75487-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="75487-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-691">- BindingEvents</span></span><br><span data-ttu-id="75487-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-692">
         - CompressedFile</span></span><br><span data-ttu-id="75487-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-694">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-695">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-695">
         - File</span></span><br><span data-ttu-id="75487-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-697">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-700">
         - PdfFile</span></span><br><span data-ttu-id="75487-701">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-701">
         - Selection</span></span><br><span data-ttu-id="75487-702">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-702">
         - Settings</span></span><br><span data-ttu-id="75487-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-703">
         - TableBindings</span></span><br><span data-ttu-id="75487-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-704">
         - TableCoercion</span></span><br><span data-ttu-id="75487-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-705">
         - TextBindings</span></span><br><span data-ttu-id="75487-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-706">
         - TextCoercion</span></span><br><span data-ttu-id="75487-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-708">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-708">Office 2016 on Mac</span></span><br><span data-ttu-id="75487-709">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-710">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="75487-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="75487-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="75487-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="75487-714">- BindingEvents</span></span><br><span data-ttu-id="75487-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-715">
         - CompressedFile</span></span><br><span data-ttu-id="75487-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="75487-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="75487-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-717">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-718">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-718">
         - File</span></span><br><span data-ttu-id="75487-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="75487-720">
         - MatrixBindings</span></span><br><span data-ttu-id="75487-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="75487-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="75487-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-723">
         - PdfFile</span></span><br><span data-ttu-id="75487-724">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-724">
         - Selection</span></span><br><span data-ttu-id="75487-725">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-725">
         - Settings</span></span><br><span data-ttu-id="75487-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="75487-726">
         - TableBindings</span></span><br><span data-ttu-id="75487-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-727">
         - TableCoercion</span></span><br><span data-ttu-id="75487-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="75487-728">
         - TextBindings</span></span><br><span data-ttu-id="75487-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-729">
         - TextCoercion</span></span><br><span data-ttu-id="75487-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="75487-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="75487-731">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="75487-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="75487-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="75487-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75487-733">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-733">Platform</span></span></th>
    <th><span data-ttu-id="75487-734">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-734">Extension points</span></span></th>
    <th><span data-ttu-id="75487-735">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="75487-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-737">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="75487-738">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-738">- Content</span></span><br><span data-ttu-id="75487-739">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-739">
         - TaskPane</span></span><br><span data-ttu-id="75487-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="75487-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="75487-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-745">- ActiveView</span></span><br><span data-ttu-id="75487-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-746">
         - CompressedFile</span></span><br><span data-ttu-id="75487-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-747">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-748">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-748">
         - File</span></span><br><span data-ttu-id="75487-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-749">
         - PdfFile</span></span><br><span data-ttu-id="75487-750">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-750">
         - Selection</span></span><br><span data-ttu-id="75487-751">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-751">
         - Settings</span></span><br><span data-ttu-id="75487-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-753">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-753">Office on Windows</span></span><br><span data-ttu-id="75487-754">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-755">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-755">- Content</span></span><br><span data-ttu-id="75487-756">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-756">
         - TaskPane</span></span><br><span data-ttu-id="75487-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="75487-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="75487-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-762">- ActiveView</span></span><br><span data-ttu-id="75487-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-763">
         - CompressedFile</span></span><br><span data-ttu-id="75487-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-764">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-765">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-765">
         - File</span></span><br><span data-ttu-id="75487-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-766">
         - PdfFile</span></span><br><span data-ttu-id="75487-767">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-767">
         - Selection</span></span><br><span data-ttu-id="75487-768">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-768">
         - Settings</span></span><br><span data-ttu-id="75487-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-770">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-770">Office 2019 on Windows</span></span><br><span data-ttu-id="75487-771">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-772">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-772">- Content</span></span><br><span data-ttu-id="75487-773">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-773">
         - TaskPane</span></span><br><span data-ttu-id="75487-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-777">- ActiveView</span></span><br><span data-ttu-id="75487-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-778">
         - CompressedFile</span></span><br><span data-ttu-id="75487-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-779">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-780">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-780">
         - File</span></span><br><span data-ttu-id="75487-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-781">
         - PdfFile</span></span><br><span data-ttu-id="75487-782">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-782">
         - Selection</span></span><br><span data-ttu-id="75487-783">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-783">
         - Settings</span></span><br><span data-ttu-id="75487-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-785">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-785">Office 2016 on Windows</span></span><br><span data-ttu-id="75487-786">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-787">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-787">- Content</span></span><br><span data-ttu-id="75487-788">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75487-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="75487-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-791">- ActiveView</span></span><br><span data-ttu-id="75487-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-792">
         - CompressedFile</span></span><br><span data-ttu-id="75487-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-793">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-794">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-794">
         - File</span></span><br><span data-ttu-id="75487-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-795">
         - PdfFile</span></span><br><span data-ttu-id="75487-796">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-796">
         - Selection</span></span><br><span data-ttu-id="75487-797">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-797">
         - Settings</span></span><br><span data-ttu-id="75487-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-799">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-799">Office 2013 on Windows</span></span><br><span data-ttu-id="75487-800">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-801">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-801">- Content</span></span><br><span data-ttu-id="75487-802">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="75487-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75487-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="75487-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-805">- ActiveView</span></span><br><span data-ttu-id="75487-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-806">
         - CompressedFile</span></span><br><span data-ttu-id="75487-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-807">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-808">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-808">
         - File</span></span><br><span data-ttu-id="75487-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-809">
         - PdfFile</span></span><br><span data-ttu-id="75487-810">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-810">
         - Selection</span></span><br><span data-ttu-id="75487-811">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-811">
         - Settings</span></span><br><span data-ttu-id="75487-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-813">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="75487-813">Office on iPad</span></span><br><span data-ttu-id="75487-814">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-815">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-815">- Content</span></span><br><span data-ttu-id="75487-816">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="75487-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-820">- ActiveView</span></span><br><span data-ttu-id="75487-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-821">
         - CompressedFile</span></span><br><span data-ttu-id="75487-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-822">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-823">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-823">
         - File</span></span><br><span data-ttu-id="75487-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-824">
         - PdfFile</span></span><br><span data-ttu-id="75487-825">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-825">
         - Selection</span></span><br><span data-ttu-id="75487-826">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-826">
         - Settings</span></span><br><span data-ttu-id="75487-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-828">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-828">Office on Mac</span></span><br><span data-ttu-id="75487-829">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="75487-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="75487-830">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-830">- Content</span></span><br><span data-ttu-id="75487-831">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-831">
         - TaskPane</span></span><br><span data-ttu-id="75487-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="75487-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="75487-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="75487-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="75487-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-837">- ActiveView</span></span><br><span data-ttu-id="75487-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-838">
         - CompressedFile</span></span><br><span data-ttu-id="75487-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-839">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-840">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-840">
         - File</span></span><br><span data-ttu-id="75487-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-841">
         - PdfFile</span></span><br><span data-ttu-id="75487-842">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-842">
         - Selection</span></span><br><span data-ttu-id="75487-843">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-843">
         - Settings</span></span><br><span data-ttu-id="75487-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-845">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-845">Office 2019 on Mac</span></span><br><span data-ttu-id="75487-846">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-847">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-847">- Content</span></span><br><span data-ttu-id="75487-848">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-848">
         - TaskPane</span></span><br><span data-ttu-id="75487-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-852">- ActiveView</span></span><br><span data-ttu-id="75487-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-853">
         - CompressedFile</span></span><br><span data-ttu-id="75487-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-854">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-855">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-855">
         - File</span></span><br><span data-ttu-id="75487-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-856">
         - PdfFile</span></span><br><span data-ttu-id="75487-857">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-857">
         - Selection</span></span><br><span data-ttu-id="75487-858">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-858">
         - Settings</span></span><br><span data-ttu-id="75487-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-860">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="75487-860">Office 2016 on Mac</span></span><br><span data-ttu-id="75487-861">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-862">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-862">- Content</span></span><br><span data-ttu-id="75487-863">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="75487-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="75487-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="75487-866">- ActiveView</span></span><br><span data-ttu-id="75487-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="75487-867">
         - CompressedFile</span></span><br><span data-ttu-id="75487-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-868">
         - DocumentEvents</span></span><br><span data-ttu-id="75487-869">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="75487-869">
         - File</span></span><br><span data-ttu-id="75487-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="75487-870">
         - PdfFile</span></span><br><span data-ttu-id="75487-871">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-871">
         - Selection</span></span><br><span data-ttu-id="75487-872">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-872">
         - Settings</span></span><br><span data-ttu-id="75487-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="75487-874">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="75487-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="75487-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="75487-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75487-876">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-876">Platform</span></span></th>
    <th><span data-ttu-id="75487-877">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-877">Extension points</span></span></th>
    <th><span data-ttu-id="75487-878">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="75487-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-880">Office na Web</span><span class="sxs-lookup"><span data-stu-id="75487-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="75487-881">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="75487-881">- Content</span></span><br><span data-ttu-id="75487-882">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-882">
         - TaskPane</span></span><br><span data-ttu-id="75487-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="75487-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="75487-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="75487-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="75487-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="75487-887">- DocumentEvents</span></span><br><span data-ttu-id="75487-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="75487-889">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="75487-889">
         - Settings</span></span><br><span data-ttu-id="75487-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="75487-891">Project</span><span class="sxs-lookup"><span data-stu-id="75487-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="75487-892">Plataforma</span><span class="sxs-lookup"><span data-stu-id="75487-892">Platform</span></span></th>
    <th><span data-ttu-id="75487-893">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="75487-893">Extension points</span></span></th>
    <th><span data-ttu-id="75487-894">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="75487-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="75487-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="75487-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-896">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-896">Office 2019 on Windows</span></span><br><span data-ttu-id="75487-897">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-898">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-900">- Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-900">- Selection</span></span><br><span data-ttu-id="75487-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-902">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-902">Office 2016 on Windows</span></span><br><span data-ttu-id="75487-903">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-904">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-906">- Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-906">- Selection</span></span><br><span data-ttu-id="75487-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="75487-908">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="75487-908">Office 2013 on Windows</span></span><br><span data-ttu-id="75487-909">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="75487-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="75487-910">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="75487-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="75487-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="75487-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="75487-912">- Seleção</span><span class="sxs-lookup"><span data-stu-id="75487-912">- Selection</span></span><br><span data-ttu-id="75487-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="75487-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="75487-914">Confira também</span><span class="sxs-lookup"><span data-stu-id="75487-914">See also</span></span>

- [<span data-ttu-id="75487-915">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="75487-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="75487-916">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="75487-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="75487-917">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="75487-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="75487-918">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="75487-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="75487-919">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="75487-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="75487-920">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="75487-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="75487-921">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="75487-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="75487-922">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="75487-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="75487-923">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="75487-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="75487-924">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="75487-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="75487-925">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="75487-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="75487-926">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="75487-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)