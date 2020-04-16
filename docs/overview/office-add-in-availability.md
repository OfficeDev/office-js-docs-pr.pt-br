---
title: Disponibilidade de host e plataforma para suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 04/13/2020
localization_priority: Priority
ms.openlocfilehash: 72da8db755fe6d1d166f66a70c8c298e5a27adff
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241053"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="8e360-103">Disponibilidade de host e plataforma para suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8e360-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="8e360-104">Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado.</span><span class="sxs-lookup"><span data-stu-id="8e360-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="8e360-105">As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que são compatíveis atualmente com cada aplicativo do Office.</span><span class="sxs-lookup"><span data-stu-id="8e360-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="8e360-106">A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.</span><span class="sxs-lookup"><span data-stu-id="8e360-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="8e360-107">Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also).</span><span class="sxs-lookup"><span data-stu-id="8e360-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="8e360-108">Excel</span><span class="sxs-lookup"><span data-stu-id="8e360-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8e360-109">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8e360-110">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8e360-111">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8e360-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-113">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="8e360-114">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-114">- TaskPane</span></span><br><span data-ttu-id="8e360-115">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-115">
        - Content</span></span><br><span data-ttu-id="8e360-116">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-116">
        - Custom Functions</span></span><br><span data-ttu-id="8e360-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="8e360-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8e360-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8e360-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8e360-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8e360-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8e360-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="8e360-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="8e360-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-130">
        - BindingEvents</span></span><br><span data-ttu-id="8e360-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-131">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-132">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-133">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-133">
        - File</span></span><br><span data-ttu-id="8e360-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-134">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-136">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-136">
        - Selection</span></span><br><span data-ttu-id="8e360-137">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-137">
        - Settings</span></span><br><span data-ttu-id="8e360-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-138">
        - TableBindings</span></span><br><span data-ttu-id="8e360-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-139">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-140">
        - TextBindings</span></span><br><span data-ttu-id="8e360-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-142">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-142">Office on Windows</span></span><br><span data-ttu-id="8e360-143">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-144">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-144">- TaskPane</span></span><br><span data-ttu-id="8e360-145">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-145">
        - Content</span></span><br><span data-ttu-id="8e360-146">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-146">
        - Custom Functions</span></span><br><span data-ttu-id="8e360-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </span><span class="sxs-lookup"><span data-stu-id="8e360-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="8e360-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8e360-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8e360-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8e360-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8e360-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="8e360-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-161">
        - BindingEvents</span></span><br><span data-ttu-id="8e360-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-162">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-163">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-164">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-164">
        - File</span></span><br><span data-ttu-id="8e360-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-165">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-167">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-167">
        - Selection</span></span><br><span data-ttu-id="8e360-168">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-168">
        - Settings</span></span><br><span data-ttu-id="8e360-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-169">
        - TableBindings</span></span><br><span data-ttu-id="8e360-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-170">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-171">
        - TextBindings</span></span><br><span data-ttu-id="8e360-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-173">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-173">Office 2019 on Windows</span></span><br><span data-ttu-id="8e360-174">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8e360-175">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-175">- TaskPane</span></span><br><span data-ttu-id="8e360-176">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-176">
        - Content</span></span><br><span data-ttu-id="8e360-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8e360-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-188">- BindingEvents</span></span><br><span data-ttu-id="8e360-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-189">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-190">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-191">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-191">
        - File</span></span><br><span data-ttu-id="8e360-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-192">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-194">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-194">
        - Selection</span></span><br><span data-ttu-id="8e360-195">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-195">
        - Settings</span></span><br><span data-ttu-id="8e360-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-196">
        - TableBindings</span></span><br><span data-ttu-id="8e360-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-197">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-198">
        - TextBindings</span></span><br><span data-ttu-id="8e360-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-200">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-200">Office 2016 on Windows</span></span><br><span data-ttu-id="8e360-201">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8e360-202">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-202">- TaskPane</span></span><br><span data-ttu-id="8e360-203">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-203">
        - Content</span></span></td>
    <td><span data-ttu-id="8e360-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8e360-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-207">- BindingEvents</span></span><br><span data-ttu-id="8e360-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-208">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-209">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-210">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-210">
        - File</span></span><br><span data-ttu-id="8e360-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-211">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-213">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-213">
        - Selection</span></span><br><span data-ttu-id="8e360-214">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-214">
        - Settings</span></span><br><span data-ttu-id="8e360-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-215">
        - TableBindings</span></span><br><span data-ttu-id="8e360-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-216">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-217">
        - TextBindings</span></span><br><span data-ttu-id="8e360-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-219">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-219">Office 2013 on Windows</span></span><br><span data-ttu-id="8e360-220">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8e360-221">
        - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-221">
        - TaskPane</span></span><br><span data-ttu-id="8e360-222">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="8e360-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e360-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="8e360-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-225">
        - BindingEvents</span></span><br><span data-ttu-id="8e360-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-226">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-227">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-228">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-228">
        - File</span></span><br><span data-ttu-id="8e360-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-229">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-231">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-231">
        - Selection</span></span><br><span data-ttu-id="8e360-232">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-232">
        - Settings</span></span><br><span data-ttu-id="8e360-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-233">
        - TableBindings</span></span><br><span data-ttu-id="8e360-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-234">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-235">
        - TextBindings</span></span><br><span data-ttu-id="8e360-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-237">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="8e360-237">Office on iPad</span></span><br><span data-ttu-id="8e360-238">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="8e360-239">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-239">- TaskPane</span></span><br><span data-ttu-id="8e360-240">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-240">
        - Content</span></span></td>
    <td><span data-ttu-id="8e360-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8e360-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8e360-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8e360-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8e360-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-253">- BindingEvents</span></span><br><span data-ttu-id="8e360-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-254">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-255">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-255">
        - File</span></span><br><span data-ttu-id="8e360-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-256">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-258">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-258">
        - Selection</span></span><br><span data-ttu-id="8e360-259">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-259">
        - Settings</span></span><br><span data-ttu-id="8e360-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-260">
        - TableBindings</span></span><br><span data-ttu-id="8e360-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-261">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-262">
        - TextBindings</span></span><br><span data-ttu-id="8e360-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-264">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-264">Office on Mac</span></span><br><span data-ttu-id="8e360-265">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="8e360-266">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-266">- TaskPane</span></span><br><span data-ttu-id="8e360-267">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-267">
        - Content</span></span><br><span data-ttu-id="8e360-268">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-268">
        - Custom Functions</span></span><br><span data-ttu-id="8e360-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8e360-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="8e360-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="8e360-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="8e360-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="8e360-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="8e360-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-283">- BindingEvents</span></span><br><span data-ttu-id="8e360-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-284">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-285">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-286">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-286">
        - File</span></span><br><span data-ttu-id="8e360-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-287">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-289">
        - PdfFile</span></span><br><span data-ttu-id="8e360-290">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-290">
        - Selection</span></span><br><span data-ttu-id="8e360-291">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-291">
        - Settings</span></span><br><span data-ttu-id="8e360-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-292">
        - TableBindings</span></span><br><span data-ttu-id="8e360-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-293">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-294">
        - TextBindings</span></span><br><span data-ttu-id="8e360-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-296">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-296">Office 2019 on Mac</span></span><br><span data-ttu-id="8e360-297">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8e360-298">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-298">- TaskPane</span></span><br><span data-ttu-id="8e360-299">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-299">
        - Content</span></span><br><span data-ttu-id="8e360-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="8e360-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="8e360-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="8e360-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="8e360-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="8e360-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="8e360-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="8e360-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="8e360-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-311">- BindingEvents</span></span><br><span data-ttu-id="8e360-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-312">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-313">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-314">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-314">
        - File</span></span><br><span data-ttu-id="8e360-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-315">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-317">
        - PdfFile</span></span><br><span data-ttu-id="8e360-318">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-318">
        - Selection</span></span><br><span data-ttu-id="8e360-319">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-319">
        - Settings</span></span><br><span data-ttu-id="8e360-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-320">
        - TableBindings</span></span><br><span data-ttu-id="8e360-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-321">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-322">
        - TextBindings</span></span><br><span data-ttu-id="8e360-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-324">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-324">Office 2016 on Mac</span></span><br><span data-ttu-id="8e360-325">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="8e360-326">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-326">- TaskPane</span></span><br><span data-ttu-id="8e360-327">
        - Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-327">
        - Content</span></span></td>
    <td><span data-ttu-id="8e360-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="8e360-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8e360-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="8e360-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-331">- BindingEvents</span></span><br><span data-ttu-id="8e360-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-332">
        - CompressedFile</span></span><br><span data-ttu-id="8e360-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-333">
        - DocumentEvents</span></span><br><span data-ttu-id="8e360-334">
        - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-334">
        - File</span></span><br><span data-ttu-id="8e360-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-335">
        - MatrixBindings</span></span><br><span data-ttu-id="8e360-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="8e360-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-337">
        - PdfFile</span></span><br><span data-ttu-id="8e360-338">
        - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-338">
        - Selection</span></span><br><span data-ttu-id="8e360-339">
        - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-339">
        - Settings</span></span><br><span data-ttu-id="8e360-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-340">
        - TableBindings</span></span><br><span data-ttu-id="8e360-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-341">
        - TableCoercion</span></span><br><span data-ttu-id="8e360-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-342">
        - TextBindings</span></span><br><span data-ttu-id="8e360-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="8e360-344">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="8e360-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="8e360-345">Funções personalizadas (somente Excel)</span><span class="sxs-lookup"><span data-stu-id="8e360-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="8e360-346">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="8e360-347">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="8e360-348">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="8e360-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-350">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-350">Office on the web</span></span></td>
    <td><span data-ttu-id="8e360-351">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="8e360-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-353">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-353">Office on Windows</span></span><br><span data-ttu-id="8e360-354">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="8e360-355">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="8e360-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-357">Office para Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-357">Office for Mac</span></span><br><span data-ttu-id="8e360-358">(conectado ao Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="8e360-359">
        - Funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="8e360-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="8e360-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="8e360-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="8e360-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e360-362">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-362">Platform</span></span></th>
    <th><span data-ttu-id="8e360-363">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-363">Extension points</span></span></th>
    <th><span data-ttu-id="8e360-364">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e360-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-366">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-366">Office on the web</span></span><br><span data-ttu-id="8e360-367">(moderno)</span><span class="sxs-lookup"><span data-stu-id="8e360-367">(modern)</span></span></td>
    <td> <span data-ttu-id="8e360-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e360-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="8e360-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="8e360-381">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-382">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-382">Office on the web</span></span><br><span data-ttu-id="8e360-383">(clássico)</span><span class="sxs-lookup"><span data-stu-id="8e360-383">(classic)</span></span></td>
    <td> <span data-ttu-id="8e360-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8e360-395">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-396">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-396">Office on Windows</span></span><br><span data-ttu-id="8e360-397">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8e360-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="8e360-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="8e360-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e360-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="8e360-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="8e360-412">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-413">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-413">Office 2019 on Windows</span></span><br><span data-ttu-id="8e360-414">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8e360-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="8e360-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="8e360-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e360-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="8e360-428">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-429">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-429">Office 2016 on Windows</span></span><br><span data-ttu-id="8e360-430">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="8e360-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Módulos</a></span><span class="sxs-lookup"><span data-stu-id="8e360-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="8e360-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="8e360-441">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-442">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-442">Office 2013 on Windows</span></span><br><span data-ttu-id="8e360-443">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="8e360-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Caixa de correio 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="8e360-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Caixa de correio 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="8e360-452">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-453">Office no iOS</span><span class="sxs-lookup"><span data-stu-id="8e360-453">Office on iOS</span></span><br><span data-ttu-id="8e360-454">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8e360-462">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-463">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-463">Office on Mac</span></span><br><span data-ttu-id="8e360-464">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="8e360-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Caixa de correio 1.7</a></span><span class="sxs-lookup"><span data-stu-id="8e360-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="8e360-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Caixa de correio 1.8</a></span><span class="sxs-lookup"><span data-stu-id="8e360-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="8e360-478">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-479">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-479">Office 2019 on Mac</span></span><br><span data-ttu-id="8e360-480">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8e360-492">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-493">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-493">Office 2016 on Mac</span></span><br><span data-ttu-id="8e360-494">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Composição da mensagem</a></span><span class="sxs-lookup"><span data-stu-id="8e360-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="8e360-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="8e360-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a></span><span class="sxs-lookup"><span data-stu-id="8e360-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="8e360-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="8e360-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Caixa de correio 1.6</a></span><span class="sxs-lookup"><span data-stu-id="8e360-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="8e360-506">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-507">Outlook no Android</span><span class="sxs-lookup"><span data-stu-id="8e360-507">Office on Android</span></span><br><span data-ttu-id="8e360-508">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Mensagem lida</a></span><span class="sxs-lookup"><span data-stu-id="8e360-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="8e360-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Organizador de compromissos (Redigir): reunião on-line (visualização)</span><span class="sxs-lookup"><span data-stu-id="8e360-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="8e360-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"> Caixa de correio 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="8e360-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Caixa de correio 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="8e360-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"> Caixa de correio 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="8e360-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"> Caixa de correio 1.4</a></span><span class="sxs-lookup"><span data-stu-id="8e360-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="8e360-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"> Caixa de correio 1.5</a></span><span class="sxs-lookup"><span data-stu-id="8e360-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="8e360-517">Não disponível</span><span class="sxs-lookup"><span data-stu-id="8e360-517">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="8e360-518">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="8e360-518">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8e360-519">O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e360-519">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="8e360-520">Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e360-520">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="8e360-521">Word</span><span class="sxs-lookup"><span data-stu-id="8e360-521">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e360-522">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-522">Platform</span></span></th>
    <th><span data-ttu-id="8e360-523">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-523">Extension points</span></span></th>
    <th><span data-ttu-id="8e360-524">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-524">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e360-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-526">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-526">Office on the web</span></span></td>
    <td> <span data-ttu-id="8e360-527">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-527">- TaskPane</span></span><br><span data-ttu-id="8e360-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="8e360-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-535">- BindingEvents</span></span><br><span data-ttu-id="8e360-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-537">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-538">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-538">
         - File</span></span><br><span data-ttu-id="8e360-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-540">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-543">
         - PdfFile</span></span><br><span data-ttu-id="8e360-544">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-544">
         - Selection</span></span><br><span data-ttu-id="8e360-545">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-545">
         - Settings</span></span><br><span data-ttu-id="8e360-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-546">
         - TableBindings</span></span><br><span data-ttu-id="8e360-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-547">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-548">
         - TextBindings</span></span><br><span data-ttu-id="8e360-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-549">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-550">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-551">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-551">Office on Windows</span></span><br><span data-ttu-id="8e360-552">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-552">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-553">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-553">- TaskPane</span></span><br><span data-ttu-id="8e360-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="8e360-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-561">- BindingEvents</span></span><br><span data-ttu-id="8e360-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-562">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-564">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-565">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-565">
         - File</span></span><br><span data-ttu-id="8e360-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-567">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-570">
         - PdfFile</span></span><br><span data-ttu-id="8e360-571">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-571">
         - Selection</span></span><br><span data-ttu-id="8e360-572">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-572">
         - Settings</span></span><br><span data-ttu-id="8e360-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-573">
         - TableBindings</span></span><br><span data-ttu-id="8e360-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-574">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-575">
         - TextBindings</span></span><br><span data-ttu-id="8e360-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-576">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-578">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-578">Office 2019 on Windows</span></span><br><span data-ttu-id="8e360-579">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-580">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-580">- TaskPane</span></span><br><span data-ttu-id="8e360-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-587">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-587">- BindingEvents</span></span><br><span data-ttu-id="8e360-588">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-588">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-589">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-589">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-590">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-590">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-591">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-591">
         - File</span></span><br><span data-ttu-id="8e360-592">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-592">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-593">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-593">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-594">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-594">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-595">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-595">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-596">
         - PdfFile</span></span><br><span data-ttu-id="8e360-597">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-597">
         - Selection</span></span><br><span data-ttu-id="8e360-598">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-598">
         - Settings</span></span><br><span data-ttu-id="8e360-599">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-599">
         - TableBindings</span></span><br><span data-ttu-id="8e360-600">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-600">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-601">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-601">
         - TextBindings</span></span><br><span data-ttu-id="8e360-602">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-602">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-603">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-603">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-604">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-604">Office 2016 on Windows</span></span><br><span data-ttu-id="8e360-605">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-605">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-606">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-606">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8e360-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-610">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-610">- BindingEvents</span></span><br><span data-ttu-id="8e360-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-611">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-612">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-612">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-613">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-613">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-614">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-614">
         - File</span></span><br><span data-ttu-id="8e360-615">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-615">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-616">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-616">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-617">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-617">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-618">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-618">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-619">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-619">
         - PdfFile</span></span><br><span data-ttu-id="8e360-620">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-620">
         - Selection</span></span><br><span data-ttu-id="8e360-621">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-621">
         - Settings</span></span><br><span data-ttu-id="8e360-622">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-622">
         - TableBindings</span></span><br><span data-ttu-id="8e360-623">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-623">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-624">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-624">
         - TextBindings</span></span><br><span data-ttu-id="8e360-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-625">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-626">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-626">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-627">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-627">Office 2013 on Windows</span></span><br><span data-ttu-id="8e360-628">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-628">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-629">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-629">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e360-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="8e360-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-632">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-632">- BindingEvents</span></span><br><span data-ttu-id="8e360-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-633">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-634">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-634">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-635">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-636">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-636">
         - File</span></span><br><span data-ttu-id="8e360-637">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-637">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-638">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-638">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-639">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-639">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-640">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-640">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-641">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-641">
         - PdfFile</span></span><br><span data-ttu-id="8e360-642">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-642">
         - Selection</span></span><br><span data-ttu-id="8e360-643">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-643">
         - Settings</span></span><br><span data-ttu-id="8e360-644">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-644">
         - TableBindings</span></span><br><span data-ttu-id="8e360-645">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-645">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-646">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-646">
         - TextBindings</span></span><br><span data-ttu-id="8e360-647">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-647">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-648">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-648">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-649">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="8e360-649">Office on iPad</span></span><br><span data-ttu-id="8e360-650">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-650">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-651">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="8e360-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-657">- BindingEvents</span></span><br><span data-ttu-id="8e360-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-658">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-660">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-661">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-661">
         - File</span></span><br><span data-ttu-id="8e360-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-663">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-666">
         - PdfFile</span></span><br><span data-ttu-id="8e360-667">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-667">
         - Selection</span></span><br><span data-ttu-id="8e360-668">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-668">
         - Settings</span></span><br><span data-ttu-id="8e360-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-669">
         - TableBindings</span></span><br><span data-ttu-id="8e360-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-670">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-671">
         - TextBindings</span></span><br><span data-ttu-id="8e360-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-672">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-674">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-674">Office on Mac</span></span><br><span data-ttu-id="8e360-675">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-675">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-676">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-676">- TaskPane</span></span><br><span data-ttu-id="8e360-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="8e360-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-684">- BindingEvents</span></span><br><span data-ttu-id="8e360-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-685">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-687">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-688">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-688">
         - File</span></span><br><span data-ttu-id="8e360-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-690">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-693">
         - PdfFile</span></span><br><span data-ttu-id="8e360-694">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-694">
         - Selection</span></span><br><span data-ttu-id="8e360-695">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-695">
         - Settings</span></span><br><span data-ttu-id="8e360-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-696">
         - TableBindings</span></span><br><span data-ttu-id="8e360-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-697">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-698">
         - TextBindings</span></span><br><span data-ttu-id="8e360-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-699">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-701">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-701">Office 2019 on Mac</span></span><br><span data-ttu-id="8e360-702">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-703">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-703">- TaskPane</span></span><br><span data-ttu-id="8e360-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="8e360-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="8e360-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="8e360-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="8e360-710">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-710">- BindingEvents</span></span><br><span data-ttu-id="8e360-711">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-711">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-712">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-712">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-713">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-713">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-714">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-714">
         - File</span></span><br><span data-ttu-id="8e360-715">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-715">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-716">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-716">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-717">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-717">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-718">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-718">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-719">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-719">
         - PdfFile</span></span><br><span data-ttu-id="8e360-720">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-720">
         - Selection</span></span><br><span data-ttu-id="8e360-721">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-721">
         - Settings</span></span><br><span data-ttu-id="8e360-722">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-722">
         - TableBindings</span></span><br><span data-ttu-id="8e360-723">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-723">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-724">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-724">
         - TextBindings</span></span><br><span data-ttu-id="8e360-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-725">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-726">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-726">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-727">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-727">Office 2016 on Mac</span></span><br><span data-ttu-id="8e360-728">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-728">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-729">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-729">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="8e360-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="8e360-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="8e360-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-733">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-733">- BindingEvents</span></span><br><span data-ttu-id="8e360-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-734">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-735">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="8e360-735">
         - CustomXmlParts</span></span><br><span data-ttu-id="8e360-736">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-736">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-737">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-737">
         - File</span></span><br><span data-ttu-id="8e360-738">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-738">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-739">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-739">
         - MatrixBindings</span></span><br><span data-ttu-id="8e360-740">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-740">
         - MatrixCoercion</span></span><br><span data-ttu-id="8e360-741">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-741">
         - OoxmlCoercion</span></span><br><span data-ttu-id="8e360-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-742">
         - PdfFile</span></span><br><span data-ttu-id="8e360-743">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-743">
         - Selection</span></span><br><span data-ttu-id="8e360-744">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-744">
         - Settings</span></span><br><span data-ttu-id="8e360-745">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-745">
         - TableBindings</span></span><br><span data-ttu-id="8e360-746">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-746">
         - TableCoercion</span></span><br><span data-ttu-id="8e360-747">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="8e360-747">
         - TextBindings</span></span><br><span data-ttu-id="8e360-748">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-748">
         - TextCoercion</span></span><br><span data-ttu-id="8e360-749">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="8e360-749">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="8e360-750">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="8e360-750">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="8e360-751">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="8e360-751">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e360-752">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-752">Platform</span></span></th>
    <th><span data-ttu-id="8e360-753">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-753">Extension points</span></span></th>
    <th><span data-ttu-id="8e360-754">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-754">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e360-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-756">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-756">Office on the web</span></span></td>
    <td> <span data-ttu-id="8e360-757">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-757">- Content</span></span><br><span data-ttu-id="8e360-758">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-758">
         - TaskPane</span></span><br><span data-ttu-id="8e360-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8e360-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="8e360-764">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-764">- ActiveView</span></span><br><span data-ttu-id="8e360-765">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-765">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-766">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-766">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-767">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-767">
         - File</span></span><br><span data-ttu-id="8e360-768">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-768">
         - PdfFile</span></span><br><span data-ttu-id="8e360-769">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-769">
         - Selection</span></span><br><span data-ttu-id="8e360-770">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-770">
         - Settings</span></span><br><span data-ttu-id="8e360-771">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-771">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-772">Office no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-772">Office on Windows</span></span><br><span data-ttu-id="8e360-773">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-773">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-774">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-774">- Content</span></span><br><span data-ttu-id="8e360-775">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-775">
         - TaskPane</span></span><br><span data-ttu-id="8e360-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8e360-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="8e360-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-781">- ActiveView</span></span><br><span data-ttu-id="8e360-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-782">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-783">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-784">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-784">
         - File</span></span><br><span data-ttu-id="8e360-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-785">
         - PdfFile</span></span><br><span data-ttu-id="8e360-786">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-786">
         - Selection</span></span><br><span data-ttu-id="8e360-787">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-787">
         - Settings</span></span><br><span data-ttu-id="8e360-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-789">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-789">Office 2019 on Windows</span></span><br><span data-ttu-id="8e360-790">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-791">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-791">- Content</span></span><br><span data-ttu-id="8e360-792">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-792">
         - TaskPane</span></span><br><span data-ttu-id="8e360-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-796">- ActiveView</span></span><br><span data-ttu-id="8e360-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-797">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-798">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-799">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-799">
         - File</span></span><br><span data-ttu-id="8e360-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-800">
         - PdfFile</span></span><br><span data-ttu-id="8e360-801">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-801">
         - Selection</span></span><br><span data-ttu-id="8e360-802">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-802">
         - Settings</span></span><br><span data-ttu-id="8e360-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-804">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-804">Office 2016 on Windows</span></span><br><span data-ttu-id="8e360-805">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-805">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-806">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-806">- Content</span></span><br><span data-ttu-id="8e360-807">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e360-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="8e360-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-810">- ActiveView</span></span><br><span data-ttu-id="8e360-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-811">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-812">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-813">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-813">
         - File</span></span><br><span data-ttu-id="8e360-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-814">
         - PdfFile</span></span><br><span data-ttu-id="8e360-815">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-815">
         - Selection</span></span><br><span data-ttu-id="8e360-816">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-816">
         - Settings</span></span><br><span data-ttu-id="8e360-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-818">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-818">Office 2013 on Windows</span></span><br><span data-ttu-id="8e360-819">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-819">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-820">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-820">- Content</span></span><br><span data-ttu-id="8e360-821">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-821">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="8e360-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e360-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="8e360-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-824">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-824">- ActiveView</span></span><br><span data-ttu-id="8e360-825">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-825">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-826">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-826">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-827">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-827">
         - File</span></span><br><span data-ttu-id="8e360-828">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-828">
         - PdfFile</span></span><br><span data-ttu-id="8e360-829">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-829">
         - Selection</span></span><br><span data-ttu-id="8e360-830">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-830">
         - Settings</span></span><br><span data-ttu-id="8e360-831">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-831">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-832">Office no iPad</span><span class="sxs-lookup"><span data-stu-id="8e360-832">Office on iPad</span></span><br><span data-ttu-id="8e360-833">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-833">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-834">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-834">- Content</span></span><br><span data-ttu-id="8e360-835">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-835">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8e360-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-839">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-839">- ActiveView</span></span><br><span data-ttu-id="8e360-840">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-840">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-841">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-841">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-842">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-842">
         - File</span></span><br><span data-ttu-id="8e360-843">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-843">
         - PdfFile</span></span><br><span data-ttu-id="8e360-844">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-844">
         - Selection</span></span><br><span data-ttu-id="8e360-845">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-845">
         - Settings</span></span><br><span data-ttu-id="8e360-846">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-846">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-847">Office no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-847">Office on Mac</span></span><br><span data-ttu-id="8e360-848">(conectado à assinatura do Office 365)</span><span class="sxs-lookup"><span data-stu-id="8e360-848">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="8e360-849">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-849">- Content</span></span><br><span data-ttu-id="8e360-850">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-850">
         - TaskPane</span></span><br><span data-ttu-id="8e360-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="8e360-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="8e360-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="8e360-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="8e360-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-856">- ActiveView</span></span><br><span data-ttu-id="8e360-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-857">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-858">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-859">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-859">
         - File</span></span><br><span data-ttu-id="8e360-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-860">
         - PdfFile</span></span><br><span data-ttu-id="8e360-861">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-861">
         - Selection</span></span><br><span data-ttu-id="8e360-862">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-862">
         - Settings</span></span><br><span data-ttu-id="8e360-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-863">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-864">Office 2019 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-864">Office 2019 on Mac</span></span><br><span data-ttu-id="8e360-865">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-866">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-866">- Content</span></span><br><span data-ttu-id="8e360-867">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-867">
         - TaskPane</span></span><br><span data-ttu-id="8e360-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-871">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-871">- ActiveView</span></span><br><span data-ttu-id="8e360-872">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-872">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-873">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-873">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-874">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-874">
         - File</span></span><br><span data-ttu-id="8e360-875">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-875">
         - PdfFile</span></span><br><span data-ttu-id="8e360-876">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-876">
         - Selection</span></span><br><span data-ttu-id="8e360-877">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-877">
         - Settings</span></span><br><span data-ttu-id="8e360-878">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-878">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-879">Office 2016 no Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-879">Office 2016 on Mac</span></span><br><span data-ttu-id="8e360-880">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-880">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-881">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-881">- Content</span></span><br><span data-ttu-id="8e360-882">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-882">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="8e360-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="8e360-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-885">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="8e360-885">- ActiveView</span></span><br><span data-ttu-id="8e360-886">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="8e360-886">
         - CompressedFile</span></span><br><span data-ttu-id="8e360-887">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-887">
         - DocumentEvents</span></span><br><span data-ttu-id="8e360-888">
         - Arquivo</span><span class="sxs-lookup"><span data-stu-id="8e360-888">
         - File</span></span><br><span data-ttu-id="8e360-889">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="8e360-889">
         - PdfFile</span></span><br><span data-ttu-id="8e360-890">
         - Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-890">
         - Selection</span></span><br><span data-ttu-id="8e360-891">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-891">
         - Settings</span></span><br><span data-ttu-id="8e360-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-892">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="8e360-893">*&ast; – Adicionado com atualizações pós-lançamento.*</span><span class="sxs-lookup"><span data-stu-id="8e360-893">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="8e360-894">OneNote</span><span class="sxs-lookup"><span data-stu-id="8e360-894">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e360-895">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-895">Platform</span></span></th>
    <th><span data-ttu-id="8e360-896">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-896">Extension points</span></span></th>
    <th><span data-ttu-id="8e360-897">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-897">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e360-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-899">Office na Web</span><span class="sxs-lookup"><span data-stu-id="8e360-899">Office on the web</span></span></td>
    <td> <span data-ttu-id="8e360-900">- Conteúdo</span><span class="sxs-lookup"><span data-stu-id="8e360-900">- Content</span></span><br><span data-ttu-id="8e360-901">
         - TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-901">
         - TaskPane</span></span><br><span data-ttu-id="8e360-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></span><span class="sxs-lookup"><span data-stu-id="8e360-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="8e360-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="8e360-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="8e360-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-906">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="8e360-906">- DocumentEvents</span></span><br><span data-ttu-id="8e360-907">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-907">
         - HtmlCoercion</span></span><br><span data-ttu-id="8e360-908">
         - Configurações</span><span class="sxs-lookup"><span data-stu-id="8e360-908">
         - Settings</span></span><br><span data-ttu-id="8e360-909">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-909">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="8e360-910">Project</span><span class="sxs-lookup"><span data-stu-id="8e360-910">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="8e360-911">Plataforma</span><span class="sxs-lookup"><span data-stu-id="8e360-911">Platform</span></span></th>
    <th><span data-ttu-id="8e360-912">Pontos de extensão</span><span class="sxs-lookup"><span data-stu-id="8e360-912">Extension points</span></span></th>
    <th><span data-ttu-id="8e360-913">Conjuntos de requisitos da API</span><span class="sxs-lookup"><span data-stu-id="8e360-913">API requirement sets</span></span></th>
    <th><span data-ttu-id="8e360-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></span><span class="sxs-lookup"><span data-stu-id="8e360-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-915">Office 2019 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-915">Office 2019 on Windows</span></span><br><span data-ttu-id="8e360-916">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-916">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-917">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-917">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-919">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-919">- Selection</span></span><br><span data-ttu-id="8e360-920">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-920">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-921">Office 2016 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-921">Office 2016 on Windows</span></span><br><span data-ttu-id="8e360-922">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-922">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-923">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-923">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-925">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-925">- Selection</span></span><br><span data-ttu-id="8e360-926">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-926">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="8e360-927">Office 2013 no Windows</span><span class="sxs-lookup"><span data-stu-id="8e360-927">Office 2013 on Windows</span></span><br><span data-ttu-id="8e360-928">(compra avulsa)</span><span class="sxs-lookup"><span data-stu-id="8e360-928">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="8e360-929">- TaskPane</span><span class="sxs-lookup"><span data-stu-id="8e360-929">- TaskPane</span></span></td>
    <td> <span data-ttu-id="8e360-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="8e360-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="8e360-931">- Seleção</span><span class="sxs-lookup"><span data-stu-id="8e360-931">- Selection</span></span><br><span data-ttu-id="8e360-932">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="8e360-932">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="8e360-933">Confira também</span><span class="sxs-lookup"><span data-stu-id="8e360-933">See also</span></span>

- [<span data-ttu-id="8e360-934">Visão geral da plataforma Suplementos do Office</span><span class="sxs-lookup"><span data-stu-id="8e360-934">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="8e360-935">Versões do Office e conjuntos de requisitos</span><span class="sxs-lookup"><span data-stu-id="8e360-935">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="8e360-936">Conjuntos de requisitos comuns da API</span><span class="sxs-lookup"><span data-stu-id="8e360-936">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="8e360-937">Conjuntos de requisitos dos comandos de suplemento</span><span class="sxs-lookup"><span data-stu-id="8e360-937">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="8e360-938">Documentação de Referência da API</span><span class="sxs-lookup"><span data-stu-id="8e360-938">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="8e360-939">Histórico de atualizações do Office 365 ProPlus</span><span class="sxs-lookup"><span data-stu-id="8e360-939">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="8e360-940">Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)</span><span class="sxs-lookup"><span data-stu-id="8e360-940">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="8e360-941">Histórico de atualizações do Office 2013 (clique para executar)</span><span class="sxs-lookup"><span data-stu-id="8e360-941">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="8e360-942">Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="8e360-942">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="8e360-943">Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)</span><span class="sxs-lookup"><span data-stu-id="8e360-943">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="8e360-944">Histórico de atualizações do Office para Mac</span><span class="sxs-lookup"><span data-stu-id="8e360-944">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="8e360-945">Criando Suplementos do Office </span><span class="sxs-lookup"><span data-stu-id="8e360-945">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)